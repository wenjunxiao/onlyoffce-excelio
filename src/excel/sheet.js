import { merge, assign } from 'lodash';
import {
  hex2rgb,
  checkPrecision,
  encodeCell
} from './utils';

const LineStyle = {
  none: 'None',
  double: 'Double',
  hair: 'Hair',
  dashdotdot: 'DashDotDot',
  dashdot: 'DashDot',
  dotted: 'Dotted',
  dashed: 'Dashed',
  thin: 'Thin',
  mediumdashdotdot: 'MediumDashDotDot',
  slantdashdot: 'SlantDashDot',
  mediumdashdot: 'MediumDashDot',
  mediumdashed: 'MediumDashed',
  medium: 'Medium',
  thick: 'Thick'
};

const getLineStyle = style => LineStyle[style.toLowerCase()] || style;
/**
 * UTC时间戳转换成Excel中是时间戳
 * @function utc2excel
 * @memberof ExcelIO.ExcelWriter.Sheet
 * @static
 * @param {Number} v UTC时间戳
 * @returns {Number} Excel中的时间戳
 */
const utc2excel = v => (v + 2209161600000) / (86400 * 1000);
/**
 * 把服务端的宽度(像素px)转换成onlyoffice中的宽度(字符数char)。
 * 根据服务端的{@link https://github.com/wenjunxiao/node-excelio/blob/main/lib/writer.js excelio}
 * 及其用到库{@link https://github.com/protobi/js-xlsx/blob/v0.8.6/xlsx.js xlsx-style}中的<code>px2char</code>
 * 方法，以及onlyoffice中<code>cell/apiBuilder.js</code>的说明，只需要使用<code>px2char</code>公式转换即可。
 * 但是在实际测试过程中，发现并不一致，需要对修正计算公式，把公式中的(px - 5)变成(px - 4)。<br/>
 * 2.在Excel文档中`xl/worksheets/sheet*.xml`中的width相同，
 *   由于主题中缺少宋体导致显示的比服务端excelio生成的文件要宽，如果在`xl/theme/theme1.xml`文件中
 *   加入对应宋体之后则和服务器端显示的宽度相同，加入的内容如下<code><pre>
    &lt;a:minorFont&gt;
      &lt;a:latin typeface="Calibri"&gt;&lt;/a:latin&gt;
      &lt;a:ea typeface="Arial"&gt;&lt;/a:ea&gt;
      &lt;a:cs typeface="Arial"&gt;&lt;/a:cs&gt;
      &lt;a:font script="Hans" typeface="宋体"/&gt;
    &lt;/a:minorFont&gt;</pre></code>
 *   但是onlyoffice生成的excel显示的宽度是真实宽度，服务端生成的要比设置的要窄。
 *   不确定是服务端的有问题还是onlyoffice有问题
 * @see https://github.com/protobi/js-xlsx/blob/master/xlsx.js
 * @function px2width
 * @memberof ExcelIO.ExcelWriter.Sheet
 * @static
 * @param {Number} px 宽度的像素值
 * @returns {Number} OnlyOffice中的宽度
 */
const px2width = px => {
  px = px || 0;
  return (((px - 5) / 7 * 100 + 0.5) | 0) / 100;
};
const char2px = (chr, sz) => { return (chr * 8 + Math.ceil(chr / 10) * 5) * Math.ceil(sz / 10); }
const char2width = (chr, sz) => px2width(char2px(chr || 0, sz));

const charsOfStr = (str) => {
  var l = str.length;
  var len = 0;
  for (var i = 0; i < l; i++) {
    if ((str.charCodeAt(i) & 0xff00) !== 0) {
      len++;
    }
    len++;
  }
  return len;
}

export {
  px2width,
  utc2excel
};

/**
 * Sheet页，只能被<code>ExcelWriter</code>创建
 * @class Sheet
 * @memberof ExcelIO.ExcelWriter
 * @param {Object} options 选项
 * @param {String} options.sheetName 当前Sheet页名称，<code>ExcelWriter</code>传入
 * @param {ExcelIO.ExcelWriter} options.owner 当前Sheet的所属Writer
 * @param {String} [options.NaN] 数字字段为<code>null</code>或<code>undefined</code>的默认值
 * @param {Boolean} [options.showGridLines] 是否显示表格线，默认<code>true</code>
 * @param {Object} [options.titleOpts] 表格标题默认选项
 * @param {Object} [options.cellOpts] 单元格默认选项
 * @param {Object} [options.alignment] 默认对齐方式<code>{horizontal: 'left',vertical: 'left'}</code>
 * @param {number} [options.width] 默认宽度
 * @param {number} [options.minWidth] 最小宽度
 * @param {number} [options.titleLine] 标题行行号，默认没有标题行
 * @param {Boolean|String} [options.border2end] 是否显示单元格边框，可以设置为单元格颜色的十六进制
 */
export default class Sheet {

  constructor(options) {
    this.rowIdx = -1;
    this.colIdx = -1;
    this.maxCol = 0;
    this._opts = assign({}, options || {});
    this.sheetName = this._opts.sheetName || 'Sheet1';
    this.owner = this._opts.owner || {};
    this.NaN = this._opts.NaN || '';
    this._px = this._opts.px === true;
    this._width = this._px ? px2width : char2width;
    this.funcs = {};
    this._colWidths = {};
    this.chunks = [
      'var cell',
    ];
  }

  /**
   * 重命名Sheet页的名称
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {String} name 新的名称
   * @returns {this}
   */
  rename (name) {
    if (name !== this.sheetName) {
      let cur = this.sheetName;
      this.sheetName = name;
      this.chunks.push(`if(sheet.GetSheet(${JSON.stringify(name)})) throw new Error('Sheet with name [${name}] already exists')`);
      this.chunks.push(`sheet.SetName(${JSON.stringify(name)})`);
      this.owner.rename(name, cur);
    }
    return this;
  }

  /**
   * 清空当前Sheet页
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @returns {this}
   */
  clear () {
    this.rowIdx = -1;
    this.colIdx = -1;
    this.maxCol = 0;
    this.chunks.push(`sheet.GetUsedRange().Clear();`);
    return this;
  }

  /**
   * 当前操作的行号
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @returns {this}
   */
  rowIndex () {
    return this.rowIdx;
  }

  /**
   * 当前操作的列号
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @returns {this}
   */
  colIndex () {
    return this.colIdx;
  }

  /**
   * 跳过指定行数
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} [rows=1] 跳过的行数
   * @param {number} [cells=0] 同时跳过的单元格，默认在行首
   * @returns {this}
   */
  skipRow (rows = 1, cells = 0) {
    if (this.colIdx > this.maxCol) {
      this.maxCol = this.colIdx;
    }
    this.rowIdx += rows;
    if (cells < 0) {
      this.colIdx += cells;
    } else {
      this.colIdx = cells - 1;
    }
    return this;
  }

  /**
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} [cells=1] 跳过的单元格
   * @returns {this}
   */
  skipCell (cells = 1) {
    this.colIdx += cells;
    return this;
  }

  /**
   * 跳转到指定的行列
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} row 指定行号
   * @param {number} [cell=0] 指定列号，默认行首
   * @returns {this}
   */
  go (row, cell = 0) {
    if (this.colIdx > this.maxCol) {
      this.maxCol = this.colIdx;
    }
    this.rowIdx = row - 1;
    this.colIdx = cell - 1;
    return this;
  }

  /**
   * 定位到下几行(默认下一行)
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} [cells=0] 同时跳转到对应的列，默认行首
   * @returns {this} 
   */
  row (cells = 0) {
    if (this.colIdx > this.maxCol) {
      this.maxCol = this.colIdx;
    }
    ++this.rowIdx;
    this.colIdx = cells - 1;
    return this;
  }

  /**
   * 在当前单元格写入数字，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} v 数字
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  number (v, options) {
    v = this.formatNumber(v);
    if (isNaN(v)) return this.cell(v, options);
    return this.cell(v, options, 'n');
  }

  /**
   * 在当前单元格写入布尔值，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} v 布尔值
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  boolean (v, options) {
    return this.cell(v, options, 'b');
  }

  /**
   * 在当前单元格写入字符串，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} v 字符串
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  string (v, options) {
    return this.cell(v, options, 's');
  }

  /**
   * 在当前单元格写入UTC时间，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} v 数字
   * @param {String} [format] 时间格式化字符串
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  utc (v, format, options) {
    if (typeof format === 'object') {
      options = format
      format = options.format
    }
    if (!format) {
      format = 'YYYY-MM-DD HH:mm:ss'
    }
    if (v instanceof Date) {
      v = utc2excel(v.valueOf());
    }
    return this.cell(v, options, 'd', format);
  }

  /**
   * 在当前单元格写入当前时区的时间，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} v 数字
   * @param {String} [format] 时间格式化字符串
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  date (v, format, options) {
    if (typeof format === 'object') {
      options = format
      format = options.format
    }
    if (!(v instanceof Date)) {
      v = new Date(v)
    }
    if (!format) {
      format = 'YYYY-MM-DD HH:mm:ss'
    }
    v = new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(),
      v.getMinutes(), v.getSeconds(), v.getMilliseconds()));
    return this.cell(utc2excel(v.valueOf()), options, 'd', format);
  }

  formatNumber (v, precision) {
    if (v === null || v === undefined) return this.NaN;
    if (isNaN(precision) || !v.toFixed) return v.toString().replace(/,/g, '');
    return v.toFixed(precision).replace(/,/g, '');
  }

  /**
   * 在当前单元格写入金额，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} v 金额
   * @param {String} currency 币种
   * @param {String} [precision] 精度
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  currency (v, currency, precision, options) {
    [precision, options] = checkPrecision(precision, options);
    v = this.formatNumber(v, precision);
    if (currency) {
      return this.cell(v, options, 'n', currency + '#,##0.00');
    } else if (/^([^\d\-.])+(.*)$/.test(v)) {
      let prefix = RegExp.$1;
      v = RegExp.$2;
      return this.cell(v, options, 'n', prefix + '#,##0.00');
    } else if (isNaN(v)) {
      return this.cell(v, options);
    }
    return this.cell(v, options, 'n', '4');
  }

  /**
   * 在当前单元格写入百分数，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} v 百分数
   * @param {String} [precision] 精度
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  percent (v, precision, options) {
    [precision, options] = checkPrecision(precision, options);
    v = this.formatNumber(v, precision);
    if (isNaN(v)) {
      return this.cell(v, options);
    }
    return this.cell(v, options, 'n', '0.00%');
  }

  /**
   * 批量写入标题单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {Array} vs 标题数组
   * @param {Object} [options={}] 单元格选项
   * @param {String} [options.newLine] 是否在下一行写入标题，默认<code>true</code>
   * @returns {this}
   */
  titles (vs, options) {
    options = typeof options === 'object' ? options : { width: options };
    if (this.rowIdx === -1 || options.newLine !== false) {
      this.row();
    }
    if (Array.isArray(options)) {
      vs.forEach((v, i) => {
        this.title(v, options[i])
      });
    } else {
      if (this._opts.titleOpts) {
        options = merge({}, this._opts.titleOpts, options);
      }
      for (let v of vs) {
        this.cell(v, options, 's');
      }
    }
    return this;
  }

  /**
   * 写入标题单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {String} v 标题
   * @param {Object|number} [options={}] 单元格选项，可直接指定为宽度
   * @returns {this}
   */
  title (v, options) {
    options = typeof options === 'object' ? options : { width: options };
    if (this._opts.titleOpts) {
      options = merge({}, this._opts.titleOpts, options);
    }
    return this.cell(v, options, 's');
  }

  /**
   * 批量写入单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {Array} vs 数据数组
   * @param {Object} [options={}] 单元格选项
   * @param {String} [options.newLine] 是否在下一行写入数据，默认<code>true</code>
   * @returns {this}
   */
  fillRow (vs, options) {
    options = typeof options === 'object' ? options : { width: options };
    if (this.rowIdx === -1 || options.newLine !== false) {
      this.row();
    }
    if (Array.isArray(options)) {
      vs.forEach((v, i) => {
        this.cell(v, options[i])
      });
    } else {
      for (let v of vs) {
        this.cell(v, options);
      }
    }
    return this;
  }

  /**
   * 批量写入表格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {Array.<Array>} data 表格数据(二维数组)
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  fill (data, options) {
    if (Array.isArray(options)) {
      for (let vs of data) {
        this.row();
        vs.forEach((v, i) => {
          this.cell(v, options[i])
        });
      }
    } else {
      for (let vs of data) {
        this.row();
        for (let v of vs) {
          this.cell(v, options);
        }
      }
    }
    return this;
  }

  /**
   * 在当前单元格写入值，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {*} v 值
   * @param {Object} [options={}] 单元格选项
   * @param {String} [type] 值类型，默认根据值的类型
   * @param {String} [format] 值的格式化
   * @returns {this}
   */
  cell (v, options, type, format) {
    const c = ++this.colIdx;
    this.chunks.push(`cell = sheet.GetRangeByNumber(${this.rowIdx},${c})`);
    if (format) {
      this.chunks.push(`cell.SetNumberFormat(${JSON.stringify(format)})`);
    }
    let fontSize = this._opts.fontSize || 10;
    if (typeof options === 'object') {
      if (options.alignment) {
        if (options.alignment.horizontal) {
          this.chunks.push(`cell.SetAlignHorizontal(${JSON.stringify(options.alignment.horizontal)})`)
        }
        if (options.alignment.vertical) {
          this.chunks.push(`cell.SetAlignVertical(${JSON.stringify(options.alignment.vertical)})`)
        }
      }
      if (options.font) {
        if (options.font.name) {
          this.chunks.push(`cell.SetFontName(${JSON.stringify(options.font.name)})`)
        }
        if (options.font.sz) {
          fontSize = options.font.sz;
          this.chunks.push(`cell.SetFontSize(${options.font.sz})`)
        }
        if (typeof options.font.bold === 'boolean') {
          this.chunks.push(`cell.SetBold(${options.font.bold})`)
        }
      }
      if (options.width) {
        this._colWidths[c] = options.width;
        this.chunks.push(`cell.SetColumnWidth(${this._width(options.width, fontSize)})`);
      }
      if (options.bgColor) {
        this.chunks.push(`cell.SetFillColor(Api.CreateColorFromRGB(${hex2rgb(options.bgColor)}))`);
      }
      if (options.fgColor) {
        this.chunks.push(`cell.SetFontColor(Api.CreateColorFromRGB(${hex2rgb(options.fgColor)}))`);
      }
    } else if (options > 0) {
      this._colWidths[c] = options;
      this.chunks.push(`cell.SetColumnWidth(${this._width(options, fontSize)})`);
    }
    if (v === undefined || v === null) {
      if (type === 'n') {
        v = '0';
      } else if (type === 'b') {
        v = 'false';
      } else {
        v = '';
      }
    } if (typeof v !== 'string') {
      v = v.toString();
    }
    if (!this._colWidths[c]) {
      let width = this._colWidths[c] = char2width(charsOfStr(v), fontSize);
      this.chunks.push(`cell.SetColumnWidth(${width})`);
    }
    this.chunks.push(`cell.SetValue(${JSON.stringify(v)})`);
    return this;
  }

  /**
   * 指定单元格的宽度
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} width 宽度
   * @param {number} [colIndex] 指定列，默认当前刚写入的单元格
   * @returns {this}
   */
  width (width, colIndex = -1) {
    let fontSize = this._opts.fontSize || 10;
    if (colIndex < 0) {
      this.chunks.push(`cell.SetColumnWidth(${this._width(width, fontSize)})`);
    } else {
      this.chunks.push(`sheet.SetColumnWidth(${colIndex},${this._width(width, fontSize)})`);
    }
    return this;
  }

  chWidth (width, colIndex = -1) {
    if (this._px) {
      return this.width(width, colIndex);
    }
    return this.width(width * 1.8, colIndex);
  }

  /**
   * 指定单元格的颜色
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {String} bgColor 背景颜色，十六进制
   * @param {String} fgColor 字体颜色，十六进制
   * @returns {this}
   */
  color (bgColor, fgColor) {
    if (bgColor) {
      this.chunks.push(`cell.SetFillColor(Api.CreateColorFromRGB(${hex2rgb(bgColor)}))`)
    }
    if (fgColor) {
      this.chunks.push(`cell.SetFontColor(Api.CreateColorFromRGB(${hex2rgb(fgColor)}))`)
    }
    return this;
  }

  /**
   * 指定单元格的字体颜色
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {String} color 背景颜色，十六进制
   * @returns {this}
   */
  fgColor (color) {
    return this.color(null, color);
  }

  /**
   * 指定单元格的背景颜色
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {String} color 背景颜色，十六进制
   * @returns {this}
   */
  bgColor (color) {
    return this.color(color, null);
  }

  /**
   * 添加水印
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @returns {this}
   */
  watermark () {
    return this;
  }

  /**
   * 不需要水印
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @returns {this}
   */
  withoutWatermark () {
    return this;
  }

  /**
   * 指定单元格的边框
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} rs 单元格开始行
   * @param {number} cs 单元格开始列
   * @param {number} re 单元格结束行
   * @param {number} ce 单元格结束列
   * @param {String} [color='#000000'] 背景颜色，十六进制，默认黑色
   * @param {String} [style] 单元格样式，默认<code>thin</code>
   * @param {Object} [options] 选项(<code>outer</code>和<code>inner</code>可同时设置，只设置一个表示只设置内部或外部边框)
   * @param {Boolean|Object} [options.outer] 是否设置外边框，<code>true</code>或者<code>{color:'',style:''}</code>，
   *  <code>color</code>或<code>style</code>可选
   * @param {Boolean|Object} [options.inner] 是否设置内边框，<code>true</code>或者<code>{color:'',style:''}</code>，
   *  <code>color</code>或<code>style</code>可选
   * @returns {this}
   */
  border (rs, cs, re, ce, color = '#000000', style = 'thin', options = {}) {
    color = `Api.CreateColorFromRGB(${hex2rgb(color)})`;
    style = getLineStyle(style);
    let opts = {};
    if (options) {
      if (options.outer) {
        opts.outer = {};
        if (options.outer.color) {
          opts.outer.color = `Api.CreateColorFromRGB(${hex2rgb(options.outer.color)})`;
        }
        if (options.outer.style) {
          opts.outer.style = getLineStyle(options.outer.style);
        }
      }
      if (options.inner) {
        opts.inner = {};
        if (options.inner.color) {
          opts.inner.color = `Api.CreateColorFromRGB(${hex2rgb(options.inner.color)})`;
        }
        if (options.inner.style) {
          opts.inner.style = getLineStyle(options.inner.style);
        }
      }
    }
    if (!this.funcs.border) {
      this.funcs.border = `function fillBorder(sheet, rs, cs, re, ce, color, style, opts){
        if (opts.outer || opts.inner) {
          if (opts.outer) {
            var style1 =  opts.outer.style || style;
            var color1 =  opts.outer.color || color;
            for (let ri = rs; ri <= re; ri++) {
              var _cell = sheet.GetRangeByNumber(ri,cs);
              _cell.SetBorders('Left',style1,color1);
              _cell = sheet.GetRangeByNumber(ri,ce);
              _cell.SetBorders('Right',style1,color1);
            }
            for (let ci = cs; ci <= ce; ci++) {
              var _cell = sheet.GetRangeByNumber(rs,ci);
              _cell.SetBorders('Top',style1,color1);
              _cell = sheet.GetRangeByNumber(re,ci);
              _cell.SetBorders('Bottom',style1,color1);
            }
          }
          if (opts.inner) {
            var style1 =  opts.inner.style || style;
            var color1 =  opts.inner.color || color;
            var _cell = sheet.GetRangeByNumber(rs,cs);
            _cell.SetBorders('Right',style1,color1);
            _cell.SetBorders('Bottom',style1,color1);
            _cell = sheet.GetRangeByNumber(rs,ce);
            _cell.SetBorders('Left',style1,color1);
            _cell.SetBorders('Bottom',style1,color1);
            _cell = sheet.GetRangeByNumber(re,cs);
            _cell.SetBorders('Right',style1,color1);
            _cell.SetBorders('Top',style1,color1);
            _cell = sheet.GetRangeByNumber(re,ce);
            _cell.SetBorders('Left',style1,color1);
            _cell.SetBorders('Top',style1,color1);
            for (let ri = rs + 1; ri < re; ri++) {
              var _cell = sheet.GetRangeByNumber(ri,cs);
              _cell.SetBorders('Top',style1,color1);
              _cell.SetBorders('Right',style1,color1);
              _cell.SetBorders('Bottom',style1,color1);
              _cell = sheet.GetRangeByNumber(ri,ce);
              _cell.SetBorders('Top',style1,color1);
              _cell.SetBorders('Bottom',style1,color1);
              _cell.SetBorders('Left',style1,color1);
            }
            for (let ci = cs + 1; ci < ce; ci++) {
              var _cell = sheet.GetRangeByNumber(rs,ci);
              _cell.SetBorders('Right',style1,color1);
              _cell.SetBorders('Bottom',style1,color1);
              _cell.SetBorders('Left',style1,color1);
              _cell = sheet.GetRangeByNumber(re,ci);
              _cell.SetBorders('Top',style1,color1);
              _cell.SetBorders('Right',style1,color1);
              _cell.SetBorders('Left',style1,color1);
            }
            for (let ri = rs + 1; ri < re; ri++) {
              for (let ci = cs + 1; ci < ce; ci++) {
                var _cell = sheet.GetRangeByNumber(ri,ci);
                _cell.SetBorders('Top',style1,color1);
                _cell.SetBorders('Right',style1,color1);
                _cell.SetBorders('Bottom',style1,color1);
                _cell.SetBorders('Left',style1,color1);
              }
            }
          }
        } else {
          for (let ri = rs; ri <= re; ri++) {
            for (let ci = cs; ci <= ce; ci++) {
              var _cell = sheet.GetRangeByNumber(ri,ci);
              _cell.SetBorders('Top',style,color);
              _cell.SetBorders('Right',style,color);
              _cell.SetBorders('Bottom',style,color);
              _cell.SetBorders('Left',style,color);
            }
          }
        }
      }`.replace(/(?:^|\n)\s*\/\/[^\n]*/g, '').replace(/\n/g, '').replace(/ +/g, ' ');
    }
    this.chunks.push(`fillBorder(sheet,${rs},${cs},${re},${ce},${color},${JSON.stringify(style)},${JSON.stringify(opts)})`);
    return this;
  }

  /**
   * 设置从指定行列到文档结束的边框
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} r 单元格开始行
   * @param {number} c 单元格开始列
   * @param {String} [color='#000000'] 背景颜色，十六进制，默认黑色
   * @param {String} [style] 单元格样式，默认<code>thin</code>
   * @param {Object} [options] 选项
   * @returns {this}
   */
  border2end (r, c, color = '#000000', style = 'thin', options = {}) {
    let maxCol = this.colIdx > this.maxCol ? this.colIdx : this.maxCol;
    return this.border(r, c, this.rowIdx, maxCol, color, style, options);
  }

  /**
   * 合并后续单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} [cells=1] 合并的单元格数，默认下一个单元格
   * @returns {this}
   */
  mergeCell (cells = 1) {
    this.merge(this.rowIdx, this.colIdx, this.rowIdx, this.colIdx + cells);
    this.colIdx += cells;
    return this;
  }

  /**
   * 合并后续行
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} [rows=1] 合并的行数，默认下一个行
   * @returns {this}
   */
  mergeRow (rows = 1) {
    this.merge(this.rowIdx, this.colIdx, this.rowIdx + rows, this.colIdx);
    return this;
  }

  /**
   * 合并单元格
   * @function
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @param {number} [rs] 合并开始行
   * @param {number} [cs] 合并开始列
   * @param {number} [re] 合并结束行
   * @param {number} [ce] 合并结束列
   * @returns {this}
   */
  merge (rs, cs, re, ce) {
    this.chunks.push(`sheet.GetRange('${encodeCell(rs, cs)}:${encodeCell(re, ce)}').Merge(false)`);
    return this;
  }

  /**
   * 单元格写入完成
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   */
  end () {
    if (this.colIdx > this.maxCol) {
      this.maxCol = this.colIdx;
    }
    if (this._opts.border2end) {
      let opts = typeof this._opts.border2end === 'object' ? this._opts.border2end : {
        color: this._opts.border2end
      }
      if (opts.color === true) {
        opts.color = '#000000';
      }
      this.border2end(0, 0, opts.color, opts.style || 'thin');
    }
    return this;
  }

  /**
   * 构建Sheet页数据
   * @param {*} showGridLines 是否显示表格线
   * @memberof ExcelIO.ExcelWriter.Sheet
   * @instance
   * @returns {String} 表格数据
   */
  build (showGridLines = true) {
    let chunks = Object.keys(this.funcs).map(name => this.funcs[name]).concat(this.chunks);
    chunks.push(`sheet.SetDisplayGridlines(${!!showGridLines})`);
    return `function(sheet){${chunks.join(';')}}`;
  }
}