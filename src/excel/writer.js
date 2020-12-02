import { assign } from 'lodash';
import Sheet from './sheet';

export {
  Sheet
};

/**
 * Excel文档生成类
 * @class ExcelWriter
 * @memberof ExcelIO
 * @param {Object} options 选项
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
export default class ExcelWriter {

  constructor(options) {
    this.Sheets = [];
    this.SheetNames = [];
    this.rowIdx = -1;
    this.colIdx = -1;
    this.maxCol = 0;
    this._opts = assign({}, options || {});
    this.NaN = this._opts.NaN || '';
  }

  /**
   * 不显示表格线
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @returns {this}
   */
  withoutGridLines () {
    this._opts.showGridLines = false;
    return this;
  }

  /**
   * 当前活动Sheet页
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @returns {this}
   */
  active () {
    if (!this.curSheet) {
      this.curSheet = new Sheet(assign({}, this._opts, {
        owner: this,
        sheetName: true
      }), this._watermark);
      this.SheetNames.push(true);
      this.Sheets.push(this.curSheet);
    }
    return this.curSheet;
  }

  /**
   * 当前操作的行号
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @returns {this}
   */
  rowIndex () {
    return this.curSheet.rowIndex();
  }

  /**
   * 当前操作的列号
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @returns {this}
   */
  colIndex () {
    return this.curSheet.colIndex();
  }

  /**
   * 获取或新建指定名称的Sheet页
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {String} name Sheet页名称
   * @returns {ExcelIO.ExcelWriter.Sheet}
   */
  newSheet (name) {
    let pos = this.SheetNames.indexOf(name);
    if (pos < 0) {
      this.curSheet = new Sheet(assign({}, this._opts, {
        owner: this,
        sheetName: name
      }), this._watermark);
      this.SheetNames.push(name);
      this.Sheets.push(this.curSheet);
      return this.curSheet.clear();
    } else {
      return this.Sheets[pos].clear();
    }
  }

  /**
   * 切换到指定Sheet页(不存在则创建)
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {String} name Sheet页名称
   * @returns {this}
   */
  sheet (name) {
    this.endSheet();
    let pos = this.SheetNames.indexOf(name);
    if (pos < 0) {
      this.curSheet = new Sheet(assign({}, this._opts, {
        owner: this,
        sheetName: name
      }), this._watermark);
      this.SheetNames.push(name);
      this.Sheets.push(this.curSheet.clear());
    } else {
      this.curSheet = this.Sheets[pos].clear();
    }
    return this;
  }

  /**
   * 重命名指定Sheet页的名称
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {String} name 新的名称
   * @param {String} from 需要修改的Sheet页名称
   * @returns {this}
   */
  rename (name, from) {
    if (this.SheetNames.indexOf(name) > -1) {
      throw new Error(`Sheet with name [${name}] already exists`);
    }
    if (from) {
      if (from !== name) {
        let pos = this.SheetNames.indexOf(from);
        this.SheetNames.splice(pos, 1, name);
        this.Sheets[pos].rename(name);
      }
    } else if (name !== this.curSheet.sheetName) {
      from = this.curSheet.sheetName;
      let pos = this.SheetNames.indexOf(from);
      this.SheetNames.splice(pos, 1, name);
      this.curSheet.rename(name);
    }
    return this;
  }

  /**
   * 跳过指定行数
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} [rows=1] 跳过的行数
   * @param {number} [cells=0] 同时跳过的单元格，默认在行首
   * @returns {this}
   */
  skipRow (rows = 1, cells = 0) {
    this.curSheet.skipRow(rows, cells);
    return this;
  }

  /**
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} [cells=1] 跳过的单元格
   * @returns {this}
   */
  skipCell (cells = 1) {
    this.curSheet.skipCell(cells);
    return this;
  }

  /**
   * 跳转到指定的行列
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} row 指定行号
   * @param {number} [cell=0] 指定列号，默认行首
   * @returns {this}
   */
  go (row, cell = 0) {
    this.curSheet.go(row, cell);
    return this;
  }

  /**
   * 定位到下几行(默认下一行)
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} [cells=0] 同时跳转到对应的列，默认行首
   * @returns {this} 
   */
  row (cells = 0) {
    this.curSheet.row(cells);
    return this;
  }

  /**
   * 在当前单元格写入数字，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} v 数字
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  number (v, options) {
    this.curSheet.number(v, options);
    return this;
  }

  /**
   * 在当前单元格写入布尔值，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} v 布尔值
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  boolean (v, options) {
    this.curSheet.boolean(v, options);
    return this;
  }

  /**
   * 在当前单元格写入字符串，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} v 字符串
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  string (v, options) {
    this.curSheet.string(v, options);
    return this;
  }

  /**
   * 在当前单元格写入UTC时间，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} v 数字
   * @param {String} [format] 时间格式化字符串
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  utc (v, format, options) {
    this.curSheet.utc(v, format, options);
    return this;
  }

  /**
   * 在当前单元格写入当前时区的时间，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} v 数字
   * @param {String} [format] 时间格式化字符串
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  date (v, format, options) {
    this.curSheet.date(v, format, options);
    return this;
  }

  formatNumber (v, precision) {
    return this.curSheet.formatNumber(v, precision);
  }

  /**
   * 在当前单元格写入金额，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} v 金额
   * @param {String} currency 币种
   * @param {String} [precision] 精度
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  currency (v, currency, precision, options) {
    this.curSheet.currency(v, currency, precision, options);
    return this;
  }

  /**
   * 在当前单元格写入百分数，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} v 百分数
   * @param {String} [precision] 精度
   * @param {Object} [options={}] 单元格选项
   * @returns {this}
   */
  percent (v, precision, options) {
    this.curSheet.percent(v, precision, options);
    return this;
  }

  /**
   * 在当前单元格写入值，并跳转到下一个单元格
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {*} v 值
   * @param {Object} [options={}] 单元格选项
   * @param {String} [type] 值类型，默认根据值的类型
   * @param {String} [format] 值的格式化
   * @returns {this}
   */
  cell (v, options, type, format) {
    this.curSheet.cell(v, options, type, format);
    return this;
  }

  /**
   * 指定单元格的宽度
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} width 宽度
   * @param {number} [colIndex] 指定列，默认当前刚写入的单元格
   * @returns {this}
   */
  width (width, colIndex = -1) {
    this.curSheet.width(width, colIndex);
    return this;
  }

  chWidth (width, colIndex = -1) {
    this.curSheet.chWidth(width, colIndex);
    return this;
  }

  /**
   * 指定单元格的颜色
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {String} bgColor 背景颜色，十六进制
   * @param {String} fgColor 字体颜色，十六进制
   * @returns {this}
   */
  color (bgColor, fgColor) {
    this.curSheet.color(bgColor, fgColor);
    return this;
  }

  /**
   * 指定单元格的边框
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} rs 单元格开始行
   * @param {number} cs 单元格开始列
   * @param {number} re 单元格结束行
   * @param {number} ce 单元格结束列
   * @param {String} [color='#000000'] 背景颜色，十六进制，默认黑色
   * @param {String} [style] 单元格样式，默认<code>thin</code>
   * @param {Object} [options] 选项
   * @returns {this}
   */
  border (rs, cs, re, ce, color = '#000000', style = 'thin', options = {}) {
    this.curSheet.border(rs, cs, re, ce, color, style, options);
    return this;
  }

  /**
   * 设置从指定行列到文档结束的边框
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} r 单元格开始行
   * @param {number} c 单元格开始列
   * @param {String} [color='#000000'] 背景颜色，十六进制，默认黑色
   * @param {String} [style] 单元格样式，默认<code>thin</code>
   * @param {Object} [options] 选项
   * @returns {this}
   */
  border2end (r, c, color = '#000000', style = 'thin', options = {}) {
    this.curSheet.border2end(r, c, color, style, options);
    return this;
  }

  /**
   * 合并后续单元格
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} [cells=1] 合并的单元格数，默认下一个单元格
   * @returns {this}
   */
  mergeCell (cells = 1) {
    this.curSheet.mergeCell(cells);
    return this;
  }

  /**
   * 合并后续行
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} [rows=1] 合并的行数，默认下一个行
   * @returns {this}
   */
  mergeRow (rows = 1) {
    this.curSheet.mergeRow(rows);
    return this;
  }

  /**
   * 合并单元格
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {number} [rs] 合并开始行
   * @param {number} [cs] 合并开始列
   * @param {number} [re] 合并结束行
   * @param {number} [ce] 合并结束列
   * @returns {this}
   */
  merge (rs, cs, re, ce) {
    this.curSheet.merge(rs, cs, re, ce);
    return this;
  }
  /**
   * 添加水印
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @returns {this}
   */
  watermark () {
    return this;
  }
  /**
   * 不需要水印
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @returns {this}
   */
  withoutWatermark () {
    return this;
  }

  /**
   * 所有Sheet都已经完成
   * @function
   * @private
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @returns {this}
   */
  endSheet () {
    for (let sheet of this.Sheets) {
      sheet.end();
    }
    this.curSheet = null;
    return this;
  }

  /**
   * 构建目标
   * @function
   * @private
   * @memberof ExcelIO.ExcelWriter
   * @instance
   */
  build2target () {
    this.endSheet();
    const tasks = [];
    const sheets = this.Sheets.map((sheet, i) => {
      if (sheet.sheetName === true) {
        tasks.push(`function(){
          var args = arguments[0] || [];
          args.push(Api.GetActiveSheet().Name);
          return args;
        }`);
      } else {
        /**
         * 检测当前需要的Sheet页是否存在，不存在则新增；
         * 是否需要通过监听`asc_onSheetsChanged`事件确定是否新增完成？
         * 暂时不需要，因为分步执行的时候存在多次交互，正常应该已经创建完成
         */
        tasks.push(`function(){
          var args = arguments[0] || [];
          if(!Api.GetSheet('${sheet.sheetName}')){
            Api.AddSheet('${sheet.sheetName}');
          }
          args.push('${sheet.sheetName}');
          return args;
        }`);
      }
      return `(${sheet.build(this._opts.showGridLines !== false)})(Api.GetSheet(args[${i}]))`;
    });
    /**
     * 因为执行`Api.AddSheet`是异步的，如果之后立即执行`Api.GetSheet`可能会返回null；
     * 因此，先获取Sheet页列表，确保所有需要的Sheet页存在，再执行具体的命令
     */
    return tasks.concat([`function(){
      var args = arguments[0] || [];
      ${sheets.join(';')}
    }`]);
  }

  /**
   * 构建并返回可执行的命令数组
   * @function
   * @memberof ExcelIO.ExcelWriter
   * @instance
   * @param {{}} options
   * @returns {Array.<String>}
   */
  build (options = {}) {
    return this.build2target(options);
  }

}