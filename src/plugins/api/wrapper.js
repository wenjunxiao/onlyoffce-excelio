/**
 * OnlyOffice文档API
 * @external DocsAPI
 * @see {@link https://api.onlyoffice.com/editors/methods}
 */
import DocsAPI from 'onlyffice';
import { debug } from '../utils';

const BACKGROUND_GUID = `${process.env.BACKGROUND_GUID}`;
const VIEW_GUID = `${process.env.VIEW_GUID}`;

/**
 * 创建OnlyOffice文档编辑器
 * @class  DocEditor
 * @memberof external:DocsAPI
 * @classdesc OnlyOffice文档编辑器对象
 * @param {string} id 用于放置编辑器div的id,<i style="color:red;"><b>使用方法之前必须确保dom节点已存在</b></i>
 * @param {{*}} config 创建编辑器需要的配置，更多详细配置参考{@link https://api.onlyoffice.com/editors/config/}
 */
const DocEditor = DocsAPI.DocEditor;

const baseUrl = (function () {
  if (document.currentScript) {
    return document.currentScript.src.split('/').slice(0, -1).join('/')
  }
  let scripts = document.getElementsByTagName('SCRIPT');
  return scripts[scripts.length - 1];
})();

const DEFAULT_CONFIG = {
  'document': {
    'permissions': {}, // 权限相关
    'fileType': 'xlsx'
  },
  'documentType': 'spreadsheet',
  'editorConfig': {
    'lang': 'zh',
    'mode': 'edit',
    'customization': {},
    'plugins': {
      'autostart': [],
      'pluginsData': []
    },
  },
  'events': {}
}

function assignDeep (obj, data) {
  if (data) {
    Object.keys(data).forEach(key => {
      const val = data[key];
      if (Array.isArray(val)) {
        obj[key] = [].slice.call(val);
      } else if (typeof val === 'object') {
        obj[key] = assignDeep(obj[key] || {}, val);
      } else {
        obj[key] = val;
      }
    });
  }
  return obj;
}

function encode (s) {
  let dict = {};
  let data = (s + '').split('');
  let out = [];
  let currChar;
  let phrase = data[0];
  let code = 256;
  for (let i = 1; i < data.length; i++) {
    currChar = data[i];
    if (dict[phrase + currChar] !== undefined) {
      phrase += currChar;
    } else {
      out.push(phrase.length > 1 ? dict[phrase] : phrase.charCodeAt(0));
      dict[phrase + currChar] = code;
      code++;
      phrase = currChar;
    }
  }
  out.push(phrase.length > 1 ? dict[phrase] : phrase.charCodeAt(0));
  for (let j = 0; j < out.length; j++) {
    out[j] = String.fromCharCode(out[j]);
  }
  return encodeURIComponent(out.join('')).replace(/%([0-9A-F]{2})/g, function (_, p1) {
    return String.fromCharCode('0x' + p1);
  });
}

export function decode (s) {
  let dict = {};
  let data = (s + '').split('');
  let currChar = data[0];
  let oldPhrase = currChar;
  let out = [currChar];
  let code = 256;
  let phrase;
  for (let i = 1; i < data.length; i++) {
    let currCode = data[i].charCodeAt(0);
    if (currCode < 256) {
      phrase = data[i];
    } else {
      phrase = dict[currCode] ? dict[currCode] : oldPhrase + currChar;
    }
    out.push(phrase);
    currChar = phrase.charAt(0);
    dict[code] = oldPhrase + currChar;
    code++;
    oldPhrase = phrase;
  }
  return out.join('');
}

const safe = {
  '+': '-',
  '/': '_',
  '=': ''
};

/**
 * 生成URL安全的Base64文档标识。主要操作是：<br/>
 * 1、命名空间进行Base64处理；<br/>
 * 2、文档地址中的<code>pathname</code>和文档标题通过<code>?</code>连接之后进行Base64处理；<br/>
 * 3、把两个Base64值通过点号(<code>.</code>)连接起来；<br/>
 * 4、把上述结果中的<code>['+', '/', '=']</code>分别替换成功<code>['-', '_', '']</code><br/>
 * 只使用<code>pathname</code>和文档标题，而不使用完整的url的原因是可能在不同域名下使用，
 * 另外url的参数中可能有其他随机参数，比如token、username等标识用户身份的信息，这样会导致同一个文档
 * 不同的人在不同地方打开会不一样，因此只采用固定的<code>pathname</code>和文档的标题。这样也可以适用于
 * 同一个模板文档链接，用于生成不同文档的场景，这样url相同，但是文档标题不同。
 * @param {String} url 文档地址
 * @param {String} title 文档标题
 * @param {String} namespace 文档命名空间
 * @param {String} version 文档版本
 * @returns {String} URL安全的Base64文档标识
 */
function safeBase64 (url, title, namespace, version) {
  const btoa = window.btoa || window.Base64 && window.Base64.btoa;
  let keyUrl = decodeURIComponent((new URL(url)).pathname) + '?' + title;
  if (version) {
    keyUrl += version;
  }
  const key = (btoa(encodeURIComponent(namespace)) + '.' + btoa(encode(keyUrl))).replace(/(\+|\/|=)/g, $0 => {
    return safe[$0];
  });
  debug('[ExcelAPI] safe base64 url key => %s %s', keyUrl, key);
  if (key.length > 254) {
    throw new Error('document url or document title too long, please use short and id');
  }
  return key;
}

function getPluginBase () {
  if ('undefined' === typeof (extensionParams))
    return null;
  return window.extensionParams['pluginBase'];
}

function isUnset (v) {
  return typeof v === 'undefined' || v === null;
}

/**
 * 下载文件另存为事件回调
 * @typedef {function} DownloadAsCallback
 * @param {Object} event 事件
 * @param {String} event.data 缓存文档地址
 * @example
 * var docEditor = ExcelAPI.createEditor('placeholder', {
 *   events: {
 *     onDownloadAs(e) {
 *       appService.uploadFile({
 *         url: e.data
 *       })
 *     }
 *   }
 * })
 */
/**
 * 封装文件地址生成器
 * @typedef {function} WrapUrlGenerator
 * @param {String} url 原始文件地址
 * @returns {String} 最终文档地址
 */

/**
 * 扩展编辑器配置，更多详细配置参考{@link https://api.onlyoffice.com/editors/config/}
 * @typedef {Object} WrappedConfig
 * @property {{*}} [pluginBase] [扩展配置]插件根地址，默认与当前api.js相同，也可以通过`window.extensionParams.pluginBase`设置
 * @property {{*}} [namespace] [扩展配置]文档的命名空间，用于生成<code>document.key</code>
 * @property {DocKeyGenerator} [url2key] [扩展配置]将<code>document.url</code>转换成<code>document.key</code>的方法，
 * 默认使用{@link safeBase64}生成文档标识
 * @property {Boolean} [autoSync] [扩展配置]是否在保存的之后触发<code>events.onDownloadAs</code>。
 * 默认值，当配置了<code>events.onDownloadAs</code>是默认为<code>true</code>，否则默认为<code>false</code>。
 * 如果设置<code>true</code>，那么必须<code>events.onDownloadAs</code>，否则会报错。
 * @property {Array.<String>} [downloadTypes] [扩展配置]支持下载的文档类型，默认只支持<code>XLSX</code>，
 * 还可以有<code>PDF</code>、<code>CSV</code>等，其他具体查看{@link https://api.onlyoffice.com/editors/conversionapi#spreadsheet-matrix OnlyOffice文档}
 * @property {Boolean|String|WrapUrlGenerator} [downloadUrl] [扩展配置]文档下载地址或生成方法，默认使用<code>onlyoffice</code>的下载地址。
 * <code>true</code>表示与<code>document.url</code>相同；可以配置成固定的下载地址；如果配置了其他类型的下载文件（
 * 下载文档的类型取决于<code>downloadTypes</code>配置），需要配置成函数根据文件类型进行分别动态处理。
 * @property {String|WrapUrlGenerator} [printUrl] [扩展配置]文档打印地址或生成方法，默认使用<code>onlyoffice</code>的打印地址。
 * 如果配置了该选项，默认将<code>document.permissions.print</code>设置为<code>true</code>
 * @property {{*}} document 打开的文档信息
 * @property {String} document.title 打开的文档标题
 * @property {String|Number} [document.id] [扩展配置]文档ID，用于生成key，当文档标题太长时使用id来代替
 * @property {String} document.url 打开的文档链接
 * @property {String|Number} [document.short] [扩展配置]文档短链，用于生成key，当文档url太长时使用短链接来代替
 * @property {String} [document.key] 打开的文档唯一标识，不设置，按照[扩展配置]<code>url2key</code>方式生成
 * @property {String|Number} [document.version] [扩展配置]文档版本，用于生成key，当文档出现断层式变化(比如用户直接上传)时需要更新版本
 * @property {{*}} [document.permissions] 文档权限配置
 * @property {Boolean} [document.permissions.print] 文档打印权限配置，默认为<code>false</code>，除非指定了<code>printUrl</code>
 * @property {{*}} editorConfig 文档编辑器设置
 * @property {String} [editorConfig.mode] 文档编辑器模式: <code style="color:red">view</code>-预览; <code style="color:red">edit</code>-编辑
 * @property {{*}} [editorConfig.customization] 个性化定制
 * @property {Boolean} [editorConfig.customization.autosave] 是否自动保存。如果配置了
 * <code>events.onDownloadAs</code>，为了避免频繁触发，默认值就是<code>false</code>；否则默认<code>true</code>
 * @property {{*}} [editorConfig.user] 当前编辑文档的用户信息
 * @property {String} [editorConfig.user.id] 当前编辑文档的用户ID
 * @property {String} [editorConfig.user.name] 当前编辑文档的用户姓名
 * @property {Object} [events] 编辑器相关事件配置
 * @property {DownloadAsCallback} [events.onDownloadAs] 用于将缓存的文档下载另存到其他地方，可以用来同步文件
 */

/**
 * 编辑器扩展
 * @class EditorWrapper
 * @memberof ExcelAPI
 * @param {string} id 用于放置编辑器div的id
 * @param {WrappedConfig} config 创建编辑器需要的配置
 */
export default class EditorWrapper {

  constructor(id, config) {
    this.reqId = 0;
    this.executing = {};
    this.listeners = {};
    this.services = {};
    this.isChanged = false;
    this.isSaved = true;
    this.isRendered = false;
    this.ready = new Promise(resolve => {
      this.readyResolve = resolve;
    });
    config = assignDeep(assignDeep({}, DEFAULT_CONFIG), config);
    const namespace = config.namespace || 'default';
    delete config.namespace;
    /**
     * 文档标识生成器，根据文档{@link https://api.onlyoffice.com/editors/config/document#key}文档最后显示
     * 只允许<code>0-9,a-z,A-Z,-._=</code>，但是目前长度最大长度是20的限制貌似没有生效，
     * 目前超过也能正常使用，不确定在那里有限制
     * @typedef {function} DocKeyGenerator
     * @param {String} urlOrShort 文档地址(或短链)
     * @param {String} titleOrId 文档标题(或id)
     * @param {String} namespace 文档命名空间
     * @param {String|Number} version 文档版本
     */
    /**
     * @type DocKeyGenerator
     */
    const url2key = config.url2key || safeBase64;
    delete config.btoa;
    let pluginBase = (config.pluginBase || getPluginBase() || baseUrl).replace(/\/+$/, '');
    delete config.pluginBase;
    config.editorConfig.plugins.pluginsData.push(pluginBase + '/background/config.json');
    config.editorConfig.plugins.pluginsData.push(pluginBase + '/view/config.json');
    if (!(/^http/.test(config.document.url))) {
      if (config.document.url[0] !== '/') {
        config.document.url = location.href.replace(/\/[^/]*$/, '/') + config.document.url;
      } else {
        config.document.url = location.protocol + '//' + location.host + config.document.url
      }
    }
    if (!config.document.key) {
      config.document.key = url2key(config.document.short || config.document.url, config.document.id || config.document.title, namespace, config.document.version);
    }
    if (config.autoSync) {
      if (!config.events.onDownloadAs) {
        throw new Error('[autoSync]设置为`true`，[events.onDownloadAs]必须设置');
      }
    } else if (isUnset(config.autoSync) && config.events.onDownloadAs) {
      config.autoSync = true;
    }
    if (isUnset(config.editorConfig.customization.autosave) && config.events.onDownloadAs) {
      config.editorConfig.customization.autosave = false;
    }
    this.downloadUrl = config.downloadUrl; // 文档下载地址
    if (this.downloadUrl === true) {
      this.downloadUrl = config.document.url;
    }
    if (Array.isArray(config.downloadTypes)) {
      this.downloadTypes = config.downloadTypes.map(s => s.toUpperCase());
    }
    this.printUrl = config.printUrl; // 文件打印地址
    if (typeof config.document.permissions.print !== 'boolean') {
      config.document.permissions.print = !!this.printUrl;
    }
    this.autoSync = config.autoSync === true;
    this.readOnly = config.editorConfig.mode === 'view';
    this.bindEvent(config.events, 'onPluginReady', onPluginReady);
    this.bindEvent(config.events, 'onPluginClose', onPluginClose);
    this.bindEvent(config.events, 'onCommandReturn', onCommandReturn);
    this.bindEvent(config.events, 'onPluginAction', onPluginAction);
    this.bindEvent(config.events, 'onChanged', onChanged);
    this.bindEvent(config.events, 'onSaved', onSaved);
    this.bindEvent(config.events, 'onDownloadUrl', onDownloadUrl);
    this.bindEvent(config.events, 'onPrintUrl', onPrintUrl);
    this.bindEvent(config.events, 'onSheetsChanged', onSheetsChanged);
    this.bindEvent(config.events, 'onActiveSheetChanged', onActiveSheetChanged);
    debug('[ExcelAPI] config =>', config);
    this.editor = new DocEditor(id, config);
    for (let el of document.getElementsByName('frameEditor')) {
      if (el.src) {
        let url = new URL(el.src);
        let frameEditorId = url.searchParams && url.searchParams.get('frameEditorId');
        if (frameEditorId === id) {
          this.iframe = el;
          break;
        }
      }
    }
  }

  bindEvent (events, name, fn) {
    const self = this;
    if (events[name]) {
      events[name] = ((_fn) => {
        return async function () {
          let ret = await fn.apply(self, arguments);
          if (ret !== false) {
            return await _fn.apply(this, arguments);
          }
          return ret;
        }
      })(events[name]);
    } else {
      events[name] = async function () {
        return await fn.apply(self, arguments);
      }
    }
  }

  postDebug (data) {
    if (data.debug !== window.localStorage.debug) {
      this.iframe.contentWindow.postMessage(JSON.stringify({
        guid: data.guid,
        type: 'onExternalPluginMessage',
        data: {
          type: 'setDebug',
          debug: window.localStorage.debug
        }
      }), '*');
    }
  }

}

function onPluginReady (e) {
  if (e.data.guid === BACKGROUND_GUID) {
    const postData = {
      type: 'initialize',
      debug: window.localStorage.debug,
      downloadTypes: this.downloadTypes
    };
    if (this.downloadUrl) {
      postData.downloadUrl = typeof this.downloadUrl === 'string' ? this.downloadUrl : true;
    }
    if (this.printUrl) {
      postData.printUrl = typeof this.downloadUrl === 'string' ? this.printUrl : true;
    }
    this.iframe.contentWindow.postMessage(JSON.stringify({
      guid: e.data.guid,
      type: 'onExternalPluginMessage',
      data: postData
    }), '*');
    debug('[ExcelAPI] Background onPluginReady =>', e.data);
    this.backgroundReady = true;
    this.readyResolve(this.editor);
    /**
     * 编辑器命令插件准备就绪，可以执行文档命令
     * @event ExcelAPI.WrappedEditor#ready
     * @type {Object}
     * @property {ExcelAPI.WrappedEditor} target 编辑器对象
     * @property {Object} data 数据
     * @property {String} data.sheets 当前编辑器打开的文档的所有Sheet页名称
     * @property {String} data.active 当前编辑器活动的Sheet页名称
     * @example
     * editor.on('ready', async(e)=>{
     *   // 以下命令执行结果相同
     *   // let active = await e.target.getActiveSheet();
     *   let active = await e.target.executeCommand('Api.GetActiveSheet().GetName()');
     *   console.log('active sheet =>', active)
     * });
     */
    this.emit('ready', {
      target: this.editor,
      data: e.data
    });
  } else if (e.data.guid === VIEW_GUID) {
    this.postDebug(e.data);
    debug('[ExcelAPI] View onPluginReady =>', e.data);
    this.viewReady = true;
    /**
     * 编辑器视图插件打开，可以执行渲染插件操作，可以使用<code>editor.onPluginOpen</code>监听
     * @event ExcelAPI.WrappedEditor#PluginOpen
     * @alias ExcelAPI.WrappedEditor#onPluginOpen
     * @type {Object}
     * @property {ExcelAPI.WrappedEditor} target 编辑器对象
     * @property {Object} data 数据
     * @property {String} data.sheets 当前编辑器打开的文档的所有Sheet页名称
     * @property {String} data.active 当前编辑器活动的Sheet页名称
     * @example
     * editor.on('PluginOpen', async(e)=>{
     *   await e.target.renderPlugin({...});
     * });
     * // 等同于上面
     * editor.onPluginOpen = async(e)=>{
     *  await e.target.renderPlugin({...});
     * };
     */
    this.emit('PluginOpen', {
      target: this.editor,
      data: e.data
    });
  } else {
    return debug('[ExcelAPI] Unknown onPluginReady =>', e.data, BACKGROUND_GUID, VIEW_GUID);
  }
}

function onPluginClose (e) {
  if (e.data.guid === BACKGROUND_GUID) {
    this.backgroundReady = false;
    this.ready = new Promise(resolve => {
      this.readyResolve = resolve;
    });
    /**
     * 编辑器命令插件关闭
     * @event ExcelAPI.WrappedEditor#close
     * @type {Object}
     * @property {ExcelAPI.WrappedEditor} target 编辑器对象
     * @property {Object} data 数据
     */
    this.emit('close', {
      target: this.editor,
      data: e.data
    });
  } else if (e.data.guid === VIEW_GUID) {
    this.viewReady = false;
    /**
     * 编辑器视图插件关闭
     * @event ExcelAPI.WrappedEditor#PluginClose
     * @type {Object}
     * @property {ExcelAPI.WrappedEditor} target 编辑器对象
     * @property {Object} data 数据
     */
    this.emit('PluginClose', {
      target: this.editor,
      data: e.data
    });
  }
}

function onCommandReturn (e) {
  const req = this.executing[e.data.id];
  delete this.executing[e.data.id];
  if (req) {
    if (req.timer) {
      clearTimeout(req.timer);
    }
    if (e.data.success) {
      req.resolve(e.data.data);
    } else {
      req.reject(e.data.error);
    }
  }
}

async function onPluginAction (e) {
  if (e.data.guid === VIEW_GUID) {
    try {
      let result = await this.services[e.data.data.name].apply(this.services, e.data.data.args);
      this.iframe.contentWindow.postMessage(JSON.stringify({
        guid: VIEW_GUID,
        type: 'onExternalPluginMessage',
        data: {
          id: e.data.id,
          type: 'onActionReturn',
          data: {
            success: true,
            data: result
          }
        }
      }), '*');
    } catch (err) {
      this.iframe.contentWindow.postMessage(JSON.stringify({
        guid: VIEW_GUID,
        type: 'onExternalPluginMessage',
        data: {
          id: e.data.id,
          type: 'onActionReturn',
          data: {
            success: false,
            error: {
              message: err.message
            }
          }
        }
      }), '*');
    }
  }
}

function onChanged (e) {
  debug('[ExcelAPI] document changed =>', e);
  if (e.data.guid === BACKGROUND_GUID) {
    this.isChanged = true;
    this.isSaved = false;
    /**
     * 文档变化事件
     * @event ExcelAPI.WrappedEditor#changed
     * @type {Object}
     * @property {ExcelAPI.WrappedEditor} target 编辑器对象
     */
    this.emit('changed', {
      target: this.editor
    });
  }
}

function onSaved (e) {
  debug('[ExcelAPI] document saved =>', this.isSaved, e);
  if (e.data.guid === BACKGROUND_GUID && !this.isSaved) {
    this.isSaved = true;
    this.isChanged = false;
    debug('[ExcelAPI] document auto sync =>', this.autoSync);
    if (this.autoSync) {
      this.editor.downloadAs();
    }
    /**
     * 文档保存事件，可以在此事件中调用<code>downloadAs</code>触发<code>onDownloadAs</code>。
     * <div style="color: red">注意如果配置了<code>config.events.onDownloadAs</code>默认会自动触发，
     * 除非把<code>config.autoSync</code>设置为<code>false</code></div>
     * @event ExcelAPI.WrappedEditor#saved
     * @type {Object}
     * @property {ExcelAPI.WrappedEditor} target 编辑器对象
     * @example
     * var docEditor = new DocsAPI.DocEditor('placeholder', {
     *   autoSync: false, // 关闭自动触发
     *   events: {
     *     onDownloadAs(e){
     *       appService.uploadFile({url: e.data})
     *     } 
     *   }
     * });
     * docEditor.on('saved', ()=>{
     *   docEditor.downloadAs();
     * });
     */
    this.emit('saved', {
      target: this.editor
    });
  }
}

function onDownloadUrl (e) {
  debug('[ExcelAPI] onDownloadUrl =>', e);
  if (e.data.guid === BACKGROUND_GUID) {
    Promise.resolve(this.downloadUrl(e.data.url)).then(url => {
      this.iframe.contentWindow.postMessage(JSON.stringify({
        guid: e.data.guid,
        type: 'onExternalPluginMessage',
        data: {
          type: 'download',
          url
        }
      }), '*');
    });
  }
}

function onPrintUrl (e) {
  debug('[ExcelAPI] onPrintUrl =>', e);
  if (e.data.guid === BACKGROUND_GUID) {
    Promise.resolve(this.printUrl(e.data.url)).then(url => {
      this.iframe.contentWindow.postMessage(JSON.stringify({
        guid: e.data.guid,
        type: 'onExternalPluginMessage',
        data: {
          type: 'print',
          url,
          downloadType: e.data.downloadType
        }
      }), '*');
    });
  }
}

function onSheetsChanged (e) {
  debug('[ExcelAPI] onSheetsChanged =>', e);
  if (e.data.guid === BACKGROUND_GUID) {
    /**
     * Sheet页变化事件,也可以通过<code>config.events.onSheetsChanged</code>监听
     * @event ExcelAPI.WrappedEditor#SheetsChanged
     * @type {Object}
     * @property {Object} data 数据
     * @property {String} data.sheets 当前编辑器打开的文档的所有Sheet页名称
     * @property {String} data.active 当前编辑器活动的Sheet页名称
     * @example
     * // editor.onSheetsChanged = function(){};
     * editor.on('SheetsChanged', async(e)=>{
     *   console.log('sheets =>', e.data.sheets)
     * });
     */
    this.emit('SheetsChanged', {
      target: this.editor,
      data: e.data
    });
  }
}

function onActiveSheetChanged (e) {
  debug('[ExcelAPI] onActiveSheetChanged =>', e);
  if (e.data.guid === BACKGROUND_GUID) {
    /**
     * Sheet活动页变化事件,也可以通过<code>config.events.onActiveSheetChanged</code>监听
     * @event ExcelAPI.WrappedEditor#ActiveSheetChanged
     * @type {Object}
     * @property {Object} data 数据
     * @property {String} data.sheets 当前编辑器打开的文档的所有Sheet页名称
     * @property {String} data.active 当前编辑器活动的Sheet页名称
     * @example
     * // editor.onActiveSheetChanged = function(){};
     * editor.on('ActiveSheetChanged', async(e)=>{
     *   console.log('active sheet =>', e.data.active)
     * });
     */
    this.emit('ActiveSheetChanged', {
      target: this.editor,
      data: e.data
    });
  }
}