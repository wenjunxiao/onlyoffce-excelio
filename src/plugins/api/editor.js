import ExcelIO from 'onlyoffice-excelio';
import EditorWrapper from './wrapper';
import { debug, getParams } from '../utils';

const Excel = ExcelIO;
const ExcelWriter = ExcelIO.ExcelWriter;

const BACKGROUND_GUID = `${process.env.BACKGROUND_GUID}`;
const VIEW_GUID = `${process.env.VIEW_GUID}`;

export class ActionTimeout extends Error {
  /**
   * 操作超时错误
   * @constructor ActionTimeout
   * @memberof ExcelAPI
   * @param {String} message 超时错误
   * @param {String} action 操作
   */
  constructor(message, action) {
    super(message);
    this.action = action;
  }
}

/**
 * 扩展编辑器
 * @class WrappedEditor
 * @memberof ExcelAPI
 * @augments external:DocsAPI.DocEditor
 * @param {String} id 用于放置编辑器div的id
 * @param {WrappedConfig} config 创建编辑器需要的配置
 */
export default function WrappedEditor (id, config) {
  const self = new EditorWrapper(id, config);
  const editor = self.editor;
  /**
   * 执行OnlyOffice操作命令
   * @function executeCommand
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @param {String} command OnlyOffice操作命令，具体命令查看{@link https://api.onlyoffice.com/docbuilder/spreadsheetapi SpreadSheetApi}
   * @param {number} [timeout=0] 操作超时时间，单位毫秒
   * @returns {Promise} 命令操作结果
   * @throws {ExcelAPI.ActionTimeout} 执行超时
   * @example
   * // 获取活动页的名称
   * const activeSheetName = await docEditor.executeCommand('Api.GetActiveSheet().GetName()');
   */
  const executeCommand = editor.executeCommand = function (command, timeout = 0) {
    debug('[ExcelAPI] executeCommand[' + timeout + '] =>', command);
    if (!command) {
      return Promise.reject(new Error('Command cannot be empty'));
    }
    if (!self.backgroundReady) {
      return Promise.reject(new Error('Plugin not ready.'));
    }
    const id = ++self.reqId;
    return new Promise((resolve, reject) => {
      const req = self.executing[id] = {
        time: Date.now(), resolve, reject
      };
      if (timeout > 0) {
        req.timer = setTimeout(() => {
          delete self.executing[id];
          req.reject(new ActionTimeout('Execute command timeout', command));
        }, timeout);
      }
      self.iframe.contentWindow.postMessage(JSON.stringify({
        guid: BACKGROUND_GUID,
        type: 'onExternalPluginMessage',
        data: {
          id,
          type: 'executeCommand',
          command
        }
      }), '*');
    });
  };
  /**
   * 是否只读
   * @member {Boolean} readOnly
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @readonly
   */
  Object.defineProperty(self.editor, 'readOnly', {
    writable: false,
    value: self.readOnly
  });
  /**
   * 文档是否有变化
   * @member {Boolean} isChanged
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @readonly
   */
  Object.defineProperty(self.editor, 'isChanged', {
    get () {
      return self.isChanged
    }
  });
  /**
   * 文档是否已保存
   * @member {Boolean} isSaved
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @readonly
   */
  Object.defineProperty(self.editor, 'isSaved', {
    get () {
      return self.isSaved
    }
  });
  /**
   * 获取/设置是否自动同步
   * @member {Boolean} autoSync
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   */
  Object.defineProperty(self.editor, 'autoSync', {
    get () {
      return self.autoSync
    },
    set (v) {
      self.autoSync = !!v;
      return self.autoSync;
    }
  });
  /**
   * 就绪回调函数
   * @callback ReadyCallback
   * @param {ExcelAPI.WrappedEditor} editor 就绪的编辑器对象
   */
  /**
   * 编辑器插件准备就绪
   * @function ready
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @param {ReadyCallback|undefined} [callback] 就绪回调函数
   * @returns {Promise.<this>}
   * @example
   * docEditor.ready((editor)=>{
   *   const sheets = await editor.getSheetNames();
   *   // const sheets = await docEditor.getSheetNames();
   *   console.log('sheets =>', sheets, docEditor === editor);
   * });
   */
  editor.ready = function (callback) {
    if (callback) {
      return self.ready.then(rsp => {
        return Promise.resolve(callback(rsp)).then(() => {
          return rsp;
        });
      });
    }
    return self.ready;
  };
  /**
   * 监听事件，所有监听都可以使用<code>on</code>+<code>eventName</code>的函数代替，比如
   * <code>on('PluginOpen')</code>和<code>onPluginOpen</code>
   * 都能够监听{@link ExcelAPI.WrappedEditor#event:PluginOpen PluginOpen}事件。
   * 使用<code>on</code>+<code>eventName</code>函数代替时区分大小写(或定义成全小写)
   * @function on
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @param {String} eventName 监听的事件名称(不区分大小写)
   * @param {function} listener 监听函数
   * @returns {this}
   * @example
   * function onPluginOpen(e){}
   * // `on('name')` 不区分大小写
   * docEditor.on('PluginOpen', onPluginOpen); // 可以监听
   * docEditor.on('pluginOpen', onPluginOpen); // 可以监听
   * docEditor.on('pluginopen', onPluginOpen); // 可以监听
   * // `on + name` 函数区分大小写或全小写
   * docEditor.onPluginOpen = onPluginOpen; // 完全匹配，可以监听
   * // docEditor.onPluginopen = onPluginOpen; // 不完全匹配，不可以监听
   * docEditor.onpluginopen = onPluginOpen; // 全小写，可以监听
   */
  editor.on = function (eventName, listener) {
    eventName = eventName.toLowerCase()
    let listeners = self.listeners[eventName];
    if (!listeners) {
      listeners = self.listeners[eventName] = [];
    }
    listeners.push(listener);
    return this;
  };
  /**
   * 触发事件
   * @function emit
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @param {String} eventName 触发的事件名称
   * @param {...*} args 触发事件回调参数
   */
  self.emit = editor.emit = function (eventName, ...args) {
    let listeners = self.listeners[eventName.toLowerCase()];
    listeners && listeners.forEach(listener => listener(...args));
    const fn = editor['on' + eventName] || editor['on' + eventName.toLowerCase()];
    if (typeof fn === 'function') {
      fn(...args);
    }
    return this;
  };
  /**
   * 获取当前文档所有Sheet名称
   * @function getSheetNames
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @param {number} [timeout=0] 等待超时时间，单位毫秒
   * @returns {Promise} 返回所有Sheet页名称的Promise
   * @throws {ExcelAPI.ActionTimeout} 执行超时
   */
  editor.getSheetNames = function (timeout = 0) {
    return executeCommand('Api.GetSheets().map(i => i.Name)', timeout);
  };
  /**
   * 获取当前文档活动页名称
   * @function getActiveSheet
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @param {number} [timeout=0] 等待超时时间，单位毫秒
   * @returns {Promise} 返回当前活动Sheet页的名称
   * @throws {ExcelAPI.ActionTimeout} 执行超时
   */
  editor.getActiveSheet = function (timeout = 0) {
    return executeCommand('Api.GetActiveSheet().GetName()', timeout);
  };
  /**
   * 执行代码
   * @function executeCode
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @param {String} code 执行通过{@link module:ExcelIO}操作Excel的代码
   * @param {{*}} [args={}] 执行代码中所需要的参数
   * @param {number} [timeout=0] 执行操作超时时间，单位毫秒
   * @returns {Promise.<*>} 代码的返回结果
   * @throws {ExcelAPI.ActionTimeout} 执行超时
   * @example
import _ from 'lodash';
// ...
const data = [{v1: 'String', v2: new Date(), v3: 3.14159}];
docEditor.executeCode(`(function(name){
  let data = ${JSON.stringify(data)}
  let writer = Excel.createWriter({});
  writer.newSheet(name).row().cell('Title1').cell('Title2').cell('Title3');
  for(let d of data) {
    writer.row().cell(d.v1).date(d.v2, 'YYYY-MM-DD HH:mm:ss').currency(_.round(d.v3, 2), '$')
  }
  return writer.build({});
})(name)`, {name: 'Sheet1', _});
   */
  editor.executeCode = function (code, args = {}, timeout = 0) {
    if (!args) args = {};
    debug('[ExcelAPI] executeCode[' + timeout + '] =>', code, args);
    try {
      const names = Object.keys(args);
      const fn = new Function('ExcelIO', 'Excel', 'ExcelWriter', ...names, `return (${code})`);
      const cmd = fn(ExcelIO, Excel, ExcelWriter, ...names.map(name => args[name]));
      if (Array.isArray(cmd)) {
        const executeNext = (data) => {
          const args = typeof data === undefined ? '' : JSON.stringify(data);
          return executeCommand(`(${cmd.shift()})(${args})`, timeout).then(data => {
            if (cmd.length > 0) {
              return executeNext(data);
            }
            return data;
          });
        }
        return executeNext();
      }
      return executeCommand(cmd, timeout);
    } catch (err) {
      return Promise.reject(err);
    }
  };
  /**
   * 调用视图插件中的方法，包括<code>renderPlugin</code>时传入的<code>methods</code>中的方法
   * @function invokeView
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @param {String} name 方法名称
   * @param {...*} [args] 方法参数
   * @returns {Promise} 返回方法调用结果的Promise
   */
  const invokeView = editor.invokeView = function (name, ...args) {
    const id = ++self.reqId;
    return new Promise((resolve, reject) => {
      self.executing[id] = {
        time: Date.now(),
        resolve,
        reject
      };
      self.iframe.contentWindow.postMessage(JSON.stringify({
        guid: VIEW_GUID,
        type: 'onExternalPluginMessage',
        data: Object.assign({}, config, {
          id,
          type: 'invoke',
          name,
          args
        })
      }), '*');
    });
  };
  /**
   * 是否已经渲染完成
   * @member {Boolean} isRendered
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @readonly
   */
  Object.defineProperty(self.editor, 'isRendered', {
    get () {
      return self.isRendered;
    }
  })
  /**
   * 编辑器的插件扩展<code>renderPlugin</code>传入的<code>methods</code>，
   * 插件渲染完成之后，可以直接通过<code>editor.methods.</code>调用其中方法，
   * 所有方法都是异步调用，并且传入的参数必须是与dom无关的可序列化的参数
   * @member {Object} methods
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @readonly
   * @example
   * await docEditor.renderPlugin({
   *   methods: {
   *     appendLogToView(text) {
   *       document.getElementById('log').value += text + '\n';
   *       return document.getElementById('log').value.length;
   *     }
   *   }
   * })
   * if (docEditor.isRendered) {
   *   let len = await docEditor.methods.appendLogToView('test view log');
   *   console.log('total log length =>', len);
   * }
   */
  self.editor.methods = {};
  /**
   * 渲染插件
   * @function renderPlugin
   * @memberof ExcelAPI.WrappedEditor
   * @instance
   * @param {Object} config 渲染插件的配置信息
   * @param {String} config.template 渲染插件的html模版
   * @param {Object} [config.methods={}] 插件需要用到的与<code>插件dom</code>及其相关操作的。
   * 方法中只能使用通过<code>config.methods</code>、<code>config.template</code>和
   * <code>config.script</code>等注入到插件中的对象方法和dom。除此之外，还可以直接使用插件内置提供的方法{@link module:plugins/View}，
   * 比如<code>executeCommand('Api.GetActiveSheet().GetName()')</code>。使用<code>config.services</code>中
   * 的方法的时候最好使用<code>this.services</code>，因为项目打包的时候可能会压缩重命名
   * @param {Object} [config.services={}] 插件需要用到的与<code>插件dom</code>无关的操作服务。除了注入到插件中的方法和dom，
   * 方法中能用当前项目中的一切，包括当前项目的dom
   * @param {Array.<Object>} [config.link=[]] 渲染插件页面需要插入的<code>&lt;link&gt;</code>。
   * 对象的属性会作为标签的属性：<dl>
   * <dt><i>href</i></dt><dd>外部样式链接</dd>
   * <dt><i>async</i></dt><dd>是否需要等待该文件加载完成才触发<code>onload</code>事件，
   * 值为<code>"async"</code>表示不需要等待，默认需要等待加载完成</dd>
   * </dl>
   * 其他属性参考HTML标准说明
   * @param {Array.<Object|String>} [config.style=[]] 渲染插件页面需要插入的<code>&lt;style&gt;</code>。
   * 如果是字符串，直接作为标签的内容；如果是对象，那么除了对象的<code>html</code>属性会作为标签的内容外，
   * 其他都会作为标签的属性，具体参考HTML标准说明。
   * @param {Array.<Object|String>} [config.script=[]] 渲染插件页面需要插入的<code>&lt;script&gt;</code>。
   * 如果是字符串，直接作为标签的内容；如果是对象：<dl>
   * <dt><i>head</i></dt><dd>扩展属性，控制脚本是在<code>head</code>还是<code>body</code>中，在
   * <code>body</code>是放在<code>config.template</code>的后面</dd>
   * <dt><i>html</i></dt><dd>扩展属性，会作为标签的内容</dd>
   * <dt><i>src</i></dt><dd>脚本链接</dd>
   * <dt><i>async</i></dt><dd>除了HTML标准中定义了是否异步加载之外，在插件中还指示是否需要等待该文件加载完成才触发
   * <code>onload</code>事件，值为<code>"async"</code>表示不需要等待，默认需要等待加载完成</dd>
   * </dl>
   * 其他属性参考HTML标准的定义。
   * @param {Boolean|String} [config.iframe=false] 渲染的视图是否放在<code>iframe</code>，
   * 默认直接内嵌在插件的dom中。也可以指定一个链接来作为插件页面，此时其他模版设置均不生效。
   * 由于<code>iframe</code>中没有处理鼠标移动相关事件，界面的拖动会有问题。
   * @param {Boolean} [config.xhr=true] 是否通过ajax并行加载脚本资源，非<code>iframe</code>模式下有效，
   * 默认为<code>true</code>。如果为<code>false</code>，非异步的资源会按照顺序等上一个下载完成再下载另一个。
   * @returns {Promise} 返回渲染完成的Promise
   * @throws {ExcelAPI.ActionTimeout} 执行超时
   * @example
docEditor.on('PluginOpen', async (e) => {
  const services = {
    postService(data) {
      return new Promise((resolve, reject) => {
        let xhr = new XMLHttpRequest();
        xhr.open('post', '/api/service');
        xhr.onerror = reject;
        xhr.onload = () => resolve(JSON.parse(xhr.responseText));
        xhr.send(JSON.stringify(data));
      });
    }
  };
  await e.target.renderPlugin({
    template: `<textarea type="text" id="data" style="width:98%;"></textarea>
      <button onclick="doCommit(this)">Commit</button>
      <textarea type="text" id="result" style="width:98%;"></textarea>`,
    methods: {
      doCommit(that) {
        let data = JSON.parse(document.getElementById('data').value || '{}');
        services.postService(data).then(rsp => {
          if (rsp.success) {
            document.getElementById('result').value = '成功:' + JSON.stringify(rsp);
          } else {
            document.getElementById('result').value = '失败:' + JSON.stringify(rsp);
          }
        }).catch(err => {
          document.getElementById('result').value = '异常:' + err.message;
        });
      }
    },
    services
  })
});
   */
  editor.renderPlugin = function (config, timeout = 0) {
    if (!config) return Promise.reject(new Error('missing config'));
    if (!config.template && !config.iframe) return Promise.reject(new Error('missing template'));
    if (!self.viewReady) {
      return Promise.reject(new Error('Plugin not ready, please use this in `PluginOpen` event.'));
    }
    self.services = config.services || {};
    self.editor.methods = {};
    const methods = config.methods || {};
    const id = ++self.reqId;
    return new Promise((resolve, reject) => {
      const req = self.executing[id] = {
        time: Date.now(),
        resolve: (function (resolve) {
          return function (data) {
            self.isRendered = true;
            resolve(data);
            /**
             * 编辑器视图插件渲染完成
             * @event ExcelAPI.WrappedEditor#PluginRendered
             * @type {Object}
             * @property {ExcelAPI.WrappedEditor} target 编辑器对象
             */
            self.emit('PluginRendered', {
              target: self.editor
            })
          };
        })(resolve),
        reject: (function (reject) {
          return function (err) {
            reject(err);
            /**
             * 编辑器视图插件渲染出错
             * @event ExcelAPI.WrappedEditor#RenderingError
             * @type {Object}
             * @property {Error} error 错误
             */
            self.emit('RenderingError', err)
          };
        })(reject)
      };
      if (timeout > 0) {
        req.timer = setTimeout(() => {
          delete self.executing[id];
          req.reject(new ActionTimeout('Render plugin timeout', 'renderPlugin'));
        }, timeout);
      }
      self.iframe.contentWindow.postMessage(JSON.stringify({
        guid: VIEW_GUID,
        type: 'onExternalPluginMessage',
        data: Object.assign({}, config, {
          id,
          type: 'executeRender',
          methods: Object.keys(methods).reduce((r, n) => {
            Object.defineProperty(self.editor.methods, n, {
              writable: false,
              value: invokeView.bind(self.editor, n)
            });
            r[n] = methods[n].toString();
            return r;
          }, {}),
          services: Object.keys(self.services).reduce((r, n) => {
            r[n] = getParams(self.services[n]);
            return r;
          }, {})
        })
      }), '*');
    });
  };
  return editor;
}