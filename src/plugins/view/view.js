/**
 * 视图插件，运行在<code>iframe</code>中，所有的方法都在<code>iframe</code>的<code>window</code>中
 * @module plugins/View
 */

import debugFactory from 'debug-factory';

/**
 * 调试日志，开启调试日志的方式是在<code>localStorage.debug</code>增加<code>excel-plugin</code>。
 * 比如<code>localStorage.debug +=',excel-plugin'</code>
 * @function debug
 * @param {...*} args 日志参数
 */
window.debug = debugFactory('excel-plugin');

/**
 * 插件名称
 * @member pluginName
 * @readonly
 */
window.pluginName = `${process.env.PLUGIN_NAME}`;
window.reqId = 0; // 请求ID
window.requests = {}; // 请求数据

/**
 * 当前脚本的父级路径
 * @private
 */
const baseUrl = (function () {
  if (document.currentScript) {
    return document.currentScript.src.split('/').slice(0, -1).join('/')
  }
  let scripts = document.getElementsByTagName('SCRIPT');
  return scripts[scripts.length - 1];
})();

/**
 * iframe里面的操作
 */
window.Asc.plugin.event_iframeAction = function (data) {
  const id = ++window.reqId;
  window.requests[id] = {
    iframe: true,
    data
  };
  window.debug('[%s] onPluginAction.iframe =>', window.pluginName, id, data);
  window.parent.parent.postMessage(window.JSON.stringify({
    frameEditorId: window.frameEditorId,
    event: 'onPluginAction',
    data: {
      guid: window.Asc.plugin.guid,
      id,
      type: data.type,
      data: data.data
    }
  }), '*');
};

window.Asc.plugin.event_iframeCommand = function (data) {
  window.debug('[%s] iframeCommand =>', window.pluginName, data.id, data.command);
  executeCommand(data.command).then(rsp => {
    const iframe = document.getElementById('view');
    iframe.contentWindow.postMessage({
      id: data.id,
      data: {
        success: true,
        data: rsp
      }
    }, '*');
  }).catch(error => {
    const iframe = document.getElementById('view');
    iframe.contentWindow.postMessage({
      id: data.id,
      data: {
        success: false,
        error
      }
    }, '*');
  });
}

function genIframeScripts (script, head, htmls) {
  if (!script) return 0;
  let count = 0;
  for (let item of script) {
    if ((!!item.head) !== head) continue;
    count++;
    if (typeof item === 'string') {
      htmls.push(`<script type="text/javascript">${item}</script>`);
    } else {
      delete item.head;
      if (!item.type) item.type = 'text/javascript';
      htmls.push(`<script ${Object.keys(item).map(attr => {
        if (attr === 'html' || attr === 'head') return '';
        return `${attr}="${item[attr]}"`
      }).join(' ')}>${item.html || ''}</script>`);
    }
  }
  return count;
}

function loadResource (src) {
  let xhr = new XMLHttpRequest();
  window.debug('[%s] resource loading =>', window.pluginName, src);
  xhr.open('get', src, true);
  return new Promise((resolve, reject) => {
    xhr.onerror = reject;
    xhr.onload = () => {
      window.debug('[%s] resource loaded =>', window.pluginName, src, xhr.responseText.length);
      return resolve(xhr.responseText);
    };
    xhr.send();
  });
}

function executeInIframe (msg) {
  const methods = msg.methods || {};
  let htmls = [];
  if (Object.keys(methods).length > 0) {
    htmls.push(`<script type="text/javascript" src="${baseUrl}/runtime.js"></script>`);
  }
  if (msg.link) {
    for (let item of msg.link) {
      if (!item.rel && /.css$/.test(item.href)) {
        item.rel = 'stylesheet';
      }
      htmls.push(`<link ${Object.keys(item).map(attr => {
        return `${attr}="${item[attr]}"`
      }).join(' ')} />`);
    }
  }
  if (msg.style) {
    for (let item of msg.style) {
      if (typeof item === 'string') {
        htmls.push(`<style>${item}</style>`);
      } else {
        htmls.push(`<style ${Object.keys(item).map(attr => {
          if (attr === 'html') return '';
          return `${attr}="${item[attr]}"`
        }).join(' ')}>${item.html || ''}</style>`);
      }
    }
  }
  let count = htmls.length;
  count += genIframeScripts(msg.script, true, htmls);
  htmls.push('<script>',
    'window.services = {};',
    'window.requests = {};',
    'window.reqId = 0;'
  );
  htmls.push(`window.plugin_postMessage=function(msg){
    window.parent.postMessage(JSON.stringify({
      guid: '${window.Asc.plugin.guid}',
      type: 'onEvent',
      eventName: 'iframeAction',
      eventData: msg
    }), '*');
  }`);
  htmls.push(`window.addEventListener('message', function(e){
    let msg = e.data;
    const req = requests[msg.id];
    delete requests[msg.id];
    if (req) {
      if (msg.data.success) {
        req.resolve(msg.data.data);
      } else {
        req.reject(msg.data.error);
      }
    }
  })`);
  htmls.push(`window.executeCommand=function(command){
    let id = ++window.reqId;
    let req = window.requests[id] = {id};
    req.promise = new Promise((resolve, reject)=>{
      req.resolve = resolve;
      req.reject = reject;
    });
    window.parent.postMessage(JSON.stringify({
      guid: '${window.Asc.plugin.guid}',
      type: 'onEvent',
      eventName: 'iframeCommand',
      eventData: {
        id,
        command
      }
    }), '*');
    return req.promise;
  }`);
  Object.keys(methods).forEach(name => {
    let fn = methods[name];
    if (/^\s*\w+\s*\(/im.test(fn)) {
      fn = 'function ' + fn;
    }
    htmls.push(`window['${name}']=${fn};`);
    count++;
  });
  const services = msg.services || {};
  Object.keys(services).forEach(name => {
    count++;
    htmls.push(`window.services['${name}']= function (){
      let id = ++window.reqId;
      let req = window.requests[id] = {id};
      req.promise = new Promise((resolve, reject)=>{
        req.resolve = resolve;
        req.reject = reject;
      });
      plugin_postMessage({
        id,
        type: 'service',
        data: {
          name: '${name}',
          args: [].slice.call(arguments)
        }
      });
      return req.promise;
    }`);
  });
  htmls.push('</script>')
  htmls.push(msg.template);
  count += genIframeScripts(msg.script, false, htmls);
  const iframe = document.createElement('iframe');
  iframe.id = 'view';
  iframe.setAttribute('style', 'position:absolute;left:0;top:0px;right:0;bottom:0;width:100%;height:100%;overflow:hidden;');
  iframe.setAttribute('frameBorder', '0');
  const iframeDone = function (err) {
    clearPageLoader();
    const data = {
      guid: window.Asc.plugin.guid,
      id: msg.id,
      success: !err
    };
    if (err) {
      data.error = {
        message: err.message || err.toString()
      }
    }
    window.Asc.scope.data = data;
    window.Asc.plugin.callCommand(function () {
      window.parent.postMessage(window.JSON.stringify({
        frameEditorId: window.frameEditorId,
        event: 'onCommandReturn',
        data: Asc.scope.data
      }), '*');
    });
  };
  iframe.onload = function () {
    iframe.onload = null;
    iframe.onerror = null;
    iframeDone();
  };
  iframe.onerror = function (err) {
    iframe.onload = null;
    iframe.onerror = null;
    iframeDone(err);
  };
  if (typeof msg.iframe === 'string') {
    if (count > 0 || msg.template) {
      return loadResource(msg.iframe).then(html => {
        return html.replace(/(<\/body>)/, $0 => {
          return htmls.join('\n') + $0;
        });
      }).then(html => {
        return 'data:text/html;charset=utf-8,' + encodeURIComponent(html);
      }).then(src => {
        iframe.src = src;
        document.getElementById('container').appendChild(iframe);
      });
    } else {
      iframe.src = msg.iframe;
    }
  } else {
    iframe.src = 'data:text/html;charset=utf-8,' + encodeURIComponent(htmls.join('\n'));
  }
  document.getElementById('container').appendChild(iframe);
}

function clearPageLoader () {
  document.getElementById('container').style.opacity = 1;
  let loader = document.getElementById('view-loader');
  if (loader) {
    loader.parentNode.removeChild(loader);
  }
}

function fillAttrs (el, item) {
  Object.keys(item).forEach(attr => {
    if (attr !== 'head') {
      if (/^html$/i.test(attr)) {
        el.innerHTML = item[attr];
      } else {
        el.setAttribute(attr, item[attr]);
      }
    }
  });
}

async function createScripts (script, head) {
  if (!script) return;
  const container = head ? document.head : document.getElementById('container');
  let waiting = null;
  for (let item of script) {
    if ((!!item.head) !== head) continue;
    let el = document.createElement('script');
    el.type = 'text/javascript';
    if (item.async === 'async') { // 异步加载,不受顺序控制
      fillAttrs(el, item);
    } else { // 加载外部脚本
      if (waiting) {
        await waiting;
        waiting = null;
      }
      if (item.src) {
        waiting = (function (el, src) {
          window.debug('[%s] script loading =>', window.pluginName, src);
          return new Promise((resolve, reject) => {
            el.onerror = function (err) {
              window.debug('[%s] script error =>', window.pluginName, src, err);
              reject(err);
            };
            el.onload = function () {
              window.debug('[%s] script injected =>', window.pluginName, src);
              resolve(el);
            }
          })
        })(el, item.src);
        fillAttrs(el, item);
      } else if (typeof item === 'string') {
        el.innerHTML = item;
      } else {
        fillAttrs(el, item);
      }
    }
    container.appendChild(el);
  }
  if (waiting) {
    await waiting;
  }
}

function loadScript (item, el) {
  const src = item.src;
  delete item.src;
  window.debug('[%s] script loading =>', window.pluginName, src);
  let xhr = new XMLHttpRequest();
  xhr.open('get', src, true);
  xhr.setRequestHeader('Content-Type', 'text/plain');
  fillAttrs(el, item);
  return new Promise((resolve, reject) => {
    xhr.onerror = reject;
    xhr.onload = () => {
      window.debug('[%s] script loaded =>', window.pluginName, src, xhr.responseText.length);
      return resolve({
        el,
        html: xhr.responseText,
        src
      });
    };
    xhr.send();
  });
}

async function loadScripts (script, head) {
  if (!script) return;
  const container = head ? document.head : document.getElementById('container');
  let tasks = [];
  for (let i = 0; i < script.length; i++) {
    const item = script[i];
    if ((!!item.head) !== head) continue;
    let el = document.createElement('script');
    el.type = 'text/javascript';
    container.appendChild(el);
    if (item.async === 'async') { // 异步加载,不受顺序控制
      fillAttrs(el, item);
    } else if (item.src) { // 同步加载外部脚本
      tasks.push(loadScript(item, el));
    } else if (typeof item === 'string') {
      tasks.push(Promise.resolve({
        el, html: item, src: `script[${i}]`
      }));
    } else {
      const html = item.html;
      delete item.html;
      fillAttrs(el, item);
      tasks.push(Promise.resolve({
        el, html, src: `script[${i}]`
      }));
    }
  }
  await Promise.all(tasks).then(pairs => {
    for (let p of pairs) {
      window.debug('[%s] script injected =>', window.pluginName, p.src);
      p.el.innerHTML = p.html;
    }
  });
}

async function executeInContext (msg) {
  if (window.onload) {
    window.debug('[%s] executeInContext reset onload =>', window.pluginName, msg.id, window.onload);
  }
  window.onload = null; // 当前文档的onload已经触发过了，直接置为空
  const container = document.getElementById('container');
  const errorHandler = document.body.onerror;
  let done = false;
  const promises = [];
  const executeDone = function (err) {
    clearPageLoader();
    if (err) {
      console.error('[%s] executeInContext error =>', window.pluginName, msg, err);
      window.parent.parent.postMessage(window.JSON.stringify({
        frameEditorId: window.frameEditorId,
        event: 'onError',
        data: {
          errorCode: err.code || 'VIEW_RENDER_ERROR',
          errorDescription: err.message || err.toString()
        }
      }), '*');
    } else {
      window.debug('[%s] executeInContext done =>', window.pluginName, msg.id);
    }
    if (done) return;
    done = true;
    document.body.onerror = errorHandler;
    let data = {
      id: msg.id,
      guid: window.Asc.plugin.guid,
      success: true
    };
    if (err) {
      data.success = false;
      data.error = {
        message: err.message || err.toString()
      };
    }
    window.parent.parent.postMessage(window.JSON.stringify({
      frameEditorId: window.frameEditorId,
      event: 'onCommandReturn',
      data
    }), '*');
  };
  document.body.onerror = function (err) {
    executeDone(err);
  }
  window.addEventListener('error', errorHandler, true);
  const scriptLoader = window.XMLHttpRequest && msg.xhr !== false ? loadScripts : createScripts;
  let el = document.createElement('script');
  try {
    if (msg.link) {
      for (let item of msg.link) {
        el = document.createElement('link');
        const async = item.async === 'async';
        delete item.async;
        if (!item.rel && /.css$/.test(item.href)) {
          item.rel = 'stylesheet';
        }
        Object.keys(item).forEach(attr => {
          el.setAttribute(attr, item[attr]);
        });
        if (el.href && !async) {
          promises.push(new Promise((resolve, reject) => {
            (function (el) {
              el.onerror = function (err) {
                window.debug('[%s] link error =>', window.pluginName, el.href, err);
                reject(err);
              };
              el.onload = function () {
                window.debug('[%s] link loaded =>', window.pluginName, el.href);
                resolve();
              }
            })(el);
          }));
        }
        document.head.appendChild(el);
      }
    }
    if (msg.style) {
      for (let item of msg.style) {
        el = document.createElement('style');
        if (typeof item === 'string') {
          el.innerHTML = item;
        } else {
          Object.keys(item).forEach(attr => {
            if (/^html$/im.test(attr)) {
              el.innerHTML = item[attr];
            } else {
              el.setAttribute(attr, item[attr]);
            }
          });
        }
        document.head.appendChild(el);
      }
    }
    await scriptLoader(msg.script, true);
    let htmls = [
      'var services = {};',
    ];
    const methods = msg.methods || {};
    Object.keys(methods).forEach(name => {
      let fn = methods[name];
      if (/^\s*(?:(?!function)\w)+\s*\(/im.test(fn)) {
        fn = 'function ' + fn;
      }
      htmls.push(`window['${name}']=${fn};`);
    });
    const services = msg.services || {};
    Object.keys(services).forEach(name => {
      htmls.push(`services['${name}']= function (){
        let id = ++window.reqId;
        let req = window.requests[id] = {id};
        req.promise = new Promise((resolve, reject)=>{
          req.resolve = resolve;
          req.reject = reject;
        });
        const name = '${name}';
        const args = [].slice.call(arguments)
        window.debug('[%s] onPluginAction.service =>', window.pluginName, id, name, args);
        window.parent.parent.postMessage(JSON.stringify({
          frameEditorId: window.frameEditorId,
          event: 'onPluginAction',
          data: {
            guid: window.Asc.plugin.guid,
            id,
            type: 'service',
            data: {
              name,
              args
            }
          }
        }), '*');
        return req.promise;
      }`);
    });
    el = document.createElement('script');
    el.innerHTML = htmls.join('\n');
    window.debug('[%s] scripts generated =>', window.pluginName, el.innerHTML);
    document.body.appendChild(el);
    el = document.createElement('div');
    el.innerHTML = msg.template;
    container.appendChild(el);
    await scriptLoader(msg.script, false);
    window.debug('[%s] executeInContext waiting =>', window.pluginName, msg.id);
    Promise.all(promises).then(() => {
      window.debug('[%s] document loaded =>', window.pluginName, msg.id);
      try {
        window.dispatchEvent(new Event('load', {
          cancelable: false,
          bubbles: false
        }));
        executeDone();
      } catch (err) {
        executeDone(err);
        console.error('dispatchEvent load error =>', err);
      }
    }).catch(err => {
      executeDone(err);
    });
  } catch (err) {
    executeDone(err);
  }
}

function executeCommand (command, close) {
  const id = ++window.reqId;
  const req = window.requests[id] = {};
  const guid = window.Asc.plugin.guid;
  window.debug('[%s] executeCommand =>', window.pluginName, id, command);
  let cmd = `try{
    window.postMessage(window.JSON.stringify({
      guid: '${guid}',
      type: 'onExternalPluginMessage',
      data: {
        id: ${id},
        type: 'onCommandReturn',
        data: {
          success: true,
          data: (${command})
        }
      }
    }), "*");
  }catch(e){
    window.postMessage(window.JSON.stringify({
      guid: '${guid}',
      type: 'onExternalPluginMessage',
      data: {
        id: ${id},
        type: 'onCommandReturn',
        data: {
          success: false,
          error: {
            message: e.message
          }
        }
      }
    }), "*");
  }`
  req.promise = new Promise((resolve, reject) => {
    req.resolve = resolve;
    req.reject = reject;
    window.Asc.plugin.executeCommand(close ? 'close' : 'command', cmd);
  });
  return req.promise;
}

/**
 * 执行Excel操作命令
 * @function executeCommand
 * @param {String} command Excel操作命令，具体命令插件{@link https://api.onlyoffice.com/docbuilder/spreadsheetapi SpreadSheetApi}
 */
window.executeCommand = function (command) {
  return executeCommand(command, false);
}

/**
 * 执行Excel操作并关闭插件
 * @function closeCommand
 * @param {String} command Excel操作命令，具体命令插件{@link https://api.onlyoffice.com/docbuilder/spreadsheetapi SpreadSheetApi}
 */
window.closeCommand = function (command) {
  return executeCommand(command, true);
};

window.Asc.plugin.onExternalPluginMessage = function (msg) {
  if (!msg) return;
  if (msg.type === 'onActiveSheetChanged') {
    /**
     * 视图插件中监听Sheet活动页变化事件，只能通过<code>window.onActiveSheetChanged</code>或
     * <code>window.onactivesheetchanged</code>监听
     * @event module:plugins/View#ActiveSheetChanged
     * @type {Object}
     * @property {Object} data 数据
     * @property {String} data.sheets 当前编辑器打开的文档的所有Sheet页名称
     * @property {String} data.active 当前编辑器活动的Sheet页名称
     * @example
     * window.onActiveSheetChanged = function (e) {
     *   console.log('ActiveSheetChanged =>', e.data.active);
     * }
     */
    if (window.onActiveSheetChanged) {
      window.onActiveSheetChanged(msg);
    } else if (window.onactivesheetchanged) {
      window.onactivesheetchanged(msg);
    }
  } else if (msg.type === 'onSheetsChanged') {
    /**
     * 视图插件中监听Sheet页变化事件，只能通过<code>window.onSheetsChanged</code>或
     * <code>window.onsheetschanged</code>监听
     * @event module:plugins/View#SheetsChanged
     * @type {Object}
     * @property {Object} data 数据
     * @property {String} data.sheets 当前编辑器打开的文档的所有Sheet页名称
     * @property {String} data.active 当前编辑器活动的Sheet页名称
     * @example
     * window.onSheetsChanged = function (e) {
     *   console.log('SheetsChanged =>', e.data.active);
     * }
     */
    if (window.onSheetsChanged) {
      window.onSheetsChanged(msg);
    } else if (window.onsheetschanged) {
      window.onsheetschanged(msg);
    }
  } else if (msg.type === 'executeRender') {
    window.debug('[%s] executeRender =>', window.pluginName, msg);
    // 在iframe执行渲染页面操作
    if (msg.iframe) {
      executeInIframe(msg);
    } else {
      // 在当前页面执行渲染
      executeInContext(msg);
    }
  } else if (msg.type === 'onActionReturn') {
    window.debug('[%s] onActionReturn =>', window.pluginName, msg.id, msg.data);
    const req = window.requests[msg.id];
    delete window.requests[msg.id];
    if (req) {
      if (req.iframe) {
        const iframe = document.getElementById('view');
        iframe.contentWindow.postMessage({ id: req.data.id, data: msg.data }, '*');
      } else if (msg.data.success) {
        req.resolve(msg.data.data);
      } else {
        req.reject(msg.data.error);
      }
    }
  } else if (msg.type === 'onPluginReady') {
    // 接受编辑器发来的的初始化完毕消息，并保存编辑器的Id，用来发送消息
    window.frameEditorId = msg.frameEditorId;
    // 给编辑器的父窗口发送编辑器的插件初始化完成消息
    window.parent.parent.postMessage(JSON.stringify({
      frameEditorId: window.frameEditorId,
      event: msg.type,
      data: {
        name: 'view',
        debug: window.localStorage.debug,
        guid: window.Asc.plugin.guid,
        sheets: msg.sheets,
        active: msg.active
      }
    }), '*');
  } else if (msg.type === 'onCommandReturn') {
    window.debug('[%s] onCommandReturn =>', window.pluginName, msg.id, msg.data);
    const req = window.requests[msg.id];
    delete window.requests[msg.id];
    if (req) {
      if (msg.data.success) {
        req.resolve(msg.data.data);
      } else {
        req.reject(msg.data.error);
      }
    }
  } else if (msg.type === 'invoke') {
    window.debug('[%s] invoke =>', window.pluginName, msg);
    const args = msg.args || [];
    const invokeDone = (err, data) => {
      if (err) {
        data = {
          success: false,
          error: {
            message: err.message || err.toString()
          }
        }
      } else {
        data = {
          success: true,
          data
        }
      }
      data.id = msg.id;
      data.guid = window.Asc.plugin.guid;
      window.parent.parent.postMessage(window.JSON.stringify({
        frameEditorId: window.frameEditorId,
        event: 'onCommandReturn',
        data
      }), '*');
    }
    const fn = window[msg.name];
    if (!fn) {
      return invokeDone(new Error(`[${msg.name}] method not exists`));
    }
    Promise.resolve(window[msg.name](...args)).then(data => {
      invokeDone(null, data);
    }).catch(err => {
      invokeDone(err);
    });
  } else if (msg.type === 'setDebug') {
    console.log('[%s] reset debug =>', window.pluginName, msg.debug);
    window.localStorage.debug = msg.debug;
    window.debug = debugFactory();
  }
}

window.Asc.plugin.init = () => {
  // 初始化
  window.Asc.scope.guid = window.Asc.plugin.guid;
  // 在编辑器窗口执行命令
  window.Asc.plugin.callCommand(function () {
    let guid = Asc.scope.guid;
    // 监控插件关闭事件
    window.g_asc_plugins.api.asc_registerCallback('asc_onPluginClose', plugin => {
      if (plugin.get_Guid() === guid) {
        // 发送当前插件关闭消息
        window.parent.postMessage(window.JSON.stringify({
          frameEditorId: window.frameEditorId,
          event: 'onPluginClose',
          data: {
            name: 'view',
            guid
          }
        }), '*');
      }
    });
    window.g_asc_plugins.api.asc_registerCallback('asc_onSheetsChanged', () => {
      window.parent.postMessage(window.JSON.stringify({
        guid,
        type: 'onExternalPluginMessage',
        data: {
          type: 'onSheetsChanged',
          data: {
            sheets: Api.GetSheets().map(i => i.Name),
            active: Api.GetActiveSheet().GetName()
          }
        }
      }), '*');
    });
    window.g_asc_plugins.api.asc_registerCallback('asc_onActiveSheetChanged', () => {
      window.parent.postMessage(window.JSON.stringify({
        guid,
        type: 'onExternalPluginMessage',
        data: {
          type: 'onActiveSheetChanged',
          data: {
            sheets: Api.GetSheets().map(i => i.Name),
            active: Api.GetActiveSheet().GetName()
          }
        }
      }), '*');
    });
    // 获取当前sheet页信息，并发送给当前插件窗口
    let sheets = Api.GetSheets();
    window.postMessage(window.JSON.stringify({
      guid,
      type: 'onExternalPluginMessage',
      data: {
        type: 'onPluginReady',
        frameEditorId: window.frameEditorId,
        sheets: sheets.map(i => i.Name),
        active: Api.GetActiveSheet().GetName()
      }
    }), '*');
  })
}