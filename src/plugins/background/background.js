/**
 * 后台插件，运行在<code>iframe</code>中，所有的方法都在<code>iframe</code>的<code>window</code>中
 * @module plugins/Background
 */
import debugFactory from 'debug-factory';

/**
 * 调试日志，开启调试日志的方式是在<code>localStorage.debug</code>增加`excel-plugin`。
 * 比如<code>localStorage.debug +=',excel-plugin'</code>
 * @function debug
 * @param {...*} args 日志参数
 */
window.debug = debugFactory('excel-plugin');

window.pluginName = `${process.env.PLUGIN_NAME}`;
window.Asc.plugin.onExternalPluginMessage = function (msg) {
  if (msg) {
    if (msg.type === 'executeCommand') {
      window.debug('[%s] executeCommand => %s', window.pluginName, msg.command);
      let cmd = `try{
        window.parent.postMessage(window.JSON.stringify({
          frameEditorId: window.frameEditorId,
          event: 'onCommandReturn',
          data: {
            id: ${msg.id},
            success: true,
            data: (${msg.command})
          }
        }), '*');
      }catch(e){
        window.parent.postMessage(window.JSON.stringify({
          frameEditorId: window.frameEditorId,
          event: 'onCommandReturn',
          data: {
            id: ${msg.id},
            success: false,
            error: {
              message: e.message
            }
          }
        }), '*');
      }`
      window.Asc.plugin.executeCommand(msg.close ? 'close' : 'command', cmd);
    } else if (msg.type === 'download') {
      window.debug('[%s] download =>', window.pluginName, msg);
      window.Asc.scope.url = msg.url;
      window.Asc.plugin.callCommand(function () {
        window.__getFile(Asc.scope.url);
      });
    } else if (msg.type === 'print') {
      window.debug('[%s] print =>', window.pluginName, msg);
      window.Asc.scope.url = msg.url;
      window.Asc.scope.downloadType = msg.downloadType;
      window.Asc.plugin.callCommand(function () {
        window.g_asc_plugins.api.__orig_sendEvent(Asc.scope.downloadType, Asc.scope.url, function () {
        });
      });
    } else if (msg.type === 'setDebug') {
      console.log('[%s] reset debug =>', window.pluginName, msg.debug);
      window.localStorage.debug = msg.debug;
      window.debug = debugFactory();
    } else if (msg.type === 'initialize') {
      window.debug('[%s] initialize =>', window.pluginName, msg);
      if (typeof msg.debug !== 'undefined') {
        window.localStorage.debug = msg.debug;
      }
      window.Asc.scope.data = msg;
      window.Asc.scope.guid = window.Asc.plugin.guid;
      window.Asc.plugin.callCommand(function () {
        const data = Asc.scope.data;
        // 去掉下载的其他格式，只支持传入的文件类型
        const c_oAscFileType = window.Asc.c_oAscFileType;
        if (data.downloadTypes && c_oAscFileType) {
          let filter = data.downloadTypes.map(type => `:not([format=${c_oAscFileType[type]}])`).join('');
          window.$(`.btn-doc-format${filter}`).remove();
        }
        if (data.downloadUrl) { // 拦截下载地址
          if (!window.__getFile) {
            window.__getFile = window.AscCommon.getFile;
          }
          if (typeof data.downloadUrl === 'string') {
            window.AscCommon.getFile = (function (url) {
              return function () {
                return window.__getFile(url);
              };
            })(data.downloadUrl);
          } else {
            window.AscCommon.getFile = function (filePath) {
              window.parent.postMessage(window.JSON.stringify({
                frameEditorId: window.frameEditorId,
                event: 'onDownloadUrl',
                data: {
                  guid: Asc.scope.guid,
                  url: filePath
                }
              }), '*');
            };
          }
        }
        if (data.printUrl) { // 文档打印地址
          // 触发打印 window.Common.NotificationCenter.trigger('print')
          // 拦截发送`asc_onPrintUrl`事件重新生成url
          const api = window.g_asc_plugins.api;
          if (!api.__orig_sendEvent) {
            const sendEvent = api.sendEvent;
            for (let n in api) {
              if (n !== 'sendEvent' && api[n] === sendEvent) {
                api.__orig_sendEvent = api[n];
                api.__name_sendEvent = n;
                break;
              }
            }
            if (!api.__orig_sendEvent) {
              api.__orig_sendEvent = sendEvent;
              api.__name_sendEvent = 'sendEvent';
            }
          }
          if (typeof data.printUrl === 'string') {
            api[api.__name_sendEvent] = (function (url) {
              return function (downloadType) {
                if (downloadType === 'asc_onPrintUrl') {
                  const args = [url].concat([].slice.call(arguments, 1));
                  return this.__orig_sendEvent.apply(this, args);
                } else {
                  return this.__orig_sendEvent.apply(this, arguments);
                }
              }
            })(data.printUrl);
          } else {
            api[api.__name_sendEvent] = function (downloadType, url) {
              if (downloadType === 'asc_onPrintUrl') {
                window.parent.postMessage(window.JSON.stringify({
                  frameEditorId: window.frameEditorId,
                  event: 'onPrintUrl',
                  data: {
                    guid: Asc.scope.guid,
                    downloadType,
                    url
                  }
                }), '*');
              } else {
                return this.__orig_sendEvent.apply(this, arguments);
              }
            }
          }
        }
      });
    }
  }
}

window.Asc.plugin.init = () => {
  window.Asc.scope.guid = window.Asc.plugin.guid;
  window.Asc.plugin.callCommand(function () {
    let guid = Asc.scope.guid;
    window.g_asc_plugins.api.asc_registerCallback('asc_onPluginClose', plugin => {
      if (plugin.get_Guid() === guid) {
        window.parent.postMessage(window.JSON.stringify({
          frameEditorId: window.frameEditorId,
          event: 'onPluginClose',
          data: {
            name: 'background',
            guid
          }
        }), '*');
      }
    });
    window.g_asc_plugins.api.asc_registerCallback('asc_onSheetsChanged', () => {
      window.parent.postMessage(window.JSON.stringify({
        frameEditorId: window.frameEditorId,
        event: 'onSheetsChanged',
        data: {
          name: 'background',
          guid,
          sheets: Api.GetSheets().map(i => i.Name),
          active: Api.GetActiveSheet().GetName()
        }
      }), '*');
    });
    window.g_asc_plugins.api.asc_registerCallback('asc_onActiveSheetChanged', () => {
      window.parent.postMessage(window.JSON.stringify({
        frameEditorId: window.frameEditorId,
        event: 'onActiveSheetChanged',
        data: {
          name: 'background',
          guid,
          sheets: Api.GetSheets().map(i => i.Name),
          active: Api.GetActiveSheet().GetName()
        }
      }), '*');
    });
    window.g_asc_plugins.api.asc_registerCallback('asc_onDocumentModifiedChanged', isModified => {
      if (isModified) {
        window.parent.postMessage(window.JSON.stringify({
          frameEditorId: window.frameEditorId,
          event: 'onChanged',
          data: {
            name: 'background',
            guid
          }
        }), '*');
      } else {
        window.parent.postMessage(window.JSON.stringify({
          frameEditorId: window.frameEditorId,
          event: 'onSaved',
          data: {
            name: 'background',
            guid
          }
        }), '*');
      }
    });
    let sheets = Api.GetSheets();
    window.parent.postMessage(window.JSON.stringify({
      frameEditorId: window.frameEditorId,
      event: 'onPluginReady',
      data: {
        name: 'background',
        guid,
        debug: Asc.scope.debug,
        sheets: sheets.map(i => i.Name),
        active: Api.GetActiveSheet().GetName()
      }
    }), '*');
  });
}