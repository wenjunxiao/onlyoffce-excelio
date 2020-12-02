/**
 * 对应于需要集成office的CMS系统的office编辑页面
 */
import _ from 'lodash';
import {
  uploadFile,
  getPublicUrl,
} from './app.services';

export * from './base';

export function mounted () {
  let title = location.query.title || '测试中文文件名和保存.xlsx';
  if (!(/.xlsx$/i.test(title))) {
    title += '.xlsx'
  }
  let docEditor = ExcelAPI.createEditor('placeholder', {
    autoSync: false, // 禁止自动同步文件
    // downloadUrl: true, // 下载地址同文档地址
    downloadUrl (url) {
      return url;
    },
    // 打印文件地址，默认不支持打印
    // printUrl (url) {
    //   return url;
    // },
    // downloadUrl: getPublicUrl('sandbox/example.xlsx?save'),
    document: {
      // permissions: {
      //   print: false,
      // },
      // key: 'a12345678901234567890',
      version: location.query.version || '',
      title,
      id: location.query.id,
      url: getPublicUrl(process.env.PUBLIC_PATH + 'upload/' + title + '?default=example.xlsx')
    },
    editorConfig: {
      customization: {
        autosave: { 'true': true, 'false': false }[location.query.autosave]
      },
      mode: 'edit',
      user: {
        id: "save",
        name: "Mr Saver"
      }
    },
    events: {
      onDocumentStateChange (e) {
        appendLog('[onDocumentStateChange] ' + JSON.stringify(e));
      },
      onActiveSheetChanged (e) {
        console.log('onActiveSheetChanged =>', e.data.active);
      },
      onSheetsChanged (e) {
        console.log('onSheetsChanged =>', e.data.sheets);
      },
      onDownloadAs (e) { // 临时文件已经生成就绪，可以通过链接获取文件存储到其他
        appendLog('file uplaoding => ' + e.data);
        uploadFile({ url: e.data, filename: title }).then(rsp => {
          appendLog('file uplaoded => ' + JSON.stringify(rsp));
        });
      },
      onError (err) {
        console.log('onError =>', err)
      }
    }
  });
  async function appendLog (log) {
    if (docEditor.isRendered) { // 插件视图已经渲染，可以在页面上显示日志
      let len = await docEditor.methods.appendLogToView(log);
      console.log('total log length =>', len);
    } else {
      console.log(log);
    }
  }
  docEditor.on('ActiveSheetChanged', async (e) => {
    console.log('active sheet =>', e.data.active)
  });
  docEditor.on('SheetsChanged', async (e) => {
    console.log('sheets =>', e.data.active)
  });
  docEditor.on('PluginOpen', async (e) => {
    const services = {
      downloadAs () {
        docEditor.downloadAs();
      },
      executeCommand (cmd) {
        return docEditor.executeCommand(cmd);
      },
      switchSync () { // 切换存储状态
        docEditor.autoSync = !docEditor.autoSync;
      },
      uploadFile (data) {
        uploadFile(data).then(rsp => {
          appendLog('file uplaoded => ' + JSON.stringify(rsp));
        });
      }
    };
    await docEditor.renderPlugin({
      template: `<textarea type="text" id="data" style="width:98%;" rows="10">Api.GetSheets().map(i=>i.Name)</textarea>
        <button onclick="doCommit(this)">Execute</button>
        <button onclick="switchSync(this)">SwitchSync</button>
        <button onclick="syncFile(this)">Sync</button>
        <button onclick="clearLog(this)">Clear</button>
        <input type="file" onchange="uploadFile(this)" />
        <textarea type="text" id="result" style="width:98%;" rows="10"></textarea>`,
      methods: {
        uploadFile (that) {
          const services = this.services;
          let file = that.files[0];
          let reader = new FileReader();
          reader.onload = function () {
            services.uploadFile({
              data: reader.result.replace(/^data:[^;]*;base64,/, ''),
              filename: file.name
            })
          };
          reader.readAsDataURL(file);
        },
        syncFile () {
          this.services.downloadAs();
        },
        appendLogToView (text) {
          let el = document.getElementById('result');
          el.value += text + '\n';
          return el.value.length;
        },
        doCommit () {
          this.executeCommand(document.getElementById('data').value).then(rsp => {
            this.appendLogToView('成功:' + JSON.stringify(rsp));
          }).catch(err => {
            this.appendLogToView('异常:' + err.message);
          });
        },
        clearLog () {
          document.getElementById('result').value = '';
        },
        switchSync () {
          this.services.switchSync();
        }
      },
      script: [`window.onSheetsChanged = function (e) {
        console.log('window.onSheetsChanged =>', e.data.sheets);
      };
      window.onActiveSheetChanged = function (e) {
        console.log('window.onActiveSheetChanged =>', e.data.active);
      };`],
      services
    })
  });
  // 手动存储，`autoSync=false`时可以通过该方法触发存储
  // docEditor.on('saved', () => {
  //   docEditor.downloadAs();
  // });
  window.docEditor = docEditor;
};