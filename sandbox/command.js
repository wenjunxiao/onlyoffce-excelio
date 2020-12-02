/**
 * 对应于需要集成office的CMS系统的office编辑页面
 */
import {
  getPublicUrl,
} from './app.services';
import { debug } from '../src/plugins/utils';

export * from './base';

export function mounted () {
  let docEditor = ExcelAPI.createEditor('placeholder', {
    document: {
      title: 'command.xlsx',
      url: getPublicUrl('sandbox/example.xlsx?command')
    },
    editorConfig: {
      mode: 'edit',
      user: {
        id: "command",
        name: "Mr Commander"
      }
    },
  });
  docEditor.on('PluginOpen', async (e) => {
    debug('[Command] PluginOpen =>', e);
    const services = {
      executeCommand (cmd) {
        return docEditor.executeCommand(cmd);
      }
    };
    await docEditor.renderPlugin({
      template: `<textarea type="text" id="data" style="width:98%;" rows="10">Api.GetSheets().map(i=>i.Name)</textarea>
        <button onclick="doCommit(this)">Execute</button>
        <textarea type="text" id="result" style="width:98%;" rows="10"></textarea>`,
      methods: {
        doCommit () {
          this.executeCommand(document.getElementById('data').value).then(rsp => {
            document.getElementById('result').value = '成功:' + JSON.stringify(rsp);
          }).catch(err => {
            document.getElementById('result').value = '异常:' + err.message;
          });
        }
      },
      services
    })
  });
  window.docEditor = docEditor;
};