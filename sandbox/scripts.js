/**
 * 对应于需要集成office的CMS系统的office编辑页面
 */

import _ from 'lodash';
import {
  getPublicUrl,
} from './app.services';
import { debug } from '../src/plugins/utils';

export * from './base';

export function mounted () {
  let docEditor = ExcelAPI.createEditor('placeholder', {
    document: {
      title: 'scripts.xlsx',
      url: getPublicUrl('sandbox/example.xlsx?scripts')
    },
    editorConfig: {
      mode: 'edit',
      user: {
        id: "scripts",
        name: "Mr Scripts"
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
      template: `<div id="app-vue">{{ message }}
        <textarea type="text" id="data" style="width:98%;" rows="10">Api.GetSheets().map(i=>i.Name)</textarea>
        <el-button type="primary" onclick="doCommit(this)">Execute</el-button>
        <textarea type="text" id="result" style="width:98%;" rows="10"></textarea>
      </div>`,
      methods: {
        doCommit () {
          this.executeCommand(document.getElementById('data').value).then(rsp => {
            document.getElementById('result').value = '成功:' + JSON.stringify(rsp);
          }).catch(err => {
            document.getElementById('result').value = '异常:' + err.message;
          });
        }
      },
      link: [{
        href: 'https://unpkg.com/element-ui/lib/theme-chalk/index.css'
      }],
      style: [],
      script: [
        {
          head: true,
          src: 'https://unpkg.com/vue/dist/vue.js'
        },
        {
          head: true,
          src: 'https://unpkg.com/element-ui/lib/index.js'
        },
        `window.onload = function() {
          console.log('vue on load')
          new Vue({
              el: '#app-vue',
              data: {
                message: 'Hello Vue!',
                sheets: ${JSON.stringify(e.data.sheets.map(x => { return { value: x, label: x } }))}
              }
            })
          }
        `
      ],
      services
    })
  });
  window.docEditor = docEditor;
};