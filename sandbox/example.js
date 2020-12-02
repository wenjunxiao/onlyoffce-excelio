/**
 * 对应于需要集成office的CMS系统的office编辑页面
 */
import _ from 'lodash';
import {
  getPublicUrl,
  postService
} from './app.services';
import { debug } from '../src/plugins/utils';

export * from './base';

const examples = {
  'ExcelIO-Basic': {
    code: `(function(name){
      let writer = Excel.createWriter({});
      const titleOpts = {
        // font: {
        //   bold: true,
        //   name: '微软雅黑',
        //   sz: 10
        // }
      };
      writer.withoutGridLines().sheet('Basic')
        .row()
        .cell('A', titleOpts).width(1)
        .cell('AB', titleOpts).width(2)
        .cell('ABC', titleOpts).width(3)
        .cell('ABCD', titleOpts).width(4)
        .cell('ABCDE', titleOpts).width(5)
        .cell('ABCDEF', titleOpts).width(6)
        .cell('ABCDEFG', titleOpts).width(7)
        .cell('ABCDEFGH', titleOpts).width(8)
        .cell('ABCDEFGHI', titleOpts).width(9)
        .cell('ABCDEFGHIJ', titleOpts).width(10)
        .cell('ABCDEFGHIJKLMNO', titleOpts).width(15)
        .cell('ABCDEFGHIJKLMNOPQRST', titleOpts).width(20)
        .cell('ABCDEFGHIJKLMNOPQRSTUVWXYZ', titleOpts).width(26)
        .cell('中文', titleOpts).chWidth(2)
        .cell('中文字', titleOpts).chWidth(3)
        .cell('自动计算宽度', titleOpts)
        .cell('Auto Width', titleOpts)
      ;
      writer.border2end(0, 0);
      return writer.build({});
    })(name)`,
  },
  'ExcelIO-Simple': {
    code: `(function(name){
    let data = \${JSON.stringify(data)}
    let writer = Excel.createWriter({});
    writer.newSheet(name).row().cell('Title1').cell('Title2').cell('Title3');
    for(let d of data) {
      writer.row().cell(d.v1).date(d.v2, 'YYYY-MM-DD HH:mm:ss').currency(_.round(d.v3, 2), '$')
    }
    return writer.build({});
  })(name)`,
    data: `[{v1: 'String', v2: new Date(), v3: 3.14159}]`
  },
  'ExcelIO-Full': {
    code: `(function(name){
    let writer = new ExcelWriter({
      NaN: '-'
    });
    writer.withoutGridLines()
    const titleOpts = {
      font: {
        bold: true,
        name: '微软雅黑',
        sz: 10
      }
    };
    function fillSheet(sheet, data, titleOpts){
      sheet.row().cell('Title1', titleOpts).width(8)
        .cell('Title2', titleOpts)
        .cell('Title3', titleOpts);
      for (let d of data) {
        sheet.row().cell(d.v1).color('#4cf18f' , '#f00000').currency(d.v2, '$').number(d.v3)
      }
      sheet.border2end(0, 0, '#000000');
    }
    const data = {{data}};
    fillSheet(writer.active().clear(), data, titleOpts);
    // fillSheet(writer.newSheet(name), data, titleOpts);
    // fillSheet(writer.newSheet(name).rename(name + '-rename'), data, titleOpts);
    // fillSheet(writer.active().rename('active-rename'), data, titleOpts);
    return writer.build();})('{{sheet}}')`,
    data: `${JSON.stringify([{
      v1: 11,
      v2: 12,
      v3: 13
    },
    {
      v1: 21,
      v2: 22,
      v3: 23
    }
    ], null, 2)}`
  },
  'ExcelIO-Border': {
    code: `(function(name){
    let writer = new ExcelWriter({
      NaN: '-'
    });
    writer.withoutGridLines()
    const titleOpts = {
      font: {
        bold: true,
        name: '微软雅黑',
        sz: 10
      }
    };
    function fillSheet(sheet, data, titleOpts){
      sheet.row().row(1).cell('Title1', titleOpts).width(8)
        .cell('Title2', titleOpts)
        .cell('Title3', titleOpts)
        .cell('Title4', titleOpts);
      for (let d of data) {
        sheet.row(1).cell(d.v1).color('#4cf18f' , '#f00000').currency(d.v2, '$')
          .number(d.v3).cell(d.v4)
      }
      sheet.border2end(1, 1, '000000', 'thin', {
        // outer: true,
        inner: true
      });
    }
    const data = {{data}};
    fillSheet(writer.active().clear(), data, titleOpts);
    // fillSheet(writer.newSheet(name), data, titleOpts);
    // fillSheet(writer.newSheet(name).rename(name + '-rename'), data, titleOpts);
    // fillSheet(writer.active().rename('active-rename'), data, titleOpts);
    return writer.build();})('{{sheet}}')`,
    data: `${JSON.stringify([{
      v1: 11,
      v2: 12,
      v3: 13,
      v4: 's1'
    },
    {
      v1: 21,
      v2: 22,
      v3: 23,
      v4: 's2'
    },
    {
      v1: 31,
      v2: 32,
      v3: 33,
      v4: 's3'
    }
    ], null, 2)}`
  }
};

export function mounted () {
  const url = getPublicUrl('sandbox/example.xlsx?hash');
  console.log('open url =>', url, location.query);
  let docEditor = ExcelAPI.createEditor('placeholder', {
    type: location.query.type || 'desktop',
    document: {
      version: location.query.version || '',
      title: location.query.title || 'example.xlsx',
      url
    },
    editorConfig: {
      lang: location.query.lang || 'zh',
      mode: location.query.mode || 'edit',
      user: {
        id: "excelio",
        name: "Mr ExcelIO"
      }
    }
  });
  docEditor.ready(async (editor) => {
    const sheets = await docEditor.getSheetNames();
    console.log('ready sheets =>', sheets, docEditor.readOnly, editor === docEditor);
  });
  docEditor.on('PluginOpen', (e) => {
    debug('[Example] PluginOpen =>', e);
    const services = {
      commit (data) {
        return postService(data.service, data).then(rsp => {
          return docEditor.executeCode(rsp.code, { _, name: rsp.sheet }).then(result => {
            return '执行成功:' + (result && JSON.stringify(result) || '');
          }).catch(err => {
            console.error('执行失败 =>', err);
            return '执行失败:' + err.message;
          });
        });
      },
      async commit0 (data) {
        console.log('[commit0] called => ', data, Date.now());
      },
      commit1: (data) => {
        console.log('[commit1] called => ', data, Date.now());
      },
      commit2: data => {
        console.log('[commit2] called => ', data, Date.now());
      },
      commit3: async (data) => {
        console.log('[commit3] called => ', data, Date.now());
      },
      commit4: function (data) {
        console.log('[commit4] called => ', data, Date.now());
      },
      commit5: async function commit (data) {
        console.log('[commit5] called => ', data, Date.now());
      },
      commit6: that => console.log('[commit6] called => ', that, Date.now()),
      commit7: async (that) => console.log('[commit7] called => ', that, Date.now())
    };
    docEditor.renderPlugin({
      methods: {
        doCommit () {
          let sheet = document.getElementById('sheet').value;
          let service = document.getElementById('service').value;
          let code = document.getElementById('code').value || '';
          let data = eval(document.getElementById('data').value || '');
          this[service]({
            service,
            sheet,
            code: code.replace(/\$\{\s*JSON\.stringify\(\s*(\w+)\s*\)\s*\}/img, '{{$1}}'),
            data
          });
        },
        commit (data) {
          console.log('before commit =>', this.services);
          window.services.commit(data).then(code => {
            document.getElementById('result').value = code;
          });
          console.log('after commit');
        },
        async commit0 (data) {
          console.log('before commit0');
          await services.commit0(data);
          console.log('after commit0');
        },
        commit1: (data) => {
          console.log('before commit1');
          services.commit1(data);
          console.log('after commit1');
        },
        commit2: data => {
          console.log('before commit2');
          services.commit2(data);
          console.log('after commit2');
        },
        commit3: async (data) => {
          console.log('before commit3');
          await services.commit3(data);
          console.log('after commit3');
        },
        commit4: function (data) {
          console.log('before commit4');
          services.commit4(data);
          console.log('after commit4');
        },
        commit5: async function commit (data) {
          console.log('before commit5');
          await services.commit5(data);
          console.log('after commit5');
        },
        commit6: data => services.commit6(data),
        commit7: async (data) => services.commit7(data),
        chgExample (that) {
          const sel = examples[that.value];
          document.getElementById('code').value = sel.code;
          document.getElementById('data').value = sel.data;
        }
      },
      services,
      link: [],
      script: [
        // {src: ''},
        `var examples=${JSON.stringify(examples)};
        chgExample(document.getElementById('example'))`
      ],
      style: [],
      template: `<div>
      Sheet:<select id="sheet">
      ${e.data.sheets.map(sheet => `<option value="${sheet}">${sheet}</option>`)}
      </select>
      Service:<select id="service">
      ${Object.keys(services).map(name => `<option value="${name}">${name}</option>`)}
      </select><br/>
      Example:<select id="example" onchange="chgExample(this)">
      ${Object.keys(examples).map(name => `<option value="${name}">${name}</option>`)}
      </select><button onclick="doCommit(this)">执行</button><br/>
      <textarea id="code" style="width:98%;" rows="15"></textarea><br/>
      <textarea id="data" style="width:98%;" rows="8"></textarea><br/>
      <textarea id="result" style="width:98%;" rows="8"></textarea>
      </div>`
    });
  });
  window.docEditor = docEditor;
};