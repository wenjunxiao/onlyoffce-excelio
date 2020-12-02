# onlyoffice-excelio

  把[ExcelIO](https://www.npmjs.com/package/node-excelio)适配到onlyoffice，
  使得用`ExcelIO`写的后端代码可以完全移植到前端执行，复用生成excel的逻辑。

## Installation

$ npm install onlyoffice-excelio

## Usage

  本模块分为两个部分`ExcelAPI`和`ExcelIO`。其中`ExcelIO`可以直接单独引用；而`ExcelAPI`引用了`ExcelIO`。

### ExcelAPI

  通过加载插件扩展`OnlyOffice`的前端`API`。通过URL引用构建好的资源，依赖`OnlyOffice`的[api.js](https://api.onlyoffice.com/editors/basic)，可以查看[sandbox/sample.html](sandbox/sample.html)

```html
<script src="http://127.0.0.1:5180/web-apps/apps/api/documents/api.js"></script>
<script src="http://127.0.0.1:8082/onlyoffice-excelio/plugins/api.js"></script>
<script type="text/javascript">
  window.onload = function () {
    let docEditor = ExcelAPI.createEditor('placeholder', {
      document: {
        title: 'example.xlsx',
        url: 'protocol//url/of/sample.xlsx'
      },
      editorConfig: {
        mode: 'edit', // 'view' or 'edit'
        user: {
          id: "sample",
          name: "Mr Sample"
        }
      }
    });
  };
</script>
```

#### 文档下载与打印

  `OnlyOffice`下载和打印的原理是，调用接口通知服务器生成对应类型的文件，并返回一个临时文件地址，前端通过
  该地址进行下载或打印。为了对下载和打印的文档进行额外的处理，比如增加水印等，可以通过以下方式重新对原始的
  下载和打印文档临时链接进行处理。具体配置参考[WrappedConfig](global.html#WrappedConfig)
```js
let docEditor = ExcelAPI.createEditor('placeholder', {
  downloadUrl: true, // 与document.url 相同
  // downloadUrl: 'http://url/of/download', // 指定文档下载地址
  // 当不同类型文档的下载地址不同时，动态url，文档可能是XSLX、PDF或者其他选择的类型，
  // 目前默认时XSLX，可以通过downloadTypes指定，具体查看API文档的配置说明
  // downloadUrl (url) {
  //   return `http://api.app.com/wartermak?url=` + encodeURIComponent(url);
  // },
  printUrl: 'http://url/of/pdf', // 指定文档的PDF文件地址，打印时需要
  // printUrl (url) { // 文档的对应的pdf临时文件链接
  //   return `http://api.app.com/wartermak?url=` + encodeURIComponent(url);
  // }
});
```

### ExcelIO

  可以在项目中单独引用，然后执行后端相同的代码，具体可以参考[src/plugins/api/editor.js](src/plugins/api/editor.js)的`editor.executeCode`
```js
import ExcelIO from 'onlyoffice-excelio';
import _ from 'lodash';

// 从服务端加载代码
async function loadCodeFromServer() {
  const data = [{v1: 'String', v2: new Date(), v3: 3.14159}]
  return `(function(name){
    let data = ${JSON.stringify(data)}
    let writer = Excel.createWriter({});
    writer.newSheet(name).row().cell('Title1').cell('Title2').cell('Title3');
    for(let d of data) {
      writer.row().cell(d.v1).date(d.v2, 'YYYY-MM-DD HH:mm:ss').currency(_.round(d.v3, 2), '$')
    }
    return writer.build({});
  })(name)`;
}
const code = await loadCodeFromServer();
const args = {name: 'my-sheet', _};
const names = Object.keys(args);
// Excel 是别名，作为参数传入
const fn = new Function('ExcelIO', 'Excel', 'ExcelWriter', ...names, `return (${code})`);
fn(ExcelIO, ExcelIO, ExcelIO.ExcelWriter, ...names.map(name => args[name]));
```

## Development

  运行以下命令启动开发环境，可以通过`--onlyoffice`指定`onlyoffice`的`Document Server JavaScript Api`，
  具体路径可以查看[官方文档](https://api.onlyoffice.com/editors/basic)地址，默认是beta环境。
```bash
$ npm run dev [-- --onlyoffice http://url/of/onlyoffice/documents/api.js]
```
  也可以在本地使用`docker`启动`onlyoffice`服务
```bash
$ docker run -d -p 5180:80 onlyoffice/documentserver
```
  启动成功之后就可以使用本地服务了

```bash
$ npm run dev -- --onlyoffice http://127.0.0.1:5180/web-apps/apps/api/documents/api.js
```

  为了和线上访问模式保持一致，可以使用`nginx`访问`plugins/api.js`，启动之前和有修改之后都需要执行`npm run debug`构建，
  也可以使用`npm run build`构建无调试信息的版本。
```bash
$ npm run dev -- --env nginx --plugin-url http://127.0.0.1/onlyoffice-excelio/plugins/api.js
```
  以上命令中`--env nginx`是为了`webpack-dev-server`只构建`sandbox`并且除了`sandbox`之外其他有修改的时候也不会热更新，
  其他模块的修改只能通过`npm run build`/`npm run debug`来构建。

## Build

  生产编译使用`npm run build`，也可以使用`npm run debug`支持调试日志，通过在控制台设置`localStorage.debug = 'excel-plugin,excel-api'`来启用调试日志

```js
$ npm install
$ npm run build
```

## Deploy

  Build之后直接把`dist`目录的静态资源部署到Nginx即可，为了允许`onlyoffice`编辑器页面跨域访问需要设置`Access-Control-Allow-Origin`
```conf
server {
  location ^~ /onlyoffice-excelio/ {
    add_header 'Access-Control-Allow-Origin' '*';
    alias /path/of/onlyoffice-excelio/dist/;
  }
}
```
