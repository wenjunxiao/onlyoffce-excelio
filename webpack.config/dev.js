/* eslint-disable */
const os = require('os');
const net = require('net');
const dns = require('dns');
const http = require('http');
const fs = require('fs');
const path = require('path');
const webpack = require('webpack');
const UglifyJS = require('uglify-es');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const ExcelIO = require('node-excelio');
const urlParse = require('url').parse;

const base = require('./base');
const BASE_DIR = path.dirname(__dirname);

const minify = (code, bare_returns = false) => {
  let result = UglifyJS.minify(code, {
    compress: true,
    toplevel: true,
    parse: {
      bare_returns
    }
  });
  if (result.error) throw result.error;
  return result.code;
}

const argv = process.argv.slice(1);
function getArg (name, defVal = null) {
  let pos = argv.lastIndexOf(`--${name}`);
  if (pos < 0) return defVal;
  let v = argv[pos + 1];
  if (v === undefined && typeof defVal === 'boolean') return true;
  return v;
}

const contentBase = [];
let idx = 0;
while ((idx = argv.indexOf('--base', idx)) > 0) {
  if (argv[idx + 1]) {
    contentBase.push(argv[idx + 1]);
  }
  idx = idx + 1;
}
const onlyoffice = getArg('onlyoffice', process.env.ONLYOFFICE_CDN);
const getAddress = () => {
  return new Promise(resolve => {
    let url = urlParse(onlyoffice);
    if (net.isIP(url.hostname) === 0) {
      dns.resolve4(url.hostname, (err, address => {
        resolve(address[0]);
      }));
    } else {
      resolve(url.hostname);
    }
  })
}
/**
 * 匹配度越高得分越高
 * @param {Array} addr 匹配的IP
 * @param {Array} base 基准IP
 */
const addrScore = (addr, base) => {
  for (let i = 0; i < addr.length; i++) {
    if (addr[i] !== base[i]) {
      return i;
    }
  }
  return addr.length;
}

/**
 * 返回两个IP，第一个IP用于本地开发环境访问，默认选择局域网IP；
 * 第二个IP选择onlyoffice容器可以访问的IP，选择公网或者VPN的IP：
 * 1.和onlyoffice服务网段匹配度越高表示容器能访问到的可能性越大
 * 2.选择电脑的局域网IP最有可能被容器访问
 */
const getHost = () => {
  return new Promise(resolve => {
    const host = getArg('host')
    if (host) {
      return resolve([host, host]);
    } else {
      getAddress().then(address => {
        const addrs = address.split('.');
        let ips = [];
        let ifaces = os.networkInterfaces();
        Object.keys(ifaces).forEach(ifname => {
          let alias = 0;
          ifaces[ifname].forEach(iface => {
            if ('IPv4' !== iface.family || iface.internal !== false) {
              return;
            }
            if (alias === 0) {
              ips.push({
                name: ifname,
                ip: iface.address,
                score: addrScore(iface.address.split('.'), addrs),
                netmask: iface.netmask,
                iface
              })
            }
            ++alias;
          });
        });
        if (ips.length > 0) {
          const host = (ips.filter(ip => !(/.255$/.test(ip.netmask)))[0] || ips[0]).ip
          ips.sort((a, b) => {
            return b.score - a.score;
          });
          if (ips[0].score > 0) {
            return resolve([host, ips[0].ip]);
          } else {
            return resolve([host, (ips.filter(ip => !(/(vbox|docker)/.test(ip.name)))[0] || ips[0]).ip]);
          }
        } else {
          return resolve(['localhost', 'localhost']);
        }
      });
    }
  })
}

function toArray (s) {
  if (Array.isArray(s)) return s;
  else if (s) return s.split(',');
  return [];
}

const PUBLIC_PATH = process.env.PUBLIC_PATH = '/onlyoffice-excelio/';
// dev 才build其他模块，否则只build当前模块，其他模块通过`npm run build`命令构建之后静态代理
const isDev = /dev/.test(getArg('env'));
const UPLOAD_BASE = getArg('upload-base', path.resolve(BASE_DIR, '.upload'));
module.exports = getHost().then(([host, publicHost]) => {
  const PLUGIN_URL = getArg('plugin-url', PUBLIC_PATH + 'plugins/api.js');
  console.error('start info(%s) =>', isDev, host, publicHost, PLUGIN_URL);
  const devServer = {
    publicPath: PUBLIC_PATH,
    contentBase: contentBase.concat([
      path.join(BASE_DIR, '.')
    ]),
    host: host,
    port: argv.indexOf('--port') > 0 ? argv[argv.indexOf('--port') + 1] : 0,
    open: true,
    historyApiFallback: true,
    inline: true,
    disableHostCheck: true,
    progress: true,
    headers: {
      'Access-Control-Allow-Origin': '*',
    },
    before (app) {
      const PNG = {};
      /**
       * 跳转到首页
       */
      app.get('/', (_, res) => res.redirect(PUBLIC_PATH));
      app.use(function (req, res, next) {
        console.error('[%s] %s %s %j', req.ip, req.method, req.url, req.headers['referer'] || '');
        res.setHeader('Access-Control-Allow-Origin', '*');
        next();
      });
      /**
       * 提交插件的icon图片
       */
      app.post(PUBLIC_PATH + 'plugins/:id/:name.png', (req, res) => {
        const ids = req.params.id.split(',');
        const alias = toArray(req.query.alias);
        const chunks = [];
        req.on('data', chunk => {
          chunks.push(chunk);
        });
        req.on('end', () => {
          let img = Buffer.concat(chunks).toString('binary');
          if (/^data:\w+\/\w+;base64,([\s\S]*)$/.test(img)) {
            img = Buffer.from(RegExp.$1, 'base64');
          }
          ids.forEach(id => {
            PNG[id + '.' + req.params.name] = img;
            alias.forEach(name => {
              PNG[id + '.' + name] = img;
            });
          });
          res.end('ok');
        });
      });
      /**
       * 插件的icon图片访问地址
       */
      if (['true', true, '1', 1].includes(process.env.DYNAMIC_ICON)) {
        app.get(PUBLIC_PATH + 'plugins/:id/:name.png', (req, res) => {
          const png = PNG[req.params.id + '.' + req.params.name];
          if (!png) {
            try {
              let pathname = urlParse(req.url).pathname.substr(PUBLIC_PATH.length);
              const stream = fs.createReadStream(path.resolve(BASE_DIR, './dist/' + pathname));
              return stream.on('error', err => {
                res.status(404);
                res.end(err.message);
              }).pipe(res);
            } catch (err) {
              res.status(404);
              return res.end(err.message);
            }
          }
          res.writeHead(200, {
            'Content-Type': 'image/png'
          });
          return res.end(png);
        });
      }
      if (!isDev) {
        /**
         * 非开发环境，读取build好的插件文件
         */
        app.get(PUBLIC_PATH + 'plugins/*', (req, res) => {
          let pathname = urlParse(req.url).pathname.substr(PUBLIC_PATH.length);
          try {
            const stream = fs.createReadStream(path.resolve(BASE_DIR, './dist/' + pathname));
            stream.on('error', err => {
              res.status(404);
              res.end(err.message);
            }).pipe(res);
          } catch (err) {
            res.status(404);
            res.end(err.message);
          }
        });
      }
      /**
       * 读取说明文档
       */
      app.get(PUBLIC_PATH + 'doc', (_, res) => res.redirect(PUBLIC_PATH + 'doc/index.html'));
      app.get(PUBLIC_PATH + 'doc/*', (req, res) => {
        let pathname = urlParse(req.url).pathname.substr(PUBLIC_PATH.length);
        if (pathname === 'doc/') {
          return res.redirect(PUBLIC_PATH + 'doc/index.html');
        }
        try {
          if (/\.(css|js)$/.test(pathname)) {
            const maps = {
              'css': 'text/css',
              'js': 'text/javascript'
            }
            res.setHeader('Content-Type', `${maps[RegExp.$1]};charset=utf-8;`)
          }
          const stream = fs.createReadStream(path.resolve(BASE_DIR, './dist/' + pathname));
          stream.on('error', err => {
            res.status(404);
            res.end(err.message);
          }).pipe(res);
        } catch (err) {
          res.status(404);
          res.end(err.message);
        }
      });
      app.get(PUBLIC_PATH + 'upload', (req, res) => {
        fs.readdir(UPLOAD_BASE, (err, files) => {
          if (err) {
            res.status(404);
            return res.end(err.message);
          }
          res.writeHead(200, {
            'Content-Type': 'application/json;charset=utf-8'
          });
          return res.end(Buffer.from(JSON.stringify(files), 'utf-8'));
        });
      });
      /**
       * 访问上传的文件
       */
      app.get(PUBLIC_PATH + 'upload/:filename', (req, res) => {
        const end = buffer => {
          let image = PNG['watermark.icon'];
          if (image) {
            buffer = ExcelIO.watermark(buffer, image);
          }
          res.end(buffer);
        };
        fs.readFile(path.resolve(UPLOAD_BASE, req.params.filename), (err, buffer) => {
          if (err) {
            if (req.query.default) {
              return fs.readFile(path.resolve(BASE_DIR, 'sandbox', req.query.default), (err0, buffer) => {
                if (err0) {
                  res.status(404);
                  return res.end(err.message);
                }
                return end(buffer);
              });
            } else {
              res.status(404);
              return res.end(err.message);
            }
          }
          end(buffer);
        });
      });
      /**
       * 保留旧的API访问接口
       */
      app.get('/sandbox/plugins/*', (req, res) => {
        return res.redirect(req.url.replace(/^\/sandbox\//, PUBLIC_PATH));
      });
      /**
       * 查看开发环境服务端接口
       */
      app.get('*/webpack.config/dev.js', (req, res) => {
        res.setHeader('Content-Type', 'text/plain;charset=utf-8;');
        res.end(this.before.toString());
      })
      /**
       * 查看源码
       */
      app.get('*/sandbox/*', (req, res) => {
        let url = req.url.replace(/^(.*?\/)sandbox\//, '$1');
        http.get(`http://${this.host}:${this.port}${url}`, rsp => {
          if (rsp.statusCode === 404) {
            const filename = req.url.replace(/^.*\/(sandbox\/.*?)(?:\?.*)?$/, '$1');
            res.setHeader('Content-Type', 'text/plain;charset=utf-8;');
            return fs.createReadStream(path.resolve(BASE_DIR, filename)).on('error', err => {
              res.status(404);
              res.end(err.message);
            }).pipe(res);
          }
          res.setHeader('Content-Type', 'text/plain;charset=utf-8;');
          rsp.pipe(res);
        });
      });
      /**
       * 提供给`sandbox/sample.html`的onlyoffice api 代理
       */
      app.get('*/protocol//url/of/onlyoffice/api.js', (_, res) => {
        http.get(onlyoffice, rsp => {
          rsp.pipe(res);
        });
      });
      /**
       * 提供给`sahndbox/sample.html`依赖的excel代理
       */
      app.get('*/protocol//url/of/sample.xlsx', (_, res) => {
        return res.redirect('/sandbox/example.xlsx');
      });
      /**
       * sahndbox/sample.html 依赖的js代理
       */
      app.get('*/protocol//url/of/plugins/api.js', (_, res) => {
        return res.redirect(PLUGIN_URL);
      });
      /**
       * 提供给`sample.html`的api接口
       */
      app.post('/app/api/sample', (req, res) => {
        const chunks = [];
        req.on('data', chunk => {
          chunks.push(chunk);
        });
        req.on('end', () => {
          const data = JSON.parse(Buffer.concat(chunks).toString('utf-8'));
          if (!data.statusCode) data.statusCode = 200;
          if (data.success === undefined) data.success = false;
          data.server = Date.now();
          res.writeHead(data.statusCode, {});
          res.end(JSON.stringify(data));
        });
      });
      /**
       * 模拟app的api接口
       */
      app.post('/app/api/service/:service', (req, res) => {
        const chunks = [];
        req.on('data', chunk => {
          chunks.push(chunk);
        });
        req.on('end', () => {
          const data = JSON.parse(Buffer.concat(chunks).toString('utf-8'));
          const code = data.code;
          delete data.code;
          data.code = code.trim().replace(/\{\{\s*(\w+)\s*\}\}/g, ($0, $1) => {
            const v = data[$1];
            if (v === undefined) return $0;
            if (typeof v === 'object') return JSON.stringify(v);
            return v;
          });
          if (/^(\(\s*function\s*[^{]*{)\s*([\s\S]+)(\}\)\([^);(]*\))\s*$/im.test(data.code)) {
            const prefix = RegExp.$1;
            const body = RegExp.$2;
            const suffix = RegExp.$3;
            try {
              data.code = prefix + minify(body, true) + suffix;
            } catch (err) {
              console.warn('cannot minify1 =>', body, err);
            }
          } else {
            try {
              data.code = minify(data.code);
            } catch (err) {
              if (/return\W*outside\W*of\W*function/.test(err.message)) {
                try {
                  data.code = `(function(){${minify(data.code, true)}})()`;
                } catch (err) {
                  console.warn('cannot minify2 =>', data.code, err);
                }
              } else {
                console.warn('cannot minify3 =>', data.code, err);
              }
            }
          }
          data.server = Date.now();
          res.end(Buffer.from(JSON.stringify(data), 'utf-8'));
        });
      });
      /**
       * 初始化上传文件目录
       */
      fs.exists(UPLOAD_BASE, exists => {
        if (!exists) {
          fs.mkdir(UPLOAD_BASE);
        }
      });
      /**
       * 文件上传接口
       */
      app.post('/app/api/upload', (req, res) => {
        const chunks = [];
        req.on('data', chunk => {
          chunks.push(chunk);
        });
        res.json = data => {
          return res.end(Buffer.from(JSON.stringify(data), 'utf-8'));
        };
        req.on('end', () => {
          const data = JSON.parse(Buffer.concat(chunks).toString('utf-8'));
          if (data.data) {
            const filename = data.filename;
            if (!filename) {
              res.status(400);
              return res.json({ message: 'missing filename' });
            }
            return fs.writeFile(path.resolve(UPLOAD_BASE, filename), Buffer.from(data.data, 'base64'), err => {
              if (err) {
                res.status(500);
                return res.json({ message: err.message });
              }
              res.json({
                success: true,
                data: {
                  url: `http://${this.host}:${this.port}${PUBLIC_PATH}upload/${filename}`
                }
              });
            });
          } else if (!data.url) {
            res.status(400);
            return res.json({ message: 'missing url' });
          }
          const url = urlParse(data.url);
          const filename = data.filename || 'upload@' + Date.now() + path.extname(url.pathname);
          http.get(data.url, rsp => {
            rsp.on('error', err => {
              res.status(500);
              return res.json({ message: err.message });
            });
            rsp.on('end', () => {
              res.json({
                success: true,
                data: {
                  url: `http://${this.host}:${this.port}${PUBLIC_PATH}upload/${filename}`
                }
              });
            })
            rsp.pipe(fs.createWriteStream(path.resolve(UPLOAD_BASE, filename)));
          });
        });
      });
    },
    after (app) {
      if (host !== publicHost) {
        const proxy = http.createServer(app).listen(this.port, publicHost, () => {
          console.error('proxy started =>', proxy.address());
        });
      }
    }
  };
  return [
    {
      mode: 'development',
      entry: {
        app: path.resolve(BASE_DIR, './sandbox/index.js')
      },
      devtool: 'source-map',
      output: {
        publicPath: PUBLIC_PATH,
        filename: '[name].js',
        hotUpdateChunkFilename: '[hash].hot-update.js',
      },
      devServer: process.env.BUILD_SANDBOX ? undefined : devServer,
      module: {
        rules: [{
          test: /(\.jsx|\.js)$/,
          loader: 'babel-loader',
          exclude: /(node_modules|bower_components)/,
          options: {
            presets: [
              ['@babel/preset-env', {
                targets: {
                  esmodules: true,
                }
              }]
            ]
          }
        }]
      },
      optimization: {
        minimize: true
      },
      plugins: [
        new webpack.DefinePlugin({
          'process.env': {
            PUBLIC_PATH: JSON.stringify(PUBLIC_PATH),
            PUBLIC_HOST: JSON.stringify(publicHost),
            ONLYOFFICE_URL: JSON.stringify(onlyoffice),
            PLUGIN_URL: JSON.stringify(PLUGIN_URL),
            VIEW_NAME: JSON.stringify(process.env.VIEW_NAME),
            ONLYOFFICE_WEB_APPS: JSON.stringify(onlyoffice.replace(/\/web-apps\/.*$/, '/web-apps/'))
          }
        }),
        new webpack.HotModuleReplacementPlugin({
          multiStep: false
        }),
        new HtmlWebpackPlugin({
          cache: false,
          title: 'OnlyOffice ExcelIO Sandbox',
          template: './sandbox/index.html',
          compress: true,
          minify: { //压缩HTML文件
            removeComments: false, //移除HTML中的注释
            collapseWhitespace: true //删除空白符与换行符
          }
        }),
        new HtmlWebpackPlugin({
          cache: false,
          template: './sandbox/sample.html',
          filename: 'sample.html',
          inject: true,
          chunks: [],
          onlyoffice_web_apps: onlyoffice.replace(/\/web-apps\/.*$/, '/web-apps/'),
          pluginBase: PLUGIN_URL.replace(/^(.*\/plugins\/).*$/, '$1')
        })
      ],
      resolve: {
        modules: [path.resolve('./sandbox'), 'node_modules'],
        extensions: ['.js', '.json'],
        alias: {
          'debug-factory': path.resolve(BASE_DIR, './src/debug/factory.js')
        }
      }
    }].concat(isDev ? base : []);
});