/* eslint-disable */
const _ = require('lodash');
const path = require('path');
const uuid = require('uuid');
const webpack = require('webpack');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyPlugin = require('copy-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');
const OptimizeCssAssetsPlugin = require('optimize-css-assets-webpack-plugin');
const IncludeFilePlugin = require('./plugins/include');
const pkg = require('../package.json');
const args = process.argv.slice(2);
const devMode = args.indexOf('--env') < 0 ? true : /^dev/.test(args[args.indexOf('--env') + 1]);
const BASE_DIR = path.dirname(__dirname);
const DEBUG_FACTORY = /prod/.test(process.env.NODE_ENV) ? './src/debug/factory.no.js' : './src/debug/factory.js';

function guid () {
  return `asc.{${uuid.v4().toUpperCase()}}`;
}

if (!process.env.BACKGROUND_NAME) {
  process.env.BACKGROUND_NAME = _.get(pkg, 'plugins.background.name') || 'Background';
}
if (!process.env.BACKGROUND_GUID) {
  process.env.BACKGROUND_GUID = _.get(pkg, 'plugins.background.guid') || guid();
}
if (!process.env.VIEW_NAME) {
  process.env.VIEW_NAME = _.get(pkg, 'plugins.view.name') || 'ExcelIO';
}
if (!process.env.VIEW_GUID) {
  process.env.VIEW_GUID = _.get(pkg, 'plugins.view.guid') || guid();
}

module.exports = [{
  mode: devMode ? 'development' : 'production',
  entry: {
    excelio: path.resolve(BASE_DIR, './src/excel/index.js')
  },
  devtool: devMode ? 'source-map' : 'none',
  output: {
    path: path.resolve(BASE_DIR, 'dist/excel'),
    filename: '[name].js',
    library: 'ExcelIO',
    libraryTarget: 'umd',
    umdNamedDefine: true,
    globalObject: "typeof self !== 'undefined' ? self : this"
  },
  module: {
    rules: [{
      test: /(\.jsx|\.js)$/,
      loader: 'babel-loader',
      exclude: /(node_modules|bower_components)/
    },
    {
      test: /(\.jsx|\.js)$/,
      loader: 'eslint-loader',
      exclude: /node_modules/
    }
    ]
  },
  resolve: {
    modules: [path.resolve('./src/excel'), 'node_modules'],
    extensions: ['.json', '.js']
  }
}, {
  mode: devMode ? 'development' : 'production',
  entry: {
    pluginBase: path.resolve(BASE_DIR, './src/plugins/pluginBase.js')
  },
  devtool: devMode ? 'source-map' : 'none',
  output: {
    path: path.resolve(BASE_DIR, 'dist/plugins'),
    filename: '[name].js'
  },
  module: {
    rules: [{
      test: /(\.jsx|\.js)$/,
      loader: 'babel-loader',
      exclude: /(node_modules|bower_components)/
    },
    {
      test: /(\.jsx|\.js)$/,
      loader: 'eslint-loader',
      exclude: /node_modules/
    }
    ]
  },
  resolve: {
    modules: [path.resolve('./src/plugins/'), 'node_modules'],
    extensions: ['.json', '.js']
  }
}, {
  mode: devMode ? 'development' : 'production',
  entry: {
    background: path.resolve(BASE_DIR, './src/plugins/background/background.js')
  },
  devtool: devMode ? 'source-map' : 'none',
  output: {
    path: path.resolve(BASE_DIR, 'dist/plugins/background'),
    filename: '[name].js'
  },
  resolveLoader: {
    alias: {
      'config-loader': require.resolve('./loaders/config.js')
    }
  },
  module: {
    rules: [{
      test: /(\.jsx|\.js)$/,
      loader: 'babel-loader',
      exclude: /(node_modules|bower_components)/
    }, {
      test: /config\.json$/,
      loader: 'config-loader',
      exclude: /(node_modules|bower_components)/,
      options: {
        compress: true,
        definitions: {
          'process.env': {
            BACKGROUND_NAME: JSON.stringify(process.env.BACKGROUND_NAME),
            BACKGROUND_GUID: JSON.stringify(process.env.BACKGROUND_GUID)
          }
        }
      }
    }]
  },
  plugins: [
    new webpack.DefinePlugin({
      'process.env': {
        PLUGIN_NAME: JSON.stringify(process.env.BACKGROUND_NAME)
      }
    }),
    new IncludeFilePlugin({
      'config.json': path.resolve(BASE_DIR, './src/plugins/background/config.json')
    }),
    new CopyPlugin([{
      from: './src/plugins/background/*.png', flatten: true
    }, {
      from: './src/plugins/background/translations', to: 'translations/', flatten: true
    }]),
    new HtmlWebpackPlugin({
      inject: 'head',
      hash: true,
      minify: {
        // 移除注释
        removeComments: true,
        // 删除空白符和换行符
        collapseWhitespace: true
      },
      compress: true,
      title: process.env.BACKGROUND_NAME,
      template: './src/plugins/background/index.html'
    })
  ],
  resolve: {
    modules: [path.resolve('./src/plugins/'), 'node_modules'],
    extensions: ['.json', '.js'],
    alias: {
      'debug-factory': path.resolve(BASE_DIR, DEBUG_FACTORY)
    }
  }
}, {
  mode: devMode ? 'development' : 'production',
  entry: {
    view: [
      path.resolve(BASE_DIR, './src/plugins/view/view.js'),
      path.resolve(BASE_DIR, './src/plugins/view/view.css')
    ]
  },
  devtool: devMode ? 'source-map' : 'none',
  output: {
    path: path.resolve(BASE_DIR, 'dist/plugins/view'),
    filename: '[name].js'
  },
  resolveLoader: {
    alias: {
      'config-loader': require.resolve('./loaders/config.js')
    }
  },
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
    }, {
      test: /\.css$/,
      loaders: [
        MiniCssExtractPlugin.loader, 'css-loader'
      ]
    }, {
      test: /config\.json$/,
      loader: 'config-loader',
      exclude: /(node_modules|bower_components)/,
      options: {
        compress: true,
        definitions: {
          'process.env': {
            VIEW_NAME: JSON.stringify(process.env.VIEW_NAME),
            VIEW_GUID: JSON.stringify(process.env.VIEW_GUID)
          }
        }
      }
    }]
  },
  optimization: {
    minimize: true
  },
  plugins: [
    new webpack.DefinePlugin({
      'process.env': {
        PLUGIN_NAME: JSON.stringify(process.env.VIEW_NAME)
      }
    }),
    new IncludeFilePlugin({
      'config.json': path.resolve(BASE_DIR, './src/plugins/view/config.json')
    }),
    new CopyPlugin([{
      from: './src/plugins/view/*.png', flatten: true
    }, {
      from: './src/plugins/view/runtime.js', flatten: true
    }, {
      from: './src/plugins/view/translations', to: 'translations/', flatten: true
    }]),
    new MiniCssExtractPlugin(),
    new OptimizeCssAssetsPlugin({
      cssProcessor: require('cssnano'),
      cssProcessorPluginOptions: {
        preset: ['default', { discardComments: { removeAll: true } }],
      },
      canPrint: true
    }),
    new HtmlWebpackPlugin({
      inject: 'head',
      hash: true,
      minify: {
        // 移除注释
        removeComments: true,
        // 删除空白符和换行符
        collapseWhitespace: true
      },
      compress: true,
      title: process.env.VIEW_NAME,
      template: './src/plugins/view/index.html'
    })
  ],
  resolve: {
    modules: [path.resolve('./src/plugins/'), 'node_modules'],
    extensions: ['.json', '.js', '.css'],
    alias: {
      'debug-factory': path.resolve(BASE_DIR, DEBUG_FACTORY)
    }
  }
}, {
  mode: devMode ? 'development' : 'production',
  entry: {
    api: path.resolve(BASE_DIR, './src/plugins/api/index.js')
  },
  devtool: devMode ? 'source-map' : 'none',
  output: {
    path: path.resolve(BASE_DIR, 'dist/plugins'),
    filename: '[name].js',
    library: 'ExcelAPI',
    libraryTarget: 'umd',
    umdNamedDefine: true,
    globalObject: "typeof self !== 'undefined' ? self : this"
  },
  externals: {
    onlyffice: 'DocsAPI'
  },
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
    },
    {
      test: /(\.jsx|\.js)$/,
      loader: 'eslint-loader',
      exclude: /node_modules/
    }]
  },
  plugins: [
    new webpack.DefinePlugin({
      'process.env': {
        VIEW_GUID: JSON.stringify(process.env.VIEW_GUID),
        BACKGROUND_GUID: JSON.stringify(process.env.BACKGROUND_GUID)
      }
    })
  ],
  resolve: {
    modules: [path.resolve('./src'), 'node_modules'],
    extensions: ['.json', '.js'],
    alias: {
      'onlyoffice-excelio': path.resolve(BASE_DIR, './src/excel/index.js'),
      'debug-factory': path.resolve(BASE_DIR, DEBUG_FACTORY)
    }
  }
}]