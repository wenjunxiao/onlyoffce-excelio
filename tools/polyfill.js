/* eslint-disable */
/**
 * 支持使用8.0以下版本进行构建
 */
if (parseFloat(process.version.replace(/[^\d.]+/g, '')) < 8.0) {
  require('@babel/register')({
    only: [
      /eslint/,
      /comment-json/,
      /css-loader/
    ],
    'presets': [
      ['@babel/preset-env', {
        targets: {
          esmodules: true,
        }
      }]
    ]
  });
  require('@babel/polyfill');
  require('fs').copyFileSync = require('fs-copy-file-sync');
}
