/* eslint-disable */
module.exports = require(`${/webpack-dev-server/.test(process.argv[1]) || process.env.BUILD_SANDBOX ? './dev' : './prod'}`);