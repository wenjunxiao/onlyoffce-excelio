/* eslint-disable */
const _ = require('lodash');
const loadUtils = require('loader-utils');

function spread (obj) {
  Object.keys(obj).forEach(k => {
    const v = obj[k];
    if (_.isPlainObject(v)) {
      _.set(obj, k.split('.'), spread(v));
    } else {
      _.set(obj, k.split('.'), v);
    }
  });
  return obj;
}

module.exports = function (source) {
  try {
    const options = loadUtils.getOptions(this);
    const definitions = spread(_.get(options, 'definitions') || {});
    const expand = obj => {
      if (Array.isArray(obj)) return obj.map(expand);
      if (_.isPlainObject(obj)) {
        Object.keys(obj).forEach(k => {
          const v = obj[k];
          if (typeof v === 'string') {
            obj[k] = v.replace(/\$\{\s*([^}]+?)\s*\}/g, ($0, $1) => {
              const v = _.get(definitions, $1);
              if (!v) return $0;
              try {
                return JSON.parse(v);
              } catch (_) {
                return v;
              }
            });
          } else if (Array.isArray(v)) {
            obj[k] = v.map(expand);
          }
        });
      }
      return obj;
    }
    const merged = require('comment-json').parse(source, null, true);
    const space = options.space === undefined && !options.compress ? 2 : options.space;
    const mergedJson = JSON.stringify(expand(merged), null, space);
    return JSON.stringify(mergedJson)
  } catch (err) {
    console.log('config load faild =>', source, err)
    throw err;
  }
};