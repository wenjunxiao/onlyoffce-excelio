import debugFactory from 'debug-factory';

const debug = debugFactory('excel-api');

function getParams (func) {
  var str = func.toString();
  str = str.replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/\/\/(.)*/g, '')
    .replace(/{[\s\S]*}/, '')
    .replace(/=>/g, '')
    .trim();
  var start = str.indexOf('(') + 1;
  var end = str.length - 1;
  var result = str.substring(start, end).split(', ');
  var params = [];
  result.forEach(element => {
    element = element.replace(/=[\s\S]*/g, '').trim();
    if (element.length > 0)
      params.push(element);
  });
  return params;
}

export {
  getParams,
  debug
}