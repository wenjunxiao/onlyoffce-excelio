export function hex2rgb (hex) {
  if (hex[0] === '#') {
    hex = hex.substr(1);
  }
  return parseInt('0x' + hex.slice(0, 2)) + ',' + parseInt('0x' + hex.slice(2, 4)) + ',' + parseInt('0x' + hex.slice(4, 6));
}
export function rgbToHex (r, g, b) {
  return ((r << 16) | (g << 8) | b).toString(16);
}
export function checkPrecision (precision, options) {
  if (typeof precision === 'object') {
    return [undefined, precision];
  }
  return [precision, options];
}
export function encodeCol (col) {
  let s = '';
  for (++col; col; col = Math.floor((col - 1) / 26))
    s = String.fromCharCode(((col - 1) % 26) + 65) + s;
  return s;
}
export function encodeCell (r, c) {
  return encodeCol(c) + (r + 1);
}
