import _ from 'lodash';
import { parse as qsParse } from 'qs';

function pad (s) {
  return ('0' + s).slice(-2);
}
/**
 * 格式化时间
 * @param {Date} d 时间
 */
function formatDate (d) {
  return d.getFullYear() + '-' + pad(d.getMonth()) + '-' + pad(d.getDate())
    + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
}
function createTextImage (txt, width, height) {
  txt = txt || 'ExcelIO';
  const stamp = formatDate(new Date());
  const c = document.createElement('canvas');
  const ctx0 = c.getContext('2d');
  const font0 = ctx0.font = ctx0.font.split(' ').map(s => {
    if (/\d+px/.test(s)) {
      return '30px';
    }
    return s;
  }).join(' ');
  const w0 = ctx0.measureText(txt).width;
  const h0 = parseInt(ctx0.font);
  const ctx1 = c.getContext('2d');
  const font1 = ctx1.font = ctx1.font.split(' ').map(s => {
    if (/\d+px/.test(s)) {
      return '10px';
    }
    return s;
  }).join(' ');
  const w1 = ctx1.measureText(stamp).width;
  const h1 = parseInt(ctx1.font);
  if (width && height) {
    c.width = width;
    c.height = height * c.width / width;
  } else {
    c.width = Math.max(w0, w1) + 4;
    c.height = h0 + h1 + 8;
  }
  const d0 = c.getContext('2d');
  d0.font = font0;
  d0.fillStyle = 'green';
  d0.fillText(txt, (c.width - w0) / 2, (c.height * 3 / 4 + h0) / 2);
  const d1 = c.getContext('2d');
  d1.font = font1;
  d1.fillStyle = 'red';
  d1.fillText(stamp, (c.width - w1) / 2, (c.height / 4 + h1) / 2 + h0 + 6);
  return c.toDataURL();
}

// 先生成插件的icon提交到服务端，避免加载插件之后没有图片
const xhr = new XMLHttpRequest();
xhr.open('post', `${process.env.PUBLIC_PATH}plugins/${encodeURIComponent('view,background')}/icon.png?alias=${encodeURIComponent('icon@2x')}`);
xhr.onload = function () {
  route();
};
xhr.send(createTextImage(`${process.env.VIEW_NAME}`));

function watermarked (container, content, font) {
  const canvas0 = document.createElement('canvas');
  const ctx0 = canvas0.getContext("2d");
  if (/^\s*\d+(px|rem)?\s*$/.test(font)) {
    ctx0.font = ctx0.font.split(' ').map(s => {
      if (/\d+px/.test(s)) {
        return font;
      }
      return s;
    }).join(' ');
  } else if (font) {
    ctx0.font = font;
  }
  const textWidth = Math.ceil(ctx0.measureText(content).width);
  const canvas = document.createElement('canvas');
  const angle = Math.PI / 180 * -30;
  const sin = Math.abs(Math.sin(angle));
  const cos = Math.abs(Math.cos(angle));
  const width = textWidth;
  const height = Math.ceil(sin * textWidth);
  canvas.setAttribute('width', (width + 28) + 'px');
  canvas.setAttribute('height', (height + 2) + 'px');
  const ctx = canvas.getContext("2d");
  ctx.font = ctx0.font;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillStyle = 'rgba(184, 184, 184, 0.3)';
  ctx.rotate(angle + parseFloat(ctx.font) / textWidth);
  ctx.fillText(content, textWidth / 2 - height * sin + height / 6, height * cos);
  const xhr = new XMLHttpRequest();
  xhr.open('post', `${process.env.PUBLIC_PATH}plugins/watermark/icon.png`);
  xhr.onload = function () {
    route();
  };
  xhr.send(canvas.toDataURL());
  canvas.toBlob(blob => {
    const base64Url = URL.createObjectURL(blob);
    function createWatermark (url) {
      const div = document.createElement("div");
      div.setAttribute('class', 'watermarked');
      div.setAttribute('style', `visibility:visible;position:absolute;top:0;left:0;width:100%;min-height:100%;z-index:9999;pointer-events:none;background-repeat:repeat;background-image:url('${url}')`);
      return div;
    }
    let watermarkDiv = createWatermark(base64Url);
    container.style.position = 'relative';
    container.insertBefore(watermarkDiv, container.firstChild);
    const MutationObserver = window.MutationObserver ||
      window.WebKitMutationObserver ||
      window.MozMutationObserver;
    let lock = false;
    const mo = new MutationObserver(records => {
      if (lock) return;
      lock = true
      try {
        records.forEach(record => {
          if (record.type === 'childList') {
            record.removedNodes.forEach(node => {
              if (node === watermarkDiv) {
                container.insertBefore(watermarkDiv, container.firstChild);
              }
            })
          } else if (record.type === 'attributes') {
            watermarkDiv = createWatermark(base64Url);
            container.replaceChild(watermarkDiv, record.target);
            mo.observe(watermarkDiv, { attributes: true });
          }
        })
      } finally {
        lock = false;
      }
    });
    mo.observe(container, { childList: true });
    mo.observe(watermarkDiv, { attributes: true });
  });
}

watermarked(document.body, formatDate(new Date()), '16px');

const comps = {};

function openRouter (comp, name) {
  if (comp.mount) {
    comp.mount().then(() => {
      comps[name] = comp;
      comp.mounted();
    });
  }
}

function closeRouter (name) {
  let comp = comps[name];
  console.log('close router =>', name, comp)
  if (comp && comp.unmount) {
    comp.unmount();
  }
  delete comps[name];
}

let last;
let href;
function route () {
  if (href === location.href) return;
  href = location.href;
  let name = location.hash.replace(/^#([^?]*).*$/, '$1');
  if (!name) {
    name = 'home'
  }
  location.router = name;
  if (last) {
    closeRouter(last);
  }
  let divs = document.querySelectorAll('.content>div');
  for (let div of divs) {
    if (div.id === name) {
      div.style.display = 'block';
    } else {
      div.style.display = 'none';
    }
  }
  location.query = qsParse(location.href.replace(/^[^?]*\??/, ''));
  if (name === 'example') {
    require.ensure([], require => {
      openRouter(require('./example'), name);
    });
  } else if (name === 'command') {
    require.ensure([], require => {
      openRouter(require('./command'), name);
    });
  } else if (name === 'save') {
    require.ensure([], require => {
      openRouter(require('./save'), name);
    });
  } else if (name === 'scripts') {
    require.ensure([], require => {
      openRouter(require('./scripts'), name);
    });
  }
  last = name;
}

window.onpopstate = () => {
  route();
};