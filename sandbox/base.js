/**
 * 对应于需要集成office的CMS系统的office编辑页面
 */
let inited = false;

/**
 * 挂载dom到主页
 */
export function mount () {
  if (inited) return Promise.resolve();
  return new Promise((resolve, reject) => {
    let el = document.createElement('script');
    el.src = `${process.env.ONLYOFFICE_URL}`;
    el.onerror = reject;
    el.onload = function () {
      let el = document.createElement('script');
      el.src = `${process.env.PLUGIN_URL}`;
      el.onerror = reject;
      el.onload = function () {
        inited = true;
        return resolve();
      };
      document.body.appendChild(el);
    };
    document.body.appendChild(el);
  });
}

export function unmount () {
  if (window.docEditor) {
    window.docEditor.destroyEditor();
    window.docEditor = null;
  }
}
