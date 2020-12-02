export const getPublicUrl = url => {
  if (url[0] !== '/') url = '/' + url;
  return `${location.protocol}//${process.env.PUBLIC_HOST}${location.port && `:${location.port}` || ''}${url}`;
}

function post (url, data) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open('post', url);
    xhr.onerror = function (err) {
      reject(err);
    };
    xhr.onload = function () {
      resolve(xhr.responseText);
    };
    xhr.send(data);
  });
}

export function postService (service, data) {
  data.client = Date.now();
  console.log('[%s] request => ', service, data);
  return post(`/app/api/service/${service}`, JSON.stringify(data)).then(rsp => {
    console.log('[%s] response => ', service, rsp);
    return JSON.parse(rsp);
  });
}

export function uploadFile (data) {
  data.client = Date.now();
  return post(`/app/api/upload`, JSON.stringify(data)).then(rsp => {
    return JSON.parse(rsp);
  });
}
