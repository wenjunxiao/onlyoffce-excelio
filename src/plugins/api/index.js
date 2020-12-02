/**
 * Excel编辑器API,依赖 {@link external:DocsAPI}
 * @module ExcelAPI
 */
import ExcelIO from 'onlyoffice-excelio';
import WrappedEditor from './editor';
export * from './editor';
export {
  /**
   * ExcelIO
   * @member {module:ExcelIO} ExcelIO
   */
  ExcelIO
};
export {
  /**
   * WrappedEditor
   * @member {ExcelAPI.WrappedEditor} WrappedEditor
   */
  WrappedEditor
}

/**
 * 创建文档编辑器
 * @param {String} id 用于放置编辑器div的id,<i style="color:red;"><b>使用方法之前必须确保dom节点已存在</b></i>
 * @param {WrappedConfig} config 创建编辑器需要的配置
 * @returns {ExcelAPI.WrappedEditor} 文档编辑对象
 * @example
 * // 确保`id`为`placeholder`的`div`已经存在
 * window.onload = function () {
 *   let docEditor = ExcelAPI.createEditor('placeholder', {
 *     document: {
 *       title: 'example.xlsx',
 *       url: 'http://www.example.com/path/of/example.xlsx'
 *     },
 *     editorConfig: {
 *       mode: 'edit', // 'view' or 'edit'
 *       user: {
 *         id: "excelio",
 *         name: "Mr ExcelIO"
 *       }
 *     }
 *   });
 * };
 */
export function createEditor (id, config) {
  return new WrappedEditor(id, config);
}
