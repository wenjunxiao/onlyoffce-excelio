/**
 * 适配Onlyoffice的Excel文档IO库
 * @module ExcelIO
 */

import ExcelWriter from './writer';

export default {
  /**
   * @type {ExcelIO.ExcelWriter}
   */
  ExcelWriter,
  /**
   * 创建Excel生成类实例
   * @function
   * @param {Object} [options] 选项
   * @returns {ExcelIO.ExcelWriter}
   */
  createWriter: (options = {}) => {
    return new ExcelWriter(options);
  }
};