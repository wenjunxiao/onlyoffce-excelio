/* eslint-disable */
const path = require('path');
const vm = require('vm');
const qs = require('querystring')
const _ = require('lodash');
const SingleEntryPlugin = require('webpack/lib/SingleEntryPlugin')

const pluginName = 'include-file-plugin';

function getCompilerName (context, filename) {
  const absolutePath = path.resolve(context, filename);
  const relativePath = path.relative(context, absolutePath);
  return pluginName + ' for "' + (absolutePath.length < relativePath.length ? absolutePath : relativePath) + '"';
}

class IncludeFilePlugin {
  constructor(entry, options = {}) {
    this.entry = entry || {}
    this.options = options || {}
    if (!this.options.filename) {
      this.options.filename = '[name]';
    }
  }

  apply (compiler) {
    compiler.hooks.thisCompilation.tap(pluginName, (compilation) => {
      const outputOptions = Object.assign({}, compilation.outputOptions, this.options);
      const compilerName = getCompilerName(compiler.context, outputOptions.path);
      const childCompiler = compilation.createChildCompiler(compilerName, outputOptions);
      childCompiler.context = compiler.context;
      for (const name of Object.keys(this.entry)) {
        const entry = new SingleEntryPlugin(childCompiler.context, this.entry[name], name)
        entry.apply(childCompiler);
      }
      compilation.hooks.additionalAssets.tapAsync(pluginName, (callback) => {
        childCompiler.runAsChild((err, entries, childCompilation) => {
          Object.keys(this.entry).map((name, index) => {
            const query = qs.parse(this.entry[name].replace(/^[^\?]*\?/, ''))
            const entry = childCompilation.entrypoints.get(name)
            if (entry.chunks.length > 1) {
              const error = new Error('more than expected chunks')
              error.chunks = entry.chunks
              throw error
            }
            const chunk = entry.chunks[0]
            if (!chunk) return {
              message: 'no chunk'
            }
            const filename = childCompilation.mainTemplate.hooks.assetPath.call(this.options.filename, {
              hash: childCompilation.hash,
              chunk: chunk,
              name: `${pluginName}_${index}`
            });
            if (chunk.files.length > 1) {
              let file = chunk.files.pop()
              if (/.map$/.test(file)) {
                if (chunk.files.length > 1) {
                  delete childCompilation.assets[chunk.files.pop()]
                }
                chunk.files.push(file)
              } else {
                delete childCompilation.assets[file]
                if (chunk.files.length > 1) {
                  const error = new Error('more than expected files:' + entry.files.join(','))
                  error.files = entry.files
                  throw error
                }
              }
            }
            const file = chunk.files[0]
            if (!file) return {
              message: 'no file'
            }
            if (file !== filename) {
              const error = new Error('expected file `' + filename + '`, but got `' + file + '`')
              throw error
            }
            const asset = childCompilation.assets[file]
            delete childCompilation.assets[file]
            const type = query.type || chunk.entryModule && chunk.entryModule.type
            if (/javascript/.test(type)) {
              compilation.assets[filename] = asset
              return {
                filename
              }
            }
            const source = asset.source()
            const content = this.evaluateCompilationResult(source, filename)
            compilation.assets[filename] = {
              source: () => content,
              size: () => content.length
            }
            return {
              filename,
              size: content.length
            }
          })
          callback()
        })
      });
    })
  }

  evaluateCompilationResult (source, filename) {
    if (!source) {
      throw (new Error('The child compilation didn\'t provide a result'))
    }
    // The LibraryTemplatePlugin stores the template result in a local variable.
    // To extract the result during the evaluation this part has to be removed.
    const vmContext = vm.createContext(_.extend({
      require: require
    }, global));
    let vmScript = null
    try {
      vmScript = new vm.Script(source, {
        filename
      });
    } catch (e) {
      return source
    }
    let tail = ''
    if (/(\n[^\w\n]+sourceMappingURL[^\n]*)$/i.test(source)) {
      tail = RegExp.$1
    }
    // Evaluate code and cast to string
    let newSource = vmScript.runInContext(vmContext);
    if (typeof newSource === 'object' && newSource.__esModule && newSource.default) {
      newSource = newSource.default;
    }
    return newSource && (newSource.toString() + tail)
  }
}

module.exports = IncludeFilePlugin