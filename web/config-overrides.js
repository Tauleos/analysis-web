const {override, fixBabelImports} = require('customize-cra');
const Uglify = require('uglifyjs-webpack-plugin');
const compress = require('compression-webpack-plugin');
const addMyPlugin = config =>{
  config.plugins.push(new Uglify({
    uglifyOptions: {
      compress: {}
    }
  }))
  // config.plugins.push(new compress())
  return config;
}

module.exports = override(
  fixBabelImports('import', {
    libraryName: 'antd',
    libraryDirectory: 'es',
    style: 'css',
  }),
  addMyPlugin
);