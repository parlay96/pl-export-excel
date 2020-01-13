var path = require('path');
var uglify = require('uglifyjs-webpack-plugin');

module.exports = {
  entry: "./index.js",
  output: {
    path: path.resolve(process.cwd(), './lib'),
    publicPath: '/lib/',
    filename: 'pl-export-excel.js',
    libraryExport: 'default', // 对外暴露default属性，就可以直接调用default里的属性
    library: 'PlExportExcel',
    libraryTarget: 'commonjs2'
  },
  resolve: {
    extensions: ['.js'],
  },
  plugins: [
      new uglify(),
  ]
};
