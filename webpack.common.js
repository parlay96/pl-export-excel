var path = require('path');
var uglify = require('uglifyjs-webpack-plugin');
const utils = (path__) => {
    return path.resolve(__dirname, '..', path__)
}

module.exports = {
  entry: "./index.js",
  output: {
    path: path.resolve(process.cwd(), './lib'),
    publicPath: '/lib/',
    filename: 'pl-export-excel.js',
    libraryExport: 'default',
    library: 'PlExportExcel',
    libraryTarget: 'commonjs2'
  },
  externals: {
    'xlsx-style': 'XLSX'
  },
  resolve: {
    extensions: ['.js'],
  },
};
