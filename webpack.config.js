var path = require('path');
var uglify = require('uglifyjs-webpack-plugin');

module.exports = {
    entry: "./index.js",
    output: {
        path: path.resolve(__dirname, './lib'),
        publicPath: '/lib/',
        filename: 'index.js',
        library: 'PlExportExcel',
        libraryTarget: 'umd',
        libraryExport: "default", // 对外暴露default属性，就可以直接调用default里的属性
    },
    module: {},
    resolve: {
        extensions: ['.ts', '.js'],
    },
    plugins: [
        new uglify(),
    ]
}
