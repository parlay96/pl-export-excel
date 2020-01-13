module.exports =
/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "/lib/";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_xlsx_style__ = __webpack_require__(4);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_xlsx_style___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0_xlsx_style__);
__webpack_require__(1);


function isArrayFn(value){
    if (typeof Array.isArray === "function") {
        return Array.isArray(value);
    }else{
        return Object.prototype.toString.call(value) === "[object Array]";
    }
}

function generateArray(table) {
    var out = [];
    var rows = table.querySelectorAll('tr');
    var ranges = [];
    for (var R = 0; R < rows.length; ++R) {
        var outRow = [];
        var row = rows[R];
        var columns = row.querySelectorAll('td');
        for (var C = 0; C < columns.length; ++C) {
            var cell = columns[C];
            var colspan = cell.getAttribute('colspan');
            var rowspan = cell.getAttribute('rowspan');
            var cellValue = cell.innerText;
            if (cellValue !== "" && cellValue == +cellValue) cellValue = +cellValue;

            //Skip ranges
            ranges.forEach(function (range) {
                if (R >= range.s.r && R <= range.e.r && outRow.length >= range.s.c && outRow.length <= range.e.c) {
                    for (var i = 0; i <= range.e.c - range.s.c; ++i) outRow.push(null);
                }
            });

            //Handle Row Span
            if (rowspan || colspan) {
                rowspan = rowspan || 1;
                colspan = colspan || 1;
                ranges.push({
                    s: {
                        r: R,
                        c: outRow.length
                    },
                    e: {
                        r: R + rowspan - 1,
                        c: outRow.length + colspan - 1
                    }
                });
            }
            ;

            //Handle Value
            outRow.push(cellValue !== "" ? cellValue : null);

            //Handle Colspan
            if (colspan)
                for (var k = 0; k < colspan - 1; ++k) outRow.push(null);
        }
        out.push(outRow);
    }
    return [out, ranges];
};

function datenum(v, date1904) {
    if (date1904) v += 1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data, opts) {
    var ws = {};
    var range = {
        s: {
            c: 10000000,
            r: 10000000
        },
        e: {
            c: 0,
            r: 0
        }
    };
    for (var R = 0; R != data.length; ++R) {
        for (var C = 0; C != data[R].length; ++C) {
            if (range.s.r > R) range.s.r = R;
            if (range.s.c > C) range.s.c = C;
            if (range.e.r < R) range.e.r = R;
            if (range.e.c < C) range.e.c = C;
            var cell = {
                v: data[R][C]
            };
            if (cell.v == null) continue;
            var cell_ref = __WEBPACK_IMPORTED_MODULE_0_xlsx_style___default.a.utils.encode_cell({
                c: C,
                r: R
            });

            if (typeof cell.v === 'number') cell.t = 'n';
            else if (typeof cell.v === 'boolean') cell.t = 'b';
            else if (cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = __WEBPACK_IMPORTED_MODULE_0_xlsx_style___default.a.SSF._table[14];
                cell.v = datenum(cell.v);
            } else cell.t = 's';

            ws[cell_ref] = cell;
        }
    }
    if (range.s.c < 10000000) ws['!ref'] = __WEBPACK_IMPORTED_MODULE_0_xlsx_style___default.a.utils.encode_range(range);
    return ws;
}

function Workbook() {
    if (!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

/**
 * 格式化数据
 * @param filterVal 表头在数据中的对应key数组
 * @param jsonData 需要导出表格数据
 */
function formatJson (filterVal, jsonData) {
    return jsonData.map(v => filterVal.map(j => v[j] || ''))
}

/**
 * 通过table标签渲染导出表格
 * @param id 需要导出表格的ID
 * @param filename 表格名
 * @param bookType 文件类型
 * @param styleFun(参数是当前表格ws) 样式函数方法
 */
function export_table_to_excel ({id, filename = '空', bookType = 'xlsx', styleFun = () => {}} = {}) {
    if (id) {
        var theTable = document.getElementById(id);
        var ws = __WEBPACK_IMPORTED_MODULE_0_xlsx_style___default.a.utils.table_to_sheet(theTable);
        styleFun(ws);
        var wb = __WEBPACK_IMPORTED_MODULE_0_xlsx_style___default.a.utils.book_new();
        __WEBPACK_IMPORTED_MODULE_0_xlsx_style___default.a.utils.book_append_sheet(wb, ws, "SheetJS");
        var wbout = __WEBPACK_IMPORTED_MODULE_0_xlsx_style___default.a.write(wb, {
            bookType: bookType,
            bookSST: false,
            type: 'binary'
        });
        saveAs(new Blob([s2ab(wbout)], { type: "" }), filename + "." + bookType)
    }
}

/**
 * 通过json渲染导出表格
 * @param title 标题
 * @param multiHeader多级头
 * @param header 头部
 * @param data 表格数据
 * @param filename 表格名
 * @param merges 合并单元格数组
 * @param autoWidth 宽度是否自适应
 * @param bookType 文件类型
 * @param styleFun(参数是当前表格ws) 样式函数方法
 */
function export_json_to_excel (
    // 初始值
    { title = [], multiHeader = [], header = [], data = [],
        filename = '', merges = [], autoWidth = true,
        bookType = 'xlsx', styleFun = () => {}} = {}) {
    // 文件名
    filename = filename || 'excel-list'
    // 表数据
    if (!isArrayFn(data)) {
        console.error('表数据不是数组类型，请检查data')
        return
    }
    data = [...data]
    // 头部
    if (isArrayFn(header) && header.length) {
        data.unshift(header);
    }
    // 标题向数组的开头添加一个
    if (isArrayFn(title) && title.length) {
        data.unshift(title);
    }
    // 多级头
    if (isArrayFn(multiHeader) && multiHeader.length) {
        for (let i = multiHeader.length - 1; i > -1; i--) {
            data.unshift(multiHeader[i])
        }
    }
    // 总数据
    // console.log(data)
    var ws_name = "SheetJS";
    var wb = new Workbook(),
        ws = sheet_from_array_of_arrays(data);
    // 合并单元格
    if (merges.length > 0) {
        if (!ws['!merges']) ws['!merges'] = [];
        merges.forEach(item => {
            if (typeof item === 'object') {
                ws['!merges'].push(item)
            } else {
                ws['!merges'].push(__WEBPACK_IMPORTED_MODULE_0_xlsx_style___default.a.utils.decode_range(item))
            }
        })
    }
    // 设置宽度
    if (autoWidth) {
        /*设置worksheet每列的最大宽度*/
        const colWidth = data.map(row => row.map(val => {
            /*先判断是否为null/undefined*/
            if (val == null) {
                return {
                    'wch': 10
                };
            }
            /*再判断是否为中文*/
            else if (val.toString().charCodeAt(0) > 255) {
                return {
                    'wch': val.toString().length * 2
                };
            } else {
                return {
                    'wch': val.toString().length
                };
            }
        }))
        /*以第一行为初始值*/
        let result = colWidth[0];
        for (let i = 1; i < colWidth.length; i++) {
            for (let j = 0; j < colWidth[i].length; j++) {
                if (result[j]['wch'] < colWidth[i][j]['wch']) {
                    result[j]['wch'] = colWidth[i][j]['wch'];
                }
            }
        }
        ws['!cols'] = result;
    }

    /* add worksheet to workbook */
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;
    var dataInfo = wb.Sheets[wb.SheetNames[0]];
    // 暴露样式设置
    styleFun(dataInfo)
    // 导出excel
    var wbout = __WEBPACK_IMPORTED_MODULE_0_xlsx_style___default.a.write(wb, {
        bookType: bookType,
        bookSST: false,
        type: 'binary'
    });
    saveAs(new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
    }), `${filename}.${bookType}`);
}

/**
 * 通过table ID渲染导出表格
 * @param tableId 需要导出表格的ID
 * @param fileName 文件名
 * @param sheetName 表名
 */
// 导出Excel方法（表格id，不加扩展名的文件名，sheet名）
function exportExcelMethod (tableId, fileName, sheetName) {
    tableToExcel(tableId, fileName, sheetName)
}
const tableToExcel = (function() {
    const uri = 'data:application/vnd.ms-excel;base64,'
    const template = `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><meta charset="UTF-8"></head><body>{table}</body></html>`
    const base64 = function(s) { return window.btoa(unescape(encodeURIComponent(s))) }
    const format = function(s, c) { return s.replace(/{(\w+)}/g, function(m, p) { return c[p] }) }
    return function(table, filename, sheetname) {
        if (!table.nodeType) table = document.getElementById(table)
        const ctx = { worksheet: sheetname || 'Worksheet', table: table.innerHTML }
        const aTag = document.createElement('a')
        aTag.href = uri + base64(format(template, ctx))
        aTag.download = filename
        aTag.click()
    }
})()
/* harmony default export */ __webpack_exports__["default"] = ({
    exportExcelMethod,
    export_json_to_excel,
    export_table_to_excel,
    formatJson,
});


/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

__webpack_require__(2)(__webpack_require__(3))

/***/ }),
/* 2 */
/***/ (function(module, exports) {

/*
	MIT License http://www.opensource.org/licenses/mit-license.php
	Author Tobias Koppers @sokra
*/
module.exports = function(src) {
	function log(error) {
		(typeof console !== "undefined")
		&& (console.error || console.log)("[Script Loader]", error);
	}

	// Check for IE =< 8
	function isIE() {
		return typeof attachEvent !== "undefined" && typeof addEventListener === "undefined";
	}

	try {
		if (typeof execScript !== "undefined" && isIE()) {
			execScript(src);
		} else if (typeof eval !== "undefined") {
			eval.call(null, src);
		} else {
			log("EvalError: No eval function available");
		}
	} catch (error) {
		log(error);
	}
}


/***/ }),
/* 3 */
/***/ (function(module, exports) {

module.exports = "(function(a,b){if(\"function\"==typeof define&&define.amd)define([],b);else if(\"undefined\"!=typeof exports)b();else{b(),a.FileSaver={exports:{}}.exports}})(this,function(){\"use strict\";function b(a,b){return\"undefined\"==typeof b?b={autoBom:!1}:\"object\"!=typeof b&&(console.warn(\"Deprecated: Expected third argument to be a object\"),b={autoBom:!b}),b.autoBom&&/^\\s*(?:text\\/\\S*|application\\/xml|\\S*\\/\\S*\\+xml)\\s*;.*charset\\s*=\\s*utf-8/i.test(a.type)?new Blob([\"\\uFEFF\",a],{type:a.type}):a}function c(b,c,d){var e=new XMLHttpRequest;e.open(\"GET\",b),e.responseType=\"blob\",e.onload=function(){a(e.response,c,d)},e.onerror=function(){console.error(\"could not download file\")},e.send()}function d(a){var b=new XMLHttpRequest;b.open(\"HEAD\",a,!1);try{b.send()}catch(a){}return 200<=b.status&&299>=b.status}function e(a){try{a.dispatchEvent(new MouseEvent(\"click\"))}catch(c){var b=document.createEvent(\"MouseEvents\");b.initMouseEvent(\"click\",!0,!0,window,0,0,0,80,20,!1,!1,!1,!1,0,null),a.dispatchEvent(b)}}var f=\"object\"==typeof window&&window.window===window?window:\"object\"==typeof self&&self.self===self?self:\"object\"==typeof global&&global.global===global?global:void 0,a=f.saveAs||(\"object\"!=typeof window||window!==f?function(){}:\"download\"in HTMLAnchorElement.prototype?function(b,g,h){var i=f.URL||f.webkitURL,j=document.createElement(\"a\");g=g||b.name||\"download\",j.download=g,j.rel=\"noopener\",\"string\"==typeof b?(j.href=b,j.origin===location.origin?e(j):d(j.href)?c(b,g,h):e(j,j.target=\"_blank\")):(j.href=i.createObjectURL(b),setTimeout(function(){i.revokeObjectURL(j.href)},4E4),setTimeout(function(){e(j)},0))}:\"msSaveOrOpenBlob\"in navigator?function(f,g,h){if(g=g||f.name||\"download\",\"string\"!=typeof f)navigator.msSaveOrOpenBlob(b(f,h),g);else if(d(f))c(f,g,h);else{var i=document.createElement(\"a\");i.href=f,i.target=\"_blank\",setTimeout(function(){e(i)})}}:function(a,b,d,e){if(e=e||open(\"\",\"_blank\"),e&&(e.document.title=e.document.body.innerText=\"downloading...\"),\"string\"==typeof a)return c(a,b,d);var g=\"application/octet-stream\"===a.type,h=/constructor/i.test(f.HTMLElement)||f.safari,i=/CriOS\\/[\\d]+/.test(navigator.userAgent);if((i||g&&h)&&\"object\"==typeof FileReader){var j=new FileReader;j.onloadend=function(){var a=j.result;a=i?a:a.replace(/^data:[^;]*;/,\"data:attachment/file;\"),e?e.location.href=a:location=a,e=null},j.readAsDataURL(a)}else{var k=f.URL||f.webkitURL,l=k.createObjectURL(a);e?e.location=l:location.href=l,e=null,setTimeout(function(){k.revokeObjectURL(l)},4E4)}});f.saveAs=a.saveAs=a,\"undefined\"!=typeof module&&(module.exports=a)});\n\n//# sourceMappingURL=FileSaver.min.js.map"

/***/ }),
/* 4 */
/***/ (function(module, exports) {

module.exports = require("XLSX");

/***/ })
/******/ ])["default"];