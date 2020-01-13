require('script-loader!file-saver');
import XLSX from 'xlsx-style'

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
            var cell_ref = XLSX.utils.encode_cell({
                c: C,
                r: R
            });

            if (typeof cell.v === 'number') cell.t = 'n';
            else if (typeof cell.v === 'boolean') cell.t = 'b';
            else if (cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            } else cell.t = 's';

            ws[cell_ref] = cell;
        }
    }
    if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
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
        var ws = XLSX.utils.table_to_sheet(theTable);
        styleFun(ws);
        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
        var wbout = XLSX.write(wb, {
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
                ws['!merges'].push(XLSX.utils.decode_range(item))
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
    var wbout = XLSX.write(wb, {
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
export default {
    exportExcelMethod,
    export_json_to_excel,
    export_table_to_excel,
    formatJson,
}
