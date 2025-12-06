import isArray from "lodash/isArray";
import isFunction from "lodash/isFunction";
import isObject from "lodash/isObject";
import cloneDeep from "lodash/cloneDeep";
import { saveAs } from "file-saver";
import { IExcelOptions, IFormatData, ISubTitle, IExcelToJsonOptions } from "./type";

class Workbook {
  SheetNames: string[] = [];
  Sheets: { [sheet: string]: any } = {};
  constructor() {}
}

const baseResourceUrl = "https://cdn.jsdelivr.net/npm/pl-export-excel@1.1.7/dist/xlsx.core.min.js";
const loadScript = (src: string): Promise<void> => {
  // Check if there are already script tags with the same src
  if (document.querySelector(`script[src="${src}"]`)) {
    return Promise.resolve();
  }
  return new Promise((resolve, reject) => {
    const scriptElement = document.createElement("script");
    scriptElement.src = src;
    scriptElement.async = true;
    scriptElement.crossOrigin = "anonymous";

    scriptElement.onload = () => resolve();
    scriptElement.onerror = () => reject(new Error(`Failed to load ${src}`));
    if (!document || !document.body) {
      return reject(new Error("Document body not found"));
    }
    document.body.appendChild(scriptElement);
  });
};

function isTwoDimensionalArray(arr) {
  if (!Array.isArray(arr)) {
    return false;
  }
  return arr.every((item) => Array.isArray(item));
}

const formatDateToYYYYMMDD = (date: Date): string => {
  if (date && date instanceof Date) {
    const year = date.getFullYear();
    // padStart: 在字符串开头填充指定字符，直到字符串达到指定长度。
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  return "";
};
/**
 * format data
 * @param headers Header configuration
 * @param datas table data
 */
const formatData = (datas: IExcelOptions["datas"], headers: IExcelOptions["headers"]): IFormatData => {
  const listen = datas.map((item) => {
    return headers.map((header) => (header?.key ? item[header?.key] || "" : ""));
  });
  const h = headers.map((item) => item.title);

  listen.unshift(h);
  // console.log(listen, headers, datas);
  return listen;
};

const expandConfig = (options: IExcelOptions) => {
  const defaults = {
    bookType: "xlsx",
    autoWidth: true,
    filename: "excel-list",
    xlsxStyleResourceUrl: baseResourceUrl
  } as Partial<IExcelOptions>;
  return Object.assign(defaults, options);
};

/** Convert string to ArrayBuffer */
const s2ab = (s: string) => {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
  return buf;
};

const loadResource = (url: string): Promise<any> => {
  return Promise.all([loadScript(url)]);
};

const two_array_to_sheet = (data: IFormatData) => {
  const ws = {};
  const range = {
    s: {
      c: 10000000,
      r: 10000000
    },
    e: {
      c: 0,
      r: 0
    }
  };

  for (let R = 0; R != data.length; ++R) {
    for (let C = 0; C != data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R;
      if (range.s.c > C) range.s.c = C;
      if (range.e.r < R) range.e.r = R;
      if (range.e.c < C) range.e.c = C;
      const cell: any = {
        v: data[R][C]
      };
      if (cell.v == null) continue;
      const cell_ref = XLSX.utils.encode_cell({
        c: C,
        r: R
      });

      if (typeof cell.v === "number") cell.t = "n";
      else if (typeof cell.v === "boolean") cell.t = "b";
      else if (cell.v instanceof Date) {
        cell.t = "n";
        // Built in format table
        cell.z = XLSX.SSF._table[14];
        cell.v = formatDateToYYYYMMDD(cell.v);
      } else cell.t = "s";
      ws[cell_ref] = cell;
    }
  }
  // convert cell range
  if (range.s.c < 10000000) ws["!ref"] = XLSX.utils.encode_range(range);
  return ws;
};

/**
 * Handling title
 * @param titleConfig title configuration
 * @param defaultColNum Default is the length of the headers field
 */
const handleTitle = (titleConfig: ISubTitle, defaultColNum: number) => {
  if (!titleConfig?.title) return null;
  const data: ISubTitle["title"][][] = [];
  const colNum = titleConfig?.colNum || defaultColNum || 1;
  const rowNum = titleConfig.rowNum || 1;

  data.push([titleConfig.title]);
  // Fill in row
  if (rowNum > 1) {
    for (let i = 1; i < rowNum; i++) {
      data.push([""]);
    }
  }
  //  Fill in col
  for (let i = 0; i < data.length; i++) {
    if (colNum > 1 && data[i]) {
      data[i].push(...new Array(colNum - 1).fill(""));
    }
  }

  titleConfig.colNum = colNum;
  titleConfig.rowNum = rowNum;

  // console.log(data, rowNum, colNum);
  return data;
};

/** handle title merges and style  */
const handleTitleMergesAndStyle = (ws: any, titleConfig: ISubTitle): any => {
  const exist = titleConfig.colNum && titleConfig.rowNum && titleConfig.title;
  if (!exist) return ws;
  try {
    const lvs = cloneDeep(ws);
    if (!lvs["!merges"]) lvs["!merges"] = [];
    lvs["!merges"].push({
      s: {
        c: 0,
        r: 0
      },
      e: {
        c: titleConfig.colNum - 1,
        r: titleConfig.rowNum - 1
      }
    });
    if (lvs["A1"]) {
      lvs["A1"].s = {
        font: {
          name: "微软雅黑",
          sz: 18,
          color: { rgb: "333333" },
          bold: true,
          italic: false,
          underline: false
        },
        alignment: {
          horizontal: "center",
          vertical: "center"
        }
      };
    }
    return lvs;
  } catch (error) {
    console.log(error);
    return ws;
  }
};

const mergesCells = (ws: any, merges: IExcelOptions["merges"]) => {
  if (!merges?.length) return ws;
  try {
    const lvs = cloneDeep(ws);
    if (merges.length > 0) {
      if (!lvs["!merges"]) lvs["!merges"] = [];
      merges.forEach((item) => {
        if (typeof item === "object") {
          lvs["!merges"].push(item);
        } else {
          lvs["!merges"].push(XLSX.utils.decode_range(item));
        }
      });
    }
    return lvs;
  } catch (error) {
    console.log(error);
    return ws;
  }
};

/** Column configuration information */
const handleAutoWidth = (ws: any, list: IFormatData) => {
  try {
    const lvs = cloneDeep(ws);
    /** Set the maximum width of each column in the worksheet */
    const colWidth = list.map((row) =>
      row.map((val: any) => {
        if (val == null) {
          return {
            wch: 10
          };
          // 超过 255 的字符属于扩展 Unicode 字符
        } else if (val.toString().charCodeAt(0) > 255) {
          return {
            wch: val.toString().length * 2
          };
        } else {
          return {
            wch: val.toString().length
          };
        }
      })
    );
    /** Starting with the first line as the initial value */
    let result = colWidth[0];
    for (let i = 1; i < colWidth.length; i++) {
      for (let j = 0; j < colWidth[i].length; j++) {
        if (result[j]) {
          if (result[j]["wch"] < colWidth[i][j]["wch"]) {
            result[j]["wch"] = colWidth[i][j]["wch"];
          }
        }
      }
    }

    lvs["!cols"] = result;
    return lvs;
  } catch (error) {
    console.log(error);
    return ws;
  }
};

export const exportJsonToExcel = async (options: IExcelOptions) => {
  const mergeOptions = expandConfig(options);
  const { titleConfig, multiHeader, filename, bookType, merges, autoWidth, styleCb, xlsxStyleResourceUrl } = mergeOptions;
  const createWorkbook = (): Workbook => {
    return new Workbook();
  };
  const init = async () => {
    // is data array
    if (!isArray(mergeOptions.datas)) {
      console.error("The table data is not of array type, please check the 'datas' fields");
      return { ws: null, list: [] };
    }

    if (!isArray(mergeOptions.headers) || !mergeOptions.headers?.length) {
      console.error("The header is empty, please check the 'headers' fields");
      return { ws: null, list: [] };
    }

    try {
      await loadResource(xlsxStyleResourceUrl);
      if (!(window as any)?.XLSX) {
        throw new Error("XLSX not found, Please check if xlsxStyleResource is loaded");
      }
      const list = cloneDeep(mergeOptions.datas);
      const data = formatData(list, mergeOptions.headers);
      if (!data.length) return { ws: null, list: [] };

      // Handling multi_headers
      if (isArray(multiHeader) && multiHeader.length) {
        for (let i = multiHeader.length - 1; i > -1; i--) {
          data.unshift(multiHeader[i] as any);
        }
      }

      // Handling title
      if (titleConfig) {
        const titleData = handleTitle(titleConfig, mergeOptions.headers.length);
        data.unshift(...titleData);
      }
      // array_to_sheet
      let ws = two_array_to_sheet(data);
      // console.log(ws, data);
      return { ws, list: data };
    } catch (error) {
      console.error("Failed:", error);
      return { ws: null, list: [] };
    }
  };

  const wb = createWorkbook();

  let { ws, list } = await init();

  if (!ws || !list.length) return;

  if (isArray(merges)) {
    ws = mergesCells(ws, merges);
  }

  if (autoWidth) {
    ws = handleAutoWidth(ws, list);
  }
  // Processing Title Styles
  if (titleConfig) {
    ws = handleTitleMergesAndStyle(ws, titleConfig);
  }
  // add custom style
  if (isFunction(styleCb)) {
    styleCb(ws);
  }
  // console.log(ws);
  try {
    wb.SheetNames.push("Excel");
    wb.Sheets["Excel"] = ws;
    const wbout = XLSX.write(wb, {
      bookType: bookType,
      bookSST: false,
      // binary: binary string (byte n is data.charCodeAt(n))
      type: "binary"
    });
    // console.log(wb);
    saveAs(
      new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
      }),
      `${filename}.${bookType}`
    );
  } catch (error) {
    console.error("Failed:", error);
  }
};

export const excelToJson = (options: IExcelToJsonOptions): Promise<{ originalList: any[]; formatList: any[] }> => {
  return new Promise(async (resolve, reject) => {
    const { file, keys, startRow, xlsxStyleResourceUrl } = isObject(options) ? options : ({} as IExcelToJsonOptions);
    try {
      if (!file || !(file instanceof File)) {
        throw new Error("file is required， or it's not a file");
      }
      await loadResource(xlsxStyleResourceUrl || baseResourceUrl);
      if (!(window as any)?.XLSX) {
        throw new Error("XLSX not found, Please check if xlsxStyleResource is loaded");
      }
      const reader = new FileReader();
      reader.onload = function (e: any) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheetName = workbook.SheetNames[0];
          if (!firstSheetName) {
            resolve({ formatList: [], originalList: [] });
            return;
          }
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          // Convert to array object
          if (keys?.length && isTwoDimensionalArray(jsonData)) {
            const list = startRow ? jsonData.slice(startRow) : jsonData;
            const cvt = list.map((items) => {
              const obj: any = {};
              items.forEach((item, index) => {
                obj[keys[index]] = item || "";
              });
              return obj;
            });
            resolve({ formatList: cvt, originalList: jsonData });
            return;
          }
          // Not converting
          resolve({ formatList: [], originalList: isArray(jsonData) ? jsonData : [] });
        } catch (error) {
          console.error("Failed:", error);
          reject(error);
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (error) {
      console.error("Failed:", error);
      reject(error);
    }
  });
};
