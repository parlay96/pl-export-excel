import cloneDeep from "lodash/cloneDeep";
import { IExcelOptions, IFormatData, ISubTitle } from "./type";

export const baseResourceUrl = "https://cdn.jsdelivr.net/npm/pl-export-excel@1.1.7/dist/xlsx.core.min.js";
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

export function isTwoDimensionalArray(arr) {
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
export const formatData = (datas: IExcelOptions["datas"], headers: IExcelOptions["headers"]): IFormatData => {
  const listen = datas.map((item) => {
    return headers.map((header) => (header?.key ? item[header?.key] || "" : ""));
  });
  const h = headers.map((item) => item.title);

  listen.unshift(h);
  // console.log(listen, headers, datas);
  return listen;
};

export const expandConfig = (options: IExcelOptions) => {
  const defaults = {
    bookType: "xlsx",
    autoWidth: true,
    filename: "excel-list",
    xlsxStyleResourceUrl: baseResourceUrl
  } as Partial<IExcelOptions>;
  return Object.assign(defaults, options);
};

/** Convert string to ArrayBuffer */
export const s2ab = (s: string) => {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
  return buf;
};

export const loadResource = (url: string): Promise<any> => {
  return Promise.all([loadScript(url)]);
};

export const two_array_to_sheet = (data: IFormatData) => {
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
export const handleTitle = (titleConfig: ISubTitle, defaultColNum: number) => {
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
export const handleTitleMergesAndStyle = (ws: any, titleConfig: ISubTitle): any => {
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

export const mergesCells = (ws: any, merges: IExcelOptions["merges"]) => {
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
export const handleAutoWidth = (ws: any, list: IFormatData) => {
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
