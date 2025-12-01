import cloneDeep from "lodash/cloneDeep";
import { IExcelOptions, IFormatData } from "./type";
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
 * @param filterVal The corresponding key array of the header in the data
 * @param datas table data
 */
export const formatData = (datas: IExcelOptions["datas"], keys?: string[]): IFormatData => {
  if (!keys?.length) {
    const keys = Object.keys(datas[0]); // 获取对象的键名数组
    return datas.map((item) => keys.map((key) => item[key]));
  }
  return datas.map((v) => keys.map((j) => v[j] || ""));
};

export const expandConfig = (options: IExcelOptions) => {
  const defaults = {
    bookType: "xlsx",
    autoWidth: true,
    filename: "excel-list",
    xlsxStyleResourceUrl: "https://unpkg.com/pl-export-excel@1.1.4/dist/xlsx.core.min.js"
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

export const loadResource = (xlsxStyleResourceUrl: string): Promise<any> => {
  return Promise.all([loadScript(xlsxStyleResourceUrl)]);
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

/** Column configuration information（ */
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
        if (result[j]["wch"] < colWidth[i][j]["wch"]) {
          result[j]["wch"] = colWidth[i][j]["wch"];
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
