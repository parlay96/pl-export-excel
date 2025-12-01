import isArray from "lodash/isArray";
import isFunction from "lodash/isFunction";
import cloneDeep from "lodash/cloneDeep";
import { saveAs } from "file-saver";
import { IExcelOptions } from "./type";
import { expandConfig, loadResource, formatData, s2ab, two_array_to_sheet, mergesCells, handleAutoWidth } from "./utils";

class Workbook {
  SheetNames: string[] = [];
  Sheets: { [sheet: string]: any } = {};
  constructor() {}
}
export const exportJsonToExcel = async (options: IExcelOptions) => {
  const mergeOptions = expandConfig(options);
  const { header, filename, bookType, multiHeader, merges, autoWidth, styleCb, xlsxStyleResourceUrl } = mergeOptions;
  const createWorkbook = (): Workbook => {
    return new Workbook();
  };
  const init = async () => {
    // is data array
    if (!isArray(mergeOptions.datas)) {
      console.error("The table data is not of array type, please check the 'datas' fields");
      return;
    }

    try {
      await loadResource(xlsxStyleResourceUrl);
      if (!(window as any)?.XLSX) {
        throw new Error("XLSX not found, Please check if xlsxStyleResource is loaded");
      }
      const list = cloneDeep(mergeOptions.datas);
      const data = formatData(list, mergeOptions.keys);
      // header
      if (isArray(header) && header.length) {
        data.unshift(header);
      }
      // multi Header
      if (isArray(multiHeader) && multiHeader.length) {
        for (let i = multiHeader.length - 1; i > -1; i--) {
          data.unshift(multiHeader[i] as any);
        }
      }
      let ws = two_array_to_sheet(data);
      // add style
      if (isFunction(styleCb)) styleCb(ws);
      return { ws, list: data };
    } catch (error) {
      console.error("Failed:", error);
      return { ws: null, list: [] };
    }
  };

  const wb = createWorkbook();

  let { ws, list } = await init();

  if (!ws || list.length == 0) return;

  ws = mergesCells(ws, merges);

  if (autoWidth) {
    ws = handleAutoWidth(ws, list);
  }

  try {
    wb.SheetNames.push("SheetJS");
    wb.Sheets["SheetJS"] = ws;
    const wbout = XLSX.write(wb, {
      bookType: bookType,
      bookSST: false,
      // binary: binary string (byte n is data.charCodeAt(n))
      type: "binary"
    });
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
