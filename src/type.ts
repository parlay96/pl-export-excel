type IMerges = { s: { c: number; r: number }; e: { c: number; r: number } };

export type IFormatData = (string | number)[][];

export interface IExcelOptions {
  /** 表格数据 */
  datas: { [key: string]: string | number | null }[];
  /** 头部 */
  header: string[];
  /** 需要导出表格数据中的字段key名 */
  keys?: string[];
  /** 文件名 */
  filename?: string;
  /** 多级头 */
  multiHeader?: string[][];
  /** 合并单元格数组 */
  merges?: IMerges[];
  /** 宽度是否自适应 */
  autoWidth?: boolean;
  /** 导出的文件类型 */
  bookType?: "xlsx" | "xlsm" | "xlsb" | "xlml" | "csv" | "txt" | "html";
  /** 自定义xlsx样式资源地址 */
  xlsxStyleResourceUrl?: string;
  /** 设置单元格样式 */
  styleCb?: (ws: any) => any;
}
