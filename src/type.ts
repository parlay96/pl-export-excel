/** 合并单元格配置：0开始计算 */
type IMerges = { s: { c: number; r: number }; e: { c: number; r: number } };

export type IFormatData = (string | number)[][];

/** 主标题配置 */
export type ISubTitle = {
  /** 标题文本 */
  title: string | number;
  /** 需要合并行数 */
  rowNum: number;
  /** 需要合并列数-默认为headers字段的长度 */
  colNum?: number;
};

/** 表头部配置 */
export type IHeaders = {
  /** 表头标题文本 */
  title: string | number | null;
  /** key值: 对应表格数据中的字段key名 */
  key: string;
};

export interface IExcelOptions {
  /** 表头部配置 */
  headers: IHeaders[];
  /** 表格数据 */
  datas: { [key: string]: string | number | null }[];
  /** 主标题配置 */
  titleConfig?: ISubTitle;
  /** 多级头 */
  multiHeader?: string[][];
  /** 文件名 */
  filename?: string;
  /** 合并单元格数组 */
  merges?: IMerges[];
  /** 宽度是否自适应 */
  autoWidth?: boolean;
  /** 导出的文件类型 */
  bookType?: "xlsx" | "xlsm" | "xlsb" | "xlml" | "csv" | "txt" | "html";
  /** 设置单元格样式 */
  styleCb?: (ws: any) => any;
  /** 自定义xlsx样式资源地址，基本用不到 */
  xlsxStyleResourceUrl?: string;
}

export interface IExcelToJsonOptions {
  /** 转换的文件 */
  file: File;
  /** 指定某一行之前为消除标题条目；当表格有标题时，或者某些行你想跳过时这是非常有用的。 */
  startRow?: number;
  /** 转成数组对象必填，单元格每列它所对应的key值;*/
  keys?: string[];
  /** 自定义xlsx样式资源地址，基本用不到 */
  xlsxStyleResourceUrl?: string;
}

export type IOriginalList = (string | number | null)[][];
