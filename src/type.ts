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
  styleCb?: (cellKeys: string[]) => { [key: string]: IOriginalStyles };
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

export type IOriginalStyles = Partial<{
  /** 单元格填充样式（背景填充）solid: 纯色填充 none: 无填充 */
  patternType: "solid" | "none";
  /** 填充前景色: 指定十六进制 RGB 值 */
  fgColor: string;
  /** 填充背景色: 指定十六进制 RGB 值 */
  // bgColor: string;
  /** 字体 */
  fontFamily: string;
  /** 字体大小 */
  fontSize: number;
  /** 字体颜色: 指定十六进制 RGB 值 */
  color: string;
  /** 是否加粗 */
  isBold: boolean;
  /** 是否下划线 */
  isUnderline: boolean;
  /** 是否斜体 */
  isItalic: boolean;
  /** 垂直对齐方式 */
  alignmentVertical: "bottom" | "center" | "top";
  /** 水平对齐方式 */
  alignmentHorizontal: "left" | "center" | "right";
  /** 是否自动换行 */
  wrapText: boolean;
  /** 上边框线 */
  borderTop: {
    style: IBORDER_STYLE;
    /** 边框颜色: 指定十六进制 RGB 值 */
    color: string;
  };
  /** 下边框线 */
  borderBottom: {
    style: IBORDER_STYLE;
    /** 边框颜色: 指定十六进制 RGB 值 */
    color: string;
  };
  /** 左边框线 */
  borderLeft: {
    style: IBORDER_STYLE;
    /** 边框颜色: 指定十六进制 RGB 值 */
    color: string;
  };
  /** 右边框线 */
  borderRight: {
    style: IBORDER_STYLE;
    /** 边框颜色: 指定十六进制 RGB 值 */
    color: string;
  };
}>;

export type IBORDER_STYLE =
  /** 细实线 */
  | "thin"
  /** 中实线 */
  | "medium"
  /** 粗实线 */
  | "thick"
  /** 虚线 */
  | "dashed"
  /** 点线 */
  | "dotted";
