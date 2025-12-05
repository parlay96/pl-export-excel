# pl-export-excel

> Export Excel files to JSON data and convert JSON data to Excel files

# install

> npm i pl-export-excel

# exportJsonToExcel(options) => Promise<void>

|   options：参数名    | 参数描述                                          | 必填 |   类型   |    默认值    |
| :------------------: | ------------------------------------------------- | :--: | :------: | :----------: |
|        datas         | 表格数据                                          |  是  |  Array   |      []      |
|       headers        | 表头部配置, 配置描述请看文档下面                  |  是  |  Array   |      []      |
|     titleConfig      | 主标题配置, 配置描述请看文档下面                  |  否  |  object  |      -       |
|        merges        | 合并单元格配置, 配置描述请看文档下面              |  否  |  Array   |      -       |
|     multiHeader      | 多级头, 给表格设置多行头部,需要自己去合并单元格。 |  否  |  Array   |      -       |
|       filename       | 表格名                                            |  否  |  String  | 'excel-list' |
|      autoWidth       | 宽度是否自适应                                    |  否  | Boolean  |     true     |
|       bookType       | 文件类型                                          |  否  |  String  |    'xlsx'    |
|       styleCb        | 样式[style](#cell-styles)                         |  否  | Function |      -       |
| xlsxStyleResourceUrl | 自定义xlsx样式资源地址，基本用不到                |  否  |  String  |      -       |

# excelToJson(options)=> Promise<{ originalList: any[]; formatList: any[] }>

|   options：参数名    | 参数描述                                                                           | 必填 |  类型  | 默认值 |
| :------------------: | ---------------------------------------------------------------------------------- | :--: | :----: | :----: |
|         file         | 转换的文件                                                                         |  是  |  File  |   -    |
|       startRow       | 指定某一行之前为消除标题条目；当表格有标题时，或者某些行你想跳过时这是非常有用的。 |  否  | number |   -    |
|         keys         | 转成数组对象必填，单元格每列它所对应的key值                                        |  否  | Array  |   -    |
| xlsxStyleResourceUrl | 自定义xlsx样式资源地址，基本用不到                                                 |  是  | String |   -    |

# npm

```javascript
   import { exportJsonToExcel } from 'pl-export-excel'
   // 导出按钮方法
   handleEmits () {
    const headers = [
      { title: "经销商名称", key: "names" },
      { title: "下单时间", key: "date" },
      { title: "订单编号", key: "orderNumber" },
      { title: "客户名称", key: "customerName" }
    ]
    // 表格数据
    const datas = Array.from({ length: 200 }, (_, idx) => ({
      names: idx == 2 ? "大萨达萨达撒多少啊大" : "娃哈哈",
      age: (idx + 1) * 10,
      date: "201920120",
      orderNumber: idx + 1,
      customerName: "王小虎" + idx + 1,
    }));
    exportJsonToExcel({ headers, datas });
  }
```

# cdn

```javascript
/**
 * 引入pl-export-excel
   <body>
     <script src="https://cdn.jsdelivr.net/npm/pl-export-excel/dist/index.full.min.js"></script>;
     <script>
       const { exportJsonToExcel } = PlExportExcel;
       // 导出按钮方法
       // handleEmits () {.....}
     </script>
   </body>
 */
```

### ts type declare

```ts
/** 主标题配置 */
export type ISubTitle = {
  /** 标题文本 */
  title: string | number;
  /** 需要合并行数 */
  rowNum: number;
  /** 需要合并列数-默认为headers字段的长度 */
  colNum?: number;
};

/** 表头部配置, 导出后表格头部的顺序就是数组的顺序 */
export type IHeaders = {
  /** 表头标题文本 */
  title: string | number;
  /** key值: 对应表格数据中的字段key名 */
  key: string;
}[];

/**
 * 合并单元格配置： c or r: 0开始计算
 * 示例：
 * {
     s: {
        // s为开始
        c: 0, //开始列, 0开始计算
        r: 0 //开始行, 0开始计算
      },
     e: {
        // e结束
        c: 5, //结束列, 0开始计算
        r: 1 //结束行, 0开始计算
      }
   }
 */
export type IMerges = { s: { c: number; r: number }; e: { c: number; r: number } };
```

# Cell Styles

Cell styles are specified by a style object that roughly parallels the OpenXML structure. The style object has five
top-level attributes: `fill`, `font`, `numFmt`, `alignment`, and `border`.

| Style Attribute | Sub Attributes | Values                                                                                        |
| :-------------- | :------------- | :-------------------------------------------------------------------------------------------- |
| fill            | patternType    | `"solid"` or `"none"`                                                                         |
|                 | fgColor        | `COLOR_SPEC`                                                                                  |
|                 | bgColor        | `COLOR_SPEC`                                                                                  |
| font            | name           | `"Calibri"` // default                                                                        |
|                 | sz             | `"11"` // font size in points                                                                 |
|                 | color          | `COLOR_SPEC`                                                                                  |
|                 | bold           | `true` or `false`                                                                             |
|                 | underline      | `true` or `false`                                                                             |
|                 | italic         | `true` or `false`                                                                             |
|                 | strike         | `true` or `false`                                                                             |
|                 | outline        | `true` or `false`                                                                             |
|                 | shadow         | `true` or `false`                                                                             |
|                 | vertAlign      | `true` or `false`                                                                             |
| numFmt          |                | `"0"` // integer index to built in formats, see StyleBuilder.SSF property                     |
|                 |                | `"0.00%"` // string matching a built-in format, see StyleBuilder.SSF                          |
|                 |                | `"0.0%"` // string specifying a custom format                                                 |
|                 |                | `"0.00%;\\(0.00%\\);\\-;@"` // string specifying a custom format, escaping special characters |
|                 |                | `"m/dd/yy"` // string a date format using Excel's format notation                             |
| alignment       | vertical       | `"bottom"` or `"center"` or `"top"`                                                           |
|                 | horizontal     | `"left"` or `"center"` or `"right"`                                                           |
|                 | wrapText       | `true ` or ` false`                                                                           |
|                 | readingOrder   | `2` // for right-to-left                                                                      |
|                 | textRotation   | Number from `0` to `180` or `255` (default is `0`)                                            |
|                 |                | `90` is rotated up 90 degrees                                                                 |
|                 |                | `45` is rotated up 45 degrees                                                                 |
|                 |                | `135` is rotated down 45 degrees                                                              |
|                 |                | `180` is rotated down 180 degrees                                                             |
|                 |                | `255` is special, aligned vertically                                                          |
| border          | top            | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                  |
|                 | bottom         | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                  |
|                 | left           | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                  |
|                 | right          | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                  |
|                 | diagonal       | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                  |
|                 | diagonalUp     | `true` or `false`                                                                             |
|                 | diagonalDown   | `true` or `false`                                                                             |

**COLOR_SPEC**: Colors for `fill`, `font`, and `border` are specified as objects, either:

- `{ auto: 1}` specifying automatic values
- `{ rgb: "FFFFAA00" }` specifying a hex ARGB value
- `{ theme: "1", tint: "-0.25"}` specifying an integer index to a theme color and a tint value (default 0)
- `{ indexed: 64}` default value for `fill.bgColor`

**BORDER_STYLE**: Border style is a string value which may take on one of the following values:

- `thin`
- `medium`
- `thick`
- `dotted`
- `hair`
- `dashed`
- `mediumDashed`
- `dashDot`
- `mediumDashDot`
- `dashDotDot`
- `mediumDashDotDot`
- `slantDashDot`
