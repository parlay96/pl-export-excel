# pl-export-excel

> Export Excel files to JSON data and convert JSON data to Excel files

# method of use

> npm i pl-export-excel

# exportJsonToExcel Api

|        参数名        | 参数描述                                                | 必填 |   类型   |    默认值    |
| :------------------: | ------------------------------------------------------- | :--: | :------: | :----------: |
|       headers        | 头部                                                    |  是  |  Array   |      []      |
|        datas         | 表格数据                                                |  是  |  Array   |      []      |
|     titleConfig      | 标题配置, 配置描述请看文档下面                          |  否  |  object  |      -       |
|     multiHeader      | 多级头, 给表格设置多行头部,需要自己去合并单元格。       |  否  |  Array   |      -       |
|       filename       | 表格名                                                  |  否  |  String  | 'excel-list' |
|        merges        | 合并单元格数组, 配置描述请看文档下面                    |  否  |  Array   |      -       |
|      autoWidth       | 宽度是否自适应                                          |  否  | Boolean  |     true     |
|       bookType       | 文件类型                                                |  否  |  String  |    'xlsx'    |
|       styleCb        | 样式[style](#cell-styles)                               |  否  | Function |      -       |
| xlsxStyleResourceUrl | 自定义xlsx样式资源地址，基本用不到，除非你想替换资源cdn |  否  |  String  |      -       |
|         keys         | 需要导出表格数据中的字段key名, 不配置默认导出所有字段   |  否  |  Array   |      -       |

# npm

```javascript
   import { exportJsonToExcel } from 'pl-export-excel'
   // 导出按钮方法
   handleEmits () {
    // 表格数据
    const list = Array.from({ length: 200 }, (_, idx) => ({
      names: "娃哈哈",
      date: "201920120",
      orderNumber: "1521",
      customerName: "王小虎",
      orderState: "在线",
      orderPayState: "全付款"
    }));
    exportJsonToExcel({
      headers: ["经销商名称", "下单时间", "订单编号", "客户名称", "订单状态", "付款状态"],
      datas: list,
      filename: "订单表格"
    });
   }
```

# cdn

```javascript
/**
 * 引入pl-export-excel
   <body>
     <script src="https://unpkg.com/pl-export-excel@1.1.6/dist/index.js"></script>;
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
export type ISubTitle = {
  /** 标题文本 */
  title: string | number;
  /** 需要合并行数 */
  rowNum: number;
  /** 需要合并列数-默认为headers字段的长度 */
  colNum?: number;
};
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
