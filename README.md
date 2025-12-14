# pl-export-excel

> Export Excel files to JSON data and convert JSON data to Excel files

[文档](https://parlay96.github.io/pl-export-excel/)

# install

> npm i pl-export-excel

# Cell Styles

单元格样式设置

| 属性名              | 说明                                                   | 类型    | 默认值                        |
| ------------------- | ------------------------------------------------------ | ------- | ----------------------------- |
| patternType         | 单元格填充样式（背景填充）solid: 纯色填充 none: 无填充 | string  | `"solid"` or `"none"`         |
| fgColor             | 填充前景色: 指定十六进制 RGB 值                        | string  | -                             |
| fontFamily          | 字体                                                   | string  | 宋体                          |
| fontSize            | 字体大小                                               | number  | -                             |
| color               | 字体颜色: 指定十六进制 RGB 值                          | string  | -                             |
| isBold              | 是否加粗                                               | boolean | -                             |
| isUnderline         | 是否下划线                                             | boolean | -                             |
| isItalic            | 是否斜体                                               | boolean | -                             |
| alignmentVertical   | 垂直对齐方式                                           | string  | "bottom" or "center" or "top" |
| alignmentHorizontal | 水平对齐方式                                           | string  | "left" or "center" or "right" |
| wrapText            | 是否自动换行                                           | boolean | -                             |
| borderTop           | 上边框线，配置请看文档下面                             | obj     | -                             |
| borderBottom        | 下边框线                                               | obj     | -                             |
| borderLeft          | 左边框线                                               | obj     | -                             |
| borderRight         | 右边框线                                               | obj     | -                             |

### 边框线配置

```ts
const border[xxx] = {
  style: IBORDER_STYLE;
  /** 边框颜色: 指定十六进制 RGB 值 */
  color: string;
}
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
```
