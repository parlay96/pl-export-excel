# pl-export-excel

> Export Excel files to JSON data and convert JSON data to Excel files

# method of use

> npm i pl-export-excel

# exportJsonToExcel Api

|        参数名        | 参数描述                      | 必填 |   类型   |           默认值           |
| :------------------: | ----------------------------- | :--: | :------: | :------------------------: |
|        header        | 头部                          |  是  |  Array   |             []             |
|        datas         | 表格数据                      |  是  |  Array   |             []             |
|     multiHeader      | 多级头                        |  否  |  Array   |             []             |
|       filename       | 表格名                        |  否  |  String  |             ''             |
|        merges        | 合并单元格数组                |  否  |  Array   |             []             |
|      autoWidth       | 宽度是否自适应                |  否  | Boolean  |            true            |
|       bookType       | 文件类型                      |  否  |  String  |           'xlsx'           |
|       styleCb        | 样式函数方法                  |  否  | Function | styleFun(参数是当前表格ws) |
| xlsxStyleResourceUrl | 自定义xlsx样式资源地址        |  否  |  String  |             -              |
|         keys         | 需要导出表格数据中的字段key名 |  否  |  Array   |             -              |

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
      header: ["经销商名称", "下单时间", "订单编号", "客户名称", "订单状态", "付款状态"],
      datas: list,
      multiHeader: [
        ["工作情况一览表", "", "", "", "", ""],
        ["", "", "", "", "", ""] // 为啥这里需要这样搞个空字符呢，存属上面合并列导致不得不这样写个哦
      ],
      merges: [
        {
          s: {
            //s为开始
            c: 0, //开始列
            r: 0 //可以看成开始行,实际是取值范围
          },
          e: {
            //e结束
            c: 5, //结束列
            r: 1 //结束行
          }
        }
      ],
      filename: "erp订单",
      bookType: "xlsx",
      styleCb: function (ws) {
        ws["A1"].s = {
          font: {
            name: "宋体",
            sz: 18,
            color: { rgb: "ff0000" },
            bold: true,
            italic: false,
            underline: false
          },
          alignment: {
            horizontal: "center",
            vertical: "center"
          },
          fill: {
            fgColor: { rgb: "008000" }
          }
        };
      }
    });
   }
```

# cdn

```javascript
<body>
  <div>我是内容</div>
  // 引入pl-export-excel
  <script src="https://unpkg.com/pl-export-excel@1.1.6/dist/index.js"></script>
</body>
```
