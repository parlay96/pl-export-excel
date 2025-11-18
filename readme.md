# pl-export-excel（当前版本 1.1.3）
> 一个导出表格的插件

# method of use
> npm i pl-export-excel

# pl-export-excel方法

方法名  |  说明  |  参数 | 参数表述
:----------: | -------  |  :-------: |  :-------:
formatJson | 格式化数据 | filterVal, jsonData |  filterVal 表头在数据中的对应key数组; jsonData 需要导出表格数据
exportExcelMethod  |  通过table ID渲染导出表格（不能是组件，必须是手写的table表格）  |  tableId, fileName, sheetName |  tableId 需要导出表格的ID; fileName 文件名; sheetName 表名
exportJsonToExcel |  通过json渲染导出表格（常用） | 见下面的表 | 见下面的表
exportTableToExcel | 通过table标签渲染导出表格（常用，可以是el-table,也可是pl-table） | 见下面的表 | 见下面的表


# exportTableToExcel方法参数描述

参数名  |  参数描述  |  必填  |  类型  |  默认值
:----------: | -------  |  :-------:  |  :-------: |  :-------:
id | 需要导出表格的ID |  是 | String |  —
filename | 表格名 | 否 | String | '空'
bookType | 文件类型 | 否 | String | 'xlsx'
styleFun | 样式函数方法 | 否 | Function | styleFun(参数是当前表格ws)


# exportJsonToExcel方法参数描述

参数名  |  参数描述  |  必填  |  类型  |  默认值
:----------: | -------  |  :-------:  |  :-------: |  :-------:
title | 标题 |  否 | Array |  []
multiHeader | 多级头 | 否 | Array | []
header | 头部 | 否 | Array | []
data | 表格数据 | 否 | Array | []
filename | 表格名 | 否 | String | ''
merges | 合并单元格数组 | 否 | Array | []
autoWidth | 宽度是否自适应 | 否 | Boolean | true
bookType | 文件类型 | 否 | String | 'xlsx'
styleFun | 样式函数方法 | 否 | Function | styleFun(参数是当前表格ws)


# npm方式（用法）
**实例**
``` javascript
   // 第一步
   在入口文件的index.html，引入XLSX样式
   <script src="https://unpkg.com/pl-export-excel@1.1.3/package/xlsx.core.min.js"></script>
   // 第二步,在项目中的使用
   import { exportJsonToExcel, formatJson } from 'pl-export-excel'
   // 导出按钮方法
   handleEmits () {
        // 表头
        const tHeader = ['经销商名称', '下单时间', '订单编号', '客户名称', '订单状态', '付款状态']
        // 表头在数据中的对应key
        const filterVal = ['names', 'date', 'orderNumber', 'customerName', 'orderState', 'orderPayState']
        // 表格数据
        const list = Array.from({ length: 1000 }, (_, idx) => ({ names: '娃哈哈',
            date: '201920120',
            orderNumber: '1521',
            customerName: '王小虎',
            orderState: '在线',
            orderPayState: '全付款'
        }))
        // 导出的数据
        const data = formatJson(filterVal, list)
        // 导出表格
        exportJsonToExcel({
          header: tHeader,
          data: data,
          merges: [{
              s: {//s为开始
                c: 0,//开始列
                r: 0//可以看成开始行,实际是取值范围
              },
              e: {//e结束
                c: 5,//结束列
                r: 1//结束行
              }
          }],
          multiHeader: [
            ["工作情况一览表", "", "", "", "", ""],
            ["", "", "", "", "", ""] // 为啥这里需要这样搞个空字符呢，存属上面合并列导致不得不这样写个哦
          ],
          filename: 'erp订单',
          bookType: 'xlsx',
          // 不懂怎么设置ws,看https://github.com/protobi/js-xlsx/tree/beta#cell-object
          styleFun: function (ws) {
            ws["A1"].s = {
              font: {
                name: '宋体',
                sz: 18,
                color: {rgb: "ff0000"},
                bold: true,
                italic: false,
                underline: false
              },
              alignment: {
                horizontal: "center",
                vertical: "center"
              },
              fill: {
                fgColor: {rgb: "008000"},
              },
            };
          }
        })
   }

> 注意如果你不需要 pl-export-excel参与打包(减少打包体积)
  // 第一步: 请在webpack配置
  externals: {
    'pl-export-excel': 'PlExportExcel'
  }
  // 第二步:  在入口文件的index.html
  // 引入pl-export-excel
  <script src="https://unpkg.com/pl-export-excel@1.1.3/lib/index.js"></script>
  // 引入XLSX样式
  <script src="https://unpkg.com/pl-export-excel@1.1.3/package/xlsx.core.min.js"></script>
```

# cdn方式用法
**实例**
``` javascript
  // 在html页面引入：
  <body>
    <div>我是内容</div>
    在这里引入脚本
    // 引入pl-export-excel
    <script src="https://unpkg.com/pl-export-excel@1.1.3/lib/index.js"></script>
     // 引入XLSX样式
    <script src="https://unpkg.com/pl-export-excel@1.1.3/package/xlsx.core.min.js"></script>
  </body>

  // 写法
  <body>
      <div id="myApp">
          我是按钮
      </div>
      <script>
          var btn = document.getElementById('myApp')
          btn.onclick = handleEmits
          // 导出按钮
          function handleEmits () {
              // 表头
              const tHeader = ['经销商名称', '下单时间', '订单编号', '客户名称']
              // 表头在数据中的对应key
              const filterVal = ['names', 'date', 'orderNumber', 'customerName']
              // 表格数据
              const list = Array.from({length: 10000}, () => ({
                  date: '2016-05-03',
                  name: '王小虎',
                  address: '上海市普陀区金沙江路 1516 弄'
              }))
              // 导出的数据
              const data = PlExportExcel.formatJson(filterVal, list)
              // 导出表格
              PlExportExcel.exportJsonToExcel({
                  header: tHeader,
                  data: data,
                  filename: '经销商表格',
                  autoWidth: true,
                  bookType: 'xlsx'
              })
          }
      </script>
  </body>
```

