说明：
------------------------------------------------------------------------------
2020-06-02
* 本类库是sharepoint工具类库文件:
* common.js封装了对于SharePoint列表的CRUD操作（包含同步和异步两种），以及一些工具方法
------------------------------------------------------------------------------
2020-06-02
lib 文件夹是平时开发中所用到的一些类库
* jquery-3.1.0.min.js
* jquery-ui.js
* jquery.SPServices.min.js
* kendo.all.min.js
* moment.js
* xlsx.core.min.js
------------------------------------------------------------------------------
2020-7-3新增读取excel文件数据的方法readExcel.js
使用方法：
直接调用函数Excel2Json(file, _keyMapRule, sheetName)
 * file是所需读取的excel文件
 * _keyMapRule是提前定义好的映射转换规则
 * sheetName是sheet页名称, 如果不传递，则默认读取第一个sheet页的名称
 * _keyMap示例：（如：name对应excel文件中的姓名列，phone对应excel文件中的电话列）
         ·{
          name:{
           text:"姓名",
           type:'string'
         },
         phone:{
          text:"电话",
          type:'string'
         }
        }·
-----------------------------------------------------------------------------