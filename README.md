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
* promise.js
* vue.min.js
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
2020-7-27 
* 新增防抖与节流工具函数,位于文件debounce.js中
* 新增生成caml查询字符串函数generateQueryStr，位于SPfilter-ES5.js文件中

* 使用说明：
* 为generateQueryStr函数传递两个参数，config和order
* 参数示例：
var config = [{
    field: 'ID', //需要查询的字段
    fieldType: 'Text', //所查询的字段类型
    option: 'Eq', //过滤条件
    value: '123', //字段的值
    areaData: [{
            type: 'Integer',
            value: 123
        }, {
            type: 'Text',
            value: '12'
        }] //当option为In时，需要填写该参数
}, {
    field: 'ID', //需要查询的字段
    fieldType: 'Text', //所查询的字段类型
    option: 'In', //过滤条件
    value: '123', //字段的值
    areaData: [{
            type: 'Integer',
            value: 123
        }, {
            type: 'Text',
            value: '12'
        }] //当option为In时，需要填写该参数, 该参数为数组对象
}];

var order = {
    field: 'ID',
    ascending: 'TRUE' //取值为'FLASE'或'TRUE'
}

* generateQueryStr(config, order)

-----------------------------------------------------------------------------
2020-7-30
* 新增文件异步上传代码重构方法
* 使用方法：

配合KendoUI的onUpload组件对象，给onUpload属性设置方法：onUploadFiles
如果没有使用kendoUI, 则需要给方法传递input type='file' 控件的id值
function onUploadFiles(e) {
    //默认传入对象是kendoUI的事件对象
    if (typeof e === 'object') {
        //此时默认为kendoui的事件对象
        var files = e.files;
        _uploadFilesCommonAsync(files, listName, listItemID).then(function(message) {
        /**
         * 此处请设置 listName, listItemID的参数值
         */
            console.log(message);
        }).catch(function(e) {
            console.log(e)
        })
    } else if (typeof e === 'string') {
        //获取上传控件id，根据id获取目前上传的附件
        var files = $("#" + e)[0].files; //获取所上传的文件类数组对象
        _uploadFilesCommonAsync(files, listName, listItemID).then(function(message) {
        /**
         * 此处请设置 listName, listItemID的参数值
         */
            console.log(message);
        }).catch(function(e) {
            console.log(e)
        })
    } else {
        //不能上传
        return false;
    }
}

* 注意：_uploadFilesCommonAsync该方法是上传文件到SharePoint服务器中的方法，需要传递一些参数，这些参数可以在onUploadFiles函数块中自己通过代码去获取，比如listName, listItemId等参数

* 示例
        $("#uploadFile").kendoUpload(
            {
                template: $("#fileTemplate").html(),
                async:
                {
                    saveUrl: _spPageContextInfo.webAbsoluteUrl + "/save",
                    removeUrl: _spPageContextInfo.webAbsoluteUrl + "/remove",
                    autoUpload: false
                },
                files: allfiles,
                ·upload: onUploadFiles·, //绑定上传附件方法：onUploadFiles
                validation: {
                    allowedExtensions: allowedExtensionsArray,
                    maxFileSize: filemaxsize,
                    minFileSize: 0
                },
                success: onSuccess,
                error: onError,
                complete: uploadFileComplete,
                select: onSelectSA_UploadFile
            });