/**
 * 文件导出工具类
 * 使用前提：引入xlsx.js
 * 官方github网址: https://github.com/SheetJS/sheetjs
 * 中文教程:https://www.cnblogs.com/liuxianan/p/js-excel.html 
 * 
 * 使用方法:直接调用export2Excel(array,filename)方法
 * array是需要显示的数据的数组集合，array[0]是“标题”， array[1]~array[n]是具体的数据
 * array参数示例 => [["Title","Name","Age"],["title1","Bob",12],["title2","Pater",18],...]
 */

function export2Excel(array, filename) {
    var sheet = XLSX.utils.aoa_to_sheet(array);
    if (!filename) filename = 'sla-report-list.xlsx';
    openDownloadDialog(sheet2blob(sheet), filename);
}

/**
 * 打开文件导出window窗口
 * @param {*} url 
 * @param {*} saveName 
 */
function openDownloadDialog(url, saveName) {
    if (window.navigator && window.navigator.msSaveOrOpenBlob) {
        //IE浏览器导出方法
        window.navigator.msSaveOrOpenBlob(url, saveName);
    } else {
        //非IE浏览器导出方法
        if (typeof url == 'object' && url instanceof Blob) {
            url = URL.createObjectURL(url); // 创建blob地址
        }
        var $a = document.createElement('a');
        $a.setAttribute("href", url);
        $a.setAttribute("download", saveName);
        $a.setAttribute("target", "_blank");//弹出窗体
        var evObj = document.createEvent('MouseEvents');
        evObj.initMouseEvent('click', true, true, window, 0, 0, 0, 0, 0, false, false, true, false, 0, null);
        $a.dispatchEvent(evObj);
    }
}

/**
 * 输出blob
 * @param {*} sheet 
 * @param {*} sheetName 
 */
function sheet2blob(sheet, sheetName) {
    sheetName = sheetName || 'sheet1';
    var workbook = {
        SheetNames: [sheetName],
        Sheets: {}
    };
    workbook.Sheets[sheetName] = sheet;
    // 生成excel的配置项
    var wopts = {
        bookType: 'xlsx', // 要生成的文件类型
        bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);
    var blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    // 字符串转ArrayBuffer
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    return blob;
}