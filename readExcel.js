/**
 * 读取excel文件插件
 * 前提条件：引入xlsx类库文件：github说明网址：https://github.com/SheetJS/sheetjs/blob/master/README.md
 * 仅需引入 dist/xlsx.core.min.js
 */

/**
 * 把文件按照二进制进行读取(异步)
 * @param {*file} 所获取的dom上的文件 
 */
function _readFile(file) {
    return new Promise(function(resolve) {
        var reader = new FileReader();
        reader.readAsArrayBuffer(file);
        reader.onload = function(ev) {
            resolve(ev.target.result)
        }
    })
}

/**
 * 读取文件信息，把读取到的文件转化成服务器可以解析的json数据
 * file是所需读取的excel文件
 * _keyMapRule是提前定义好的映射转换规则
 * sheetName是sheet页名称, 如果不传递，则默认读取第一个sheet页的名称
 * _keyMap示例：（如：name对应excel文件中的姓名列，phone对应excel文件中的电话列）
 * 
        // {
        //  name:{
        //   text:"姓名",
        //   type:'string'
        // },
        // phone:{
        //  text:"电话",
        //  type:"string"
        // }
        // }
 * 
 * 
 */

function Excle2Json(file, _keyMapRule, sheetName) {
    return new Promise(function(resolve, reject) {
        _readFile(file).then(function(data) {
            if (window.XLSX) {
                var xlsx = window.XLSX;
            } else {
                reject("没有引入js-xlsx库，请到官网下载并引入:https://github.com/SheetJS/sheetjs/blob/master/dist/xlsx.core.min.js")
            }
            var workbook = xlsx.read(data, { type: 'buffer' }),
                sheetName = sheetName ? sheetName : workbook.SheetNames[0];
            worksheet = workbook.Sheets[sheetName]; //默认读取第一个sheet页的数据
            //此时获取到的数据是data是原生的数据，如果excel是中文的话，我们转换出来的json数据的属性名也是中文，所以需要将其属性名转化为英文
            data = xlsx.utils.sheet_to_json(worksheet);
            //转化属性名
            //前提条件，首先在全局定义一个对象结构，该对象结构用来设定转换规则
            if (!_keyMapRule) {
                //如果没有传入映射规则，则不进行属性名转换
                resolve(data)
            } else {
                //进行属性名转换
                var arr = [];
                data.forEach(function(item) {
                    var obj = {};
                    for (var key in _keyMapRule) {
                        if (!_keyMapRule.hasOwnProperty(key)) break;
                        var v = _keyMapRule[key],
                            text = v.text,
                            type = v.type;
                        v = item[text] || "";
                        type === "string" ? (v = String(v)) : null;
                        type === "number" ? (v = Number(v)) : null;
                        obj[key] = v;
                    }
                    arr.push(obj);
                });
                resolve(arr);
            }
        })
    })
}