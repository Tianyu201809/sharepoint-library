/**
 * Babel重构_typeof方法，为了兼容各个浏览器，可忽略
 */

function _typeof(obj) {
    "@babel/helpers - typeof";
    if (typeof Symbol === "function" && typeof Symbol.iterator === "symbol") {
        _typeof = function _typeof(obj) {
            return typeof obj;
        };
    } else {
        _typeof = function _typeof(obj) {
            return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj;
        };
    }
    return _typeof(obj);
}


/**
 * 使用说明：
 * SharePoint通用过滤方法
 * CAML语句拼接
 * 返回一个拼接成功的caml查询语句
 * 目前仅支持And关键字连接所有条件
 */

/**
 * 
 * @param {Array} config 查询条件对象数组 
 * @param {Object} order  排序对象
 */

/**
 * 
 * 参数详细说明：
 * config：过滤条件数组对象集合
 * order:排序对象
 * 
 */

// var config = [{
//     field: 'ID', //需要查询的字段 （SharePoint List中的字段名称）
//     fieldType: 'Text', //所查询的字段类型（与SharePoint中的Value中的Type属性一致）
//     option: 'Eq', //过滤条件（SharePoint操作符）
//     value: '123', //字段的值
//     areaData: [{
//             type: 'Integer',
//             value: 123
//         }, {
//             type: 'Text',
//             value: '12'
//         }] //当option为In时，需要填写该参数
// }, {
//     field: 'ID', //需要查询的字段
//     fieldType: 'Text', //所查询的字段类型
//     option: 'In', //过滤条件
//     value: '123', //字段的值
//     areaData: [{
//             type: 'Integer',
//             value: 123
//         }, {
//             type: 'Text',
//             value: '12'
//         }] //当option为In时，需要填写该参数, 该参数为数组对象
// }];

// var order = {
//     field: 'ID',
//     ascending: 'TRUE' //取值为'FLASE'或'TRUE'
// }


function generateQueryStr(config, order) {
    if (_typeof(config) !== 'object') return;
    var _query;
    var camlTemplate = "<Query><Where>###query###</Where></Query>";
    var len = config.length;
    if (len === 1) {
        //如果只有一个过滤条件
        debugger;
        _query = _analysisOption(config[0]);
        _query = camlTemplate.replace('###query###', _query);
    } else if (len === 2) {
        //如果有两个过滤条件
        var queryArr = [];
        var tempQuery = '';

        for (var j = 0; j < 2; j++) {
            queryArr.push(_analysisOption(config[j]));
        }

        for (var m = 0; m < queryArr.length; m++) {
            tempQuery += queryArr[m];
        }
        _query = camlTemplate.replace('###query###', '<And>' + tempQuery + '</And>');
    } else if (len >= 3) {
        //三个以上过滤条件
        debugger;
        var _queryTemp = [];
        var camlTemplate = "<Query><Where>###query###</Where></Query>";

        for (var i = 0; i < config.length; i++) {
            _queryTemp.push(_analysisOption(config[i]));
        }
        //首先将前两个过滤条件进行拼接
        var _queryPart = '<And>' + _queryTemp[0] + _queryTemp[1] + '</And>' + '###query###'; //此处的###query###是为了给后面的条件进行占位
        var _query = camlTemplate.replace('###query###', _queryPart); //接下来按照规律进行And拼接
        var _tempQuery, _tempQuery2;
        for (var n = 2; n < config.length; n++) {
            var queryFilter = _analysisOption(config[n]); //获取过滤条件
            _tempQuery = _query.replace('<Where>', '<Where><And>');
            _tempQuery2 = _tempQuery.replace('###query###', queryFilter + '</And>' + '###query###');
            _query = _tempQuery2;
        }
        _query = _tempQuery2.replace(/###query###/g, ''); //将最后多余的占位符###query###删除掉
    }
    if (order) {
        //<OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy>
        //默认排序字段
        return _query.replace('</Where>', '</Where>' + '<OrderBy>' + "<FieldRef Name='" + order.field + "' Ascending='" + order.ascending ? (order.ascending) : 'FALSE' + "'/>" + '</OrderBy>')

    } else {
        return _query;
    }
}


/**
 * 根据config数组对象中的对象，生成条件查询字符串
 * @param {*Object}
 */
function _analysisOption(obj) {
    //所有包含的参数
    // var optArr = [
    //     "Eq", //等于
    //     "Neq", //不等于
    //     "Lt", //小于
    //     "Leq", //小于等于
    //     "Gt", //大于
    //     "Geq", //大于等于
    //     "Contains", //包含
    //     "BeginsWith", //以某字符串开头
    //     "In", //在集合范围内
    //     "IsNull", //为空
    //     "IsNotNull" //不为空
    // ];
    var _query = '';
    switch (obj.option) {
        case 'Eq':
            _query += "<Eq><FieldRef Name='" + obj.field + "' /><Value Type='" + obj.fieldType + "'>" + obj.value + "</Value></Eq>";
            break;
        case 'Neq':
            _query += "<Neq><FieldRef Name='" + obj.field + "' /><Value Type='" + obj.fieldType + "'>" + obj.value + "</Value></Neq>";
            break;
        case 'Lt':
            _query += "<Lt><FieldRef Name='" + obj.field + "' /><Value Type='" + obj.fieldType + "'>" + obj.value + "</Value></Lt>";
            break;
        case 'Leq':
            _query += "<Leq><FieldRef Name='" + obj.field + "' /><Value Type='" + obj.fieldType + "'>" + obj.value + "</Value></Leq>";
            break;
        case 'Gt':
            _query += "<Gt><FieldRef Name='" + obj.field + "' /><Value Type='" + obj.fieldType + "'>" + obj.value + "</Value></Gt>";
            break;
        case 'Geq':
            _query += "<Geq><FieldRef Name='" + obj.field + "' /><Value Type='" + obj.fieldType + "'>" + obj.value + "</Value></Geq>";
            break;
        case 'Contains':
            _query += "<Contains><FieldRef Name='" + obj.field + "' /><Value Type='" + obj.fieldType + "'>" + obj.value + "</Value></Contains>";
            break;
        case 'BeginsWith':
            _query += "<BeginsWith><FieldRef Name='" + obj.field + "' /><Value Type='" + obj.fieldType + "'>" + obj.value + "</Value></BeginsWith>";
            break;
        case 'In':
            //In条件在一定的范围内检索数据
            //需要在输入参数中添加areaData属性
            var _queryTemp = "<Values>###query###</Values>";
            var _queryValue = '';
            for (var i = 0; i < obj.areaData.length; i++) {
                _queryValue += "<Value Type='" + obj.areaData[i].type + "'>" + obj.areaData[i].value + "</Value>";
            }
            _query = _queryTemp.replace('###query###', _queryValue);
            break;
        case 'IsNull':
            _query += "<IsNull><FieldRef Name='" + obj.field + "' /><Value Type='" + obj.fieldType + "'>" + obj.value + "</Value></IsNull>";
            break;
        case 'IsNotNull':
            _query += "<IsNotNull><FieldRef Name='" + obj.field + "' /><Value Type='" + obj.fieldType + "'>" + obj.value + "</Value></IsNotNull>";
            break;
        default:
            break;
    }

    return _query;
}