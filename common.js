/**
 * 说明：
 * 工具方法类库
 * 增删改查SP list
 * sync代表同步方法，async代表异步方法
 * 使用此工具类库的前置条件为引用jquery(3.0)以上版本和引用SPService类库，还有Promise类库（如果浏览器不支持ES6需要引入）
 * version 1.0  github地址: https://github.com/Tianyu201809/sharepoint-library/tree/dev
 * 作者: Tianyu Zhang
 * 时间: 2020-06-01
 */


/**
 * 获取sharepoint list数据的同步函数
 * 注意填写arrayField参数时，list的显示字段和技术字段的名称要保持一致
 * @param {*string} listName 所查询列表的名称
 * @param {*string} query   查询条件CAML语法 "<Query><Where></Where></Query>"
 * @param {*array => ["Title","ID"...]} arrayField   所需要查询的字段，如果不填则查询所有字段 ["Title","ID"...]
 */
function getListDataSync(listName, query, arrayField) {
    var data = [];
    if (!listName) {
        return "调用getListDataAsync时，请填写listName!"
    }
    if (!arrayField) {
        return "调用getListDataAsync时，请输入查询字段列表['Title','ID'...]";
    }
    if (!query) {
        query = '<Query><Where></Where></Query>';
    }

    var _viewFields = "<ViewFields>";
    for (var k = 0; k < arrayField.length; k++) {
        _viewFields += "<FieldRef Name='" + arrayField[k] + "' />";
    }
    _viewFields += "</ViewFields>";
    $().SPServices({
        operation: 'GetListItems',
        async: false,
        listName: listName,
        CAMLViewFields: _viewFields,
        CAMLQuery: query,
        completefunc: function (xData, Status) {
            if ($(xData.responseXML).SPFilterNode("z:row").length > 0) {
                $(xData.responseXML).SPFilterNode("z:row").each(function (i, val) {
                    for (var j = 0; j < arrayField.length; j++) {
                        var key = String(arrayField[j]);
                        data[i] ? data[i] : data[i] = {};
                        data[i][key] = $(this).attr("ows_" + arrayField[j] + "") || "";
                    }
                });
            } else {
                var err = {};
                err['response'] = xData.responseXML;
                err['status'] = "error";
                data = [];
                data.push(err);
            }
        }
    });
    return data;
}


/**
 * 获取sharepoint list数据的异步函数，该函数会返回一个Promise对象
 * 注意填写arrayField参数时，list的显示字段和技术字段的名称要保持一致
 * @param {*string} listName 所查询列表的名称
 * @param {*string} query   查询条件CAML语法 "<Query><Where></Where></Query>"
 * @param {*array => ["Title","ID"...]} arrayField   所需要查询的字段，如果不填则查询所有字段 
 */
function getListDataAsync(listName, query, arrayField) {
    return new Promise(function (resolve, reject) {
        if (!listName) {
            reject("调用getListDataAsync时，请填写listName!");
            return;
        }
        if (!arrayField) {
            reject("调用getListDataAsync时，请输入查询字段列表['Title','ID'...]");
            return;
        }
        if (!query) {
            query = '<Query><Where></Where></Query>';
        }

        var _viewFields = "<ViewFields>";
        for (var k = 0; k < arrayField.length; k++) {
            _viewFields += "<FieldRef Name='" + arrayField[k] + "' />";
        }
        _viewFields += "</ViewFields>";
        $().SPServices({
            operation: 'GetListItems',
            async: true,
            listName: listName,
            CAMLViewFields: _viewFields,
            CAMLQuery: query,
            completefunc: function (xData, Status) {
                if ($(xData.responseXML).SPFilterNode("z:row").length > 0) {
                    var data = [];
                    $(xData.responseXML).SPFilterNode("z:row").each(function (i, val) {
                        for (var j = 0; j < arrayField.length; j++) {
                            var key = String(arrayField[j]);
                            data[i] ? data[i] : data[i] = {};
                            data[i][key] = $(this).attr("ows_" + arrayField[j] + "") || "";
                        }
                    });
                    resolve(data);
                } else {
                    var err = {};
                    err['response'] = xData.responseXML;
                    err['status'] = "error";
                    reject(err);
                }
            }
        });
    })
}

/**
 * 为sharepoint添加数据，同步函数
 * @param {*string} listName list名称
 * @param {*string} data  所添加的数据 [['Title','12345'],['desc','hello world'] ...]
 */

function insertDataIntoListSync(listName, data) {
    var itemID;
    var obj = {};
    $().SPServices({
        operation: 'UpdateListItems',
        async: false,
        batchCmd: 'New',
        listName: listName,
        valuepairs: data,
        completefunc: function (xData, Status) {
            if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
                itemID = $(xData.responseXML).SPFilterNode("z:row").attr("ows_ID");
                obj['ID'] = itemID;
                obj['response'] = null;
                obj['status'] = "success";
            } else {
                obj['ID'] = undefined;
                obj['response'] = xData.responseXML;
                obj['status'] = "error";
            }
        }
    });
    return obj; //返回数据添加成功后，item的id值
}


/**
 * 为sharepoint添加数据，异步函数，返回promise实例
 * @param {*string} listName list名称 必填参数
 * @param {*string} data  所添加的数据  必填参数 eg :[['Title',"hello"],['field1','test',],...]
 */
function insertDataIntoListAsync(listName, data) {
    return new Promise(function (resolve, reject) {
        if (!listName) {
            return Promise.reject();
        }
        $().SPServices({
            operation: 'UpdateListItems',
            async: true,
            batchCmd: 'New',
            listName: listName,
            valuepairs: data,
            completefunc: function (xData, Status) {
                if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
                    var obj = {};
                    var itemID = $(xData.responseXML).SPFilterNode("z:row").attr("ows_ID");
                    obj['ID'] = itemID;
                    obj['response'] = null;
                    obj['status'] = "success";
                    resolve(obj);
                } else {
                    var err = {};
                    err['ID'] = undefined;
                    err['response'] = xData.responseXML;
                    obj['status'] = "error";
                    reject(err);
                }
            }
        });
    })
}
/**
 * 删除item 同步方法，返回一个obj对象
 */

function delListItemSync(listName, itemID) {
    var obj = {};
    $().SPServices({
        operation: 'UpdateListItems',
        async: false,
        batchCmd: 'Delete', //New, Update, Delete, Moderate
        listName: listName,
        ID: itemID, //
        completefunc: function (xData, Status) {
            if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
                obj['status'] = "success";
                obj['response'] = 'ID:' + itemID + " deleted success";
            } else {
                obj['status'] = "error";
                obj['response'] = xData.responseXML;
            }
        }
    });
    return obj;
}

/**
 * 删除item 异步方法，返回Promise对象
 * @param {*string} listName SP列表名称
 * @param {*string} itemID  SP List item ID值
 */
function delListItemAsync(listName, itemID) {
    return new Promise(function (resolve, reject) {
        $().SPServices({
            operation: 'UpdateListItems',
            async: true,
            batchCmd: 'Delete', //可以包含的参数: New, Update, Delete, Moderate
            listName: listName,
            ID: itemID,
            completefunc: function (xData, Status) {
                if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
                    var obj = {};
                    obj['status'] = "success";
                    obj['response'] = 'ID:' + itemID + " deleted success";
                    resolve(obj);
                } else {
                    var err = {};
                    err['status'] = "error";
                    err['response'] = xData.responseXML;
                    reject(err);
                }
            }
        });
    })
}

/**
 * 更新SP List item 同步函数
 * @param {*string} listName 必填
 * @param {*string} itemID  必填
 * @param {*array => [['Title','123'],['field1','123'],['field2','123']...]} 必填
 */
function updateListItemSync(listName, itemID, data) {
    if (!listName) {
        return "Please input listname";
    }
    if (!itemID) {
        return "Please input itemID";
    }
    for (var i = 0; i < data.length; i++) {
        if (Object.prototype.toString.call(data[i]).indexOf('Array') === -1) {
            return "Please input data (Array format)";
        }
    }
    var obj = {};
    $().SPServices({
        operation: "UpdateListItems",
        async: false,
        listName: listName,
        ID: itemID,
        valuepairs: data,
        completefunc: function (xData, Status) {
            if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
                obj['status'] = "success";
                obj['response'] = 'ID:'+ itemID + " updated success";
                obj['ID'] = itemID;
            } else {
                obj['status'] = "error";
                obj['response'] = xData.responseXML;
                obj['ID'] = itemID;
            }
        }
    });
    return obj;
}

/**
 * 更新SP List item 异步函数，返回一个Promise对象
 * @param {*string} listName 必填
 * @param {*string} itemID  必填
 * @param {*array => [['Title','123'],['field1','123'],['field2','123']...]} 必填
 */
function updateListItemAsync(listName, itemID, data) {
    return new Promise(function (resolve, reject) {
        if (!listName) {
            return Promise.reject("Please input listname");
        }
        if (!itemID) {
            return Promise.reject("Please input itemID");
        }
        for (var i = 0; i < data.length; i++) {
            if (Object.prototype.toString.call(data[i]).indexOf('Array') === -1) {
                return Promise.reject("Please input data (Array format)");
            }
        }
        $().SPServices({
            operation: "UpdateListItems",
            async: true,
            listName: listName,
            ID: itemID,
            valuepairs: data,
            completefunc: function (xData, Status) {
                if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
                    var obj = {};
                    obj['status'] = "success";
                    obj['response'] = 'ID:'+ itemID + " updated success";
                    obj['ID'] = itemID;
                    resolve(obj);
                } else {
                    var obj = {};
                    obj['status'] = "error";
                    obj['response'] = xData.responseXML;
                    obj['ID'] = itemID;
                    reject(obj);
                }
            }
        });
    })
}


/**
 * 日期格式化，同步函数
 * 注意：此方法修改了Date原型
 * 
 * 调用方法
 * var time1 = new Date().format("yyyy-MM-dd HH:mm:ss");     
 * var time2 = new Date().format("yyyy-MM-dd");  
 */
Date.prototype.format = function (fmt) { //author: meizz   
    var o = {
        "M+": this.getMonth() + 1,               //月份   
        "d+": this.getDate(),                    //日   
        "h+": this.getHours(),                   //小时   
        "m+": this.getMinutes(),                 //分   
        "s+": this.getSeconds(),                 //秒   
        "q+": Math.floor((this.getMonth() + 3) / 3), //季度   
        "S": this.getMilliseconds()             //毫秒   
    };
    if (/(y+)/.test(fmt))
        fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    for (var k in o)
        if (new RegExp("(" + k + ")").test(fmt))
            fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
    return fmt;
}

/**
 * 日期类型的数据传递给SP后端的时候，需要进行格式化操作
 * 该方法为通用的日期转换方法
 */

function ConvertDateISO(dateVal) {
    var result = $().SPServices.SPConvertDateToISO({
        dateToConvert: new Date(dateVal),
        dateOffset: "-05:00"
    });
    return result;
}


/**
 * 获取url中传递的参数
 * 输出一个对象obj
 */
function getUrlVars() {
    var curParams = document.location.search;
    var vars = [], hash;
    var hashes;
    if (curParams.split('?').length > 2) {
        hashes = curParams.substring(curParams.lastIndexOf('?') + 1).split('&');
    }
    else {
        hashes = curParams.substr(1).split('&');
    }
    for (var i = 0; i < hashes.length; i++) {
        hash = hashes[i].split('=');
        vars.push(hash[0]);
        vars[hash[0]] = hash[1];
    }
    return vars;
}

/**
 * 获取用户所包含的Group，异步函数，返回一个Promise对象
 * 刚函数输出一个数组array， 数组中包含所查询用户在当前站点下的Group权限组
 * 不填写用户参数，则默认为当前用户
 */
//获取用户权限
function getUserGroups(username) {
    username ? username : username = $().SPServices.SPGetCurrentUser();
    return new Promise(function (resolve, reject) {
        var userInGroup = [];
        $().SPServices({
            operation: "GetGroupCollectionFromUser",
            userLoginName: username,
            async: true,
            completefunc: function (xData, Status) {
                if ($(xData.responseXML).SPFilterNode("Group").length > 0) {
                    $(xData.responseXML).SPFilterNode("Group").each(function () {
                        userInGroup.push($(this).attr("Name") || "");
                    });
                    resolve(userInGroup);
                } else {
                    reject([]); //获取权限失败
                }

            }
        });

    })
}
/**
 * 对象克隆方法（深度克隆）
 */
function cloneObj(obj) {
    return JSON.parse(JSON.stringify(obj))
}
