/**
 * 说明：
 * 工具方法类库
 * 增删改查SP list
 * sync代表同步方法，async代表异步方法
 * 使用此工具类库的前置条件为引用jquery(3.0)以上版本和引用SPService类库，还有Promise类库（如果浏览器不支持ES6需要引入）
 * version 1.2  github地址: https://github.com/Tianyu201809/sharepoint-library/tree/dev
 * 作者: Tianyu Zhang
 * 时间: 2020-06-02
 */


/**
 * 获取sharepoint list数据的同步函数
 * 注意填写arrayField参数时，list的显示字段和技术字段的名称要保持一致
 * @param {*string} listName 所查询列表的名称  必填
 * @param {*string} query   查询条件CAML语法 "<Query><Where></Where></Query>"
 * @param {*array => ["Title","ID"...]} arrayField    ["Title","ID"...]  必填
 */
function getListDataSync(listName, query, arrayField) {
    var data = [];
    if (!listName) {
        return false;
    }
    if (!arrayField) {
        return false;
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
        completefunc: function(xData, Status) {
            if ($(xData.responseXML).SPFilterNode("z:row").length > 0) {
                $(xData.responseXML).SPFilterNode("z:row").each(function(i, val) {
                    for (var j = 0; j < arrayField.length; j++) {
                        var key = String(arrayField[j]);
                        data[i] ? data[i] : data[i] = {};
                        data[i][key] = $(this).attr("ows_" + key + "") || "";
                    }
                });
            } else {
                var err = {};
                err['response'] = xData.responseText;
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
 * @param {*array => ["Title","ID"...]} arrayField  必填
 */
function getListDataAsync(listName, query, arrayField) {
    return new Promise(function(resolve, reject) {
        if (!listName) {
            reject(false);
            return;
        }
        if (!arrayField) {
            reject(false);
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
            completefunc: function(xData, Status) {
                if ($(xData.responseXML).SPFilterNode("z:row").length > 0) {
                    var data = [];
                    $(xData.responseXML).SPFilterNode("z:row").each(function(i, val) {
                        for (var j = 0; j < arrayField.length; j++) {
                            var key = String(arrayField[j]);
                            data[i] ? data[i] : data[i] = {};
                            data[i][key] = $(this).attr("ows_" + key + "") || "";
                        }
                    });
                    resolve(data);
                } else {
                    var err = {};
                    err['response'] = xData.responseText;
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
    if (!listName) {
        return false;
    }
    $().SPServices({
        operation: 'UpdateListItems',
        async: false,
        batchCmd: 'New',
        listName: listName,
        valuepairs: data,
        completefunc: function(xData, Status) {
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
    return new Promise(function(resolve, reject) {
        if (!listName) {
            return Promise.reject();
        }
        $().SPServices({
            operation: 'UpdateListItems',
            async: true,
            batchCmd: 'New',
            listName: listName,
            valuepairs: data,
            completefunc: function(xData, Status) {
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
                    err['status'] = "error";
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
    if (!listName) {
        return false;
    }
    if (!itemID) {
        return false;
    }
    $().SPServices({
        operation: 'UpdateListItems',
        async: false,
        batchCmd: 'Delete', //New, Update, Delete, Moderate
        listName: listName,
        ID: itemID, //
        completefunc: function(xData, Status) {
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
    return new Promise(function(resolve, reject) {
        if (!listName) {
            reject(false);
        }
        if (!itemID) {
            reject(false);
        }
        $().SPServices({
            operation: 'UpdateListItems',
            async: true,
            batchCmd: 'Delete', //可以包含的参数: New, Update, Delete, Moderate
            listName: listName,
            ID: itemID,
            completefunc: function(xData, Status) {
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
        return false;
    }
    if (!itemID) {
        return false;
    }
    for (var i = 0; i < data.length; i++) {
        if (Object.prototype.toString.call(data[i]).indexOf('Array') === -1) {
            return false;
        }
    }
    var obj = {};
    $().SPServices({
        operation: "UpdateListItems",
        async: false,
        listName: listName,
        ID: itemID,
        valuepairs: data,
        completefunc: function(xData, Status) {
            if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
                obj['status'] = "success";
                obj['response'] = 'ID:' + itemID + " updated success";
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
    return new Promise(function(resolve, reject) {
        if (!listName) {
            return Promise.reject(false);
        }
        if (!itemID) {
            return Promise.reject(false);
        }
        for (var i = 0; i < data.length; i++) {
            if (Object.prototype.toString.call(data[i]).indexOf('Array') === -1) {
                return Promise.reject(false);
            }
        }
        $().SPServices({
            operation: "UpdateListItems",
            async: true,
            listName: listName,
            ID: itemID,
            valuepairs: data,
            completefunc: function(xData, Status) {
                if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
                    var obj = {};
                    obj['status'] = "success";
                    obj['response'] = 'ID:' + itemID + " updated success";
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
Date.prototype.format = function(fmt) { //author: meizz   
    var o = {
        "M+": this.getMonth() + 1, //月份   
        "d+": this.getDate(), //日   
        "h+": this.getHours(), //小时   
        "m+": this.getMinutes(), //分   
        "s+": this.getSeconds(), //秒   
        "q+": Math.floor((this.getMonth() + 3) / 3), //季度   
        "S": this.getMilliseconds() //毫秒   
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
    var vars = [],
        hash;
    var hashes;
    if (curParams.split('?').length > 2) {
        hashes = curParams.substring(curParams.lastIndexOf('?') + 1).split('&');
    } else {
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
 * 获取用户所包含的Group，同步函数，返回一个Promise对象
 * 刚函数输出一个数组array， 数组中包含所查询用户在当前站点下的Group权限组
 * 不填写用户参数，则默认为当前用户
 */

function getUserGroupsSync(username) {
    username ? username : username = $().SPServices.SPGetCurrentUser();
    var userInGroup = [];
    $().SPServices({
        operation: "GetGroupCollectionFromUser",
        userLoginName: username,
        async: false,
        completefunc: function(xData, Status) {
            if ($(xData.responseXML).SPFilterNode("Group").length > 0) {
                $(xData.responseXML).SPFilterNode("Group").each(function() {
                    userInGroup.push($(this).attr("Name") || "");
                });
            }
        }
    });
    return userInGroup;
}

/**
 * 获取用户所包含的Group，异步函数，返回一个Promise对象
 * 刚函数输出一个数组array， 数组中包含所查询用户在当前站点下的Group权限组
 * 不填写用户参数，则默认为当前用户
 */

function getUserGroupsAsync(username) {
    username ? username : username = $().SPServices.SPGetCurrentUser();
    return new Promise(function(resolve, reject) {
        var userInGroup = [];
        $().SPServices({
            operation: "GetGroupCollectionFromUser",
            userLoginName: username,
            async: true,
            completefunc: function(xData, Status) {
                if ($(xData.responseXML).SPFilterNode("Group").length > 0) {
                    $(xData.responseXML).SPFilterNode("Group").each(function() {
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

/**
 * 获取数组中最大的数据
 * 如果传入参数不是数组，或者传入的数组中存在非数字的元素，返回NaN
 */
function getMaxNumFromArray(arr) {
    if (Object.prototype.toString.call(arr).indexOf('Array') === -1) {
        return NaN;
    }
    return Math.max.apply(Math, arr);
}

Date.prototype.Format = function(formatStr) {
    var str = formatStr;
    var Week = ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'];

    str = str.replace(/yyyy|YYYY/, this.getFullYear());
    str = str.replace(/yy|YY/, (this.getYear() % 100) > 9 ? (this.getYear() % 100).toString() : '0' + (this.getYear() % 100));

    str = str.replace(/MM/, (this.getMonth() + 1) > 9 ? (this.getMonth() + 1).toString() : '0' + this.getMonth() + 1);
    str = str.replace(/M/g, this.getMonth() + 1);

    str = str.replace(/w|W/g, Week[this.getDay()]);

    str = str.replace(/dd|DD/, this.getDate() > 9 ? this.getDate().toString() : '0' + this.getDate());
    str = str.replace(/d|D/g, this.getDate());

    str = str.replace(/hh|HH/, this.getHours() > 9 ? this.getHours().toString() : '0' + this.getHours());
    str = str.replace(/h|H/g, this.getHours());
    str = str.replace(/mm/, this.getMinutes() > 9 ? this.getMinutes().toString() : '0' + this.getMinutes());
    str = str.replace(/m/g, this.getMinutes());

    str = str.replace(/ss|SS/, this.getSeconds() > 9 ? this.getSeconds().toString() : '0' + this.getSeconds());
    str = str.replace(/s|S/g, this.getSeconds());

    return str;
}

//生成title
function GenerateNumber() {
    var newdate = new Date();
    var newmonth = (newdate.getMonth() + 1);
    var curNumber = newdate.Format('YYYY') + (newmonth < 10 ? ('0' + newmonth) : newmonth) + newdate.Format('DD') + "-" + newdate.Format('HHmmSS');
    var title = 'F' + curNumber;
    return title;
}

/**
 * 删除多条数据，异步方法
 * @param {必填} listname 需要删除数据的list名
 * @param {必填} itemIDArrayList 需要删除的id数组集合 [1,2,3,4...]
 * 使用方法:第一个参数传递list名称， 第二个参数传递需要删除的id数组
 * 删除成功之后，返回一个boolean值 true代表删除成功
 * 控制台中会打印出没有删除成功数据的id号
 */

/*

function deleteItemsInListAsync(listname, itemIDArrayList) {
   return new Promise(function (resolve, reject) {
       var deletedItemsIDList = [];
       try {
           if (itemIDArrayList.length == 0) {
               var obj = {};
               obj.status = 'success';
               obj.response = '没有数据被删除';
               resolve(obj);
               return;
           }
           (function loop(index) {
               $().SPServices({
                   operation: 'UpdateListItems',
                   async: true,
                   batchCmd: 'Delete', //New, Update, Delete, Moderate
                   listName: listname,
                   ID: itemIDArrayList[index],
                   completefunc: function (xData, Status) {
                       if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
                           var obj = {};
                           var itemID = $(xData.responseXML).SPFilterNode("z:row").attr("ows_ID");
                           obj.status = 'success';
                           obj.option = 'delete';
                           obj.ID = itemID;
                           deletedItemsIDList.push(obj);
                           if (index < itemIDArrayList.length - 1) {
                               console.log(index);
                               index = index + 1;
                               loop(index)
                           } else {
                               resolve(deletedItemsIDList);
                           }
                       } else {
                           //如果有没有被删除的数据，则不影响其他数据的删除
                           console.log('id为' + itemIDArrayList[index] + '的数据删除失败');
                           if (index < itemIDArrayList.length - 1) {
                               console.log(index);
                               index = index + 1;
                               loop(index)
                           } else {
                               resolve(deletedItemsIDList);
                           }
                       }
                   }
               });
           })(0)
       } catch (error) {
           console.log(error);
           reject(error)
       }
   })
}
*/

function deleteItemsInListAsync(listname, array) {
    return new Promise(function(resolve, reject) {
        var _array = [];
        for (var i = 0; i < array.length; i++) {
            (function(index) {
                _array[index] = new Promise(function(resolve, reject) {
                    $().SPServices({
                        operation: 'UpdateListItems',
                        async: true,
                        batchCmd: 'Delete', //可以包含的参数: New, Update, Delete, Moderate
                        listName: listname,
                        ID: array[index],
                        completefunc: function(xData, Status) {
                            if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
                                var obj = {};
                                var itemID = array[index];
                                obj['status'] = "success";
                                obj['response'] = 'ID:' + itemID + " deleted success";
                                obj['ID'] = itemID;
                                resolve(obj);
                            } else {
                                var err = {};
                                err['status'] = "error";
                                err['response'] = xData.responseXML;
                                err['ID'] = array[index];
                                reject(err);
                            }
                        }
                    });

                })
            })(i)
        }
        Promise.all(_array.map(function(p) {
                return p.catch(function(e) {
                    return e;
                })
            }))
            .then(function(result) {
                console.log(result);
                console.log('delete data success');
                resolve(true);
            })
            .catch(function(err) {
                console.log(err);
                console.log('delete data error');
                resolve(false);
            })
    })

}


/**
 * 检查一个元素是否包含于数组中
 * jquery 提供的方法
 * value:所查询的元素
 * array:目标数组
 * 包含于数组返回true, 不包含于数组返回false
 */

function isInclude(value, array) {
    var index = $.inArray(value, array);
    if (index >= 0) {
        return true;
    } else {
        return false;
    }
}