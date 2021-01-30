/**
 * 说明：
 * 工具方法类库
 * 增删改查SP list
 * sync代表同步方法，async代表异步方法
 * 使用此工具类库的前置条件为引用jquery(3.0)以上版本和引用SPService类库，还有Promise类库（如果浏览器不支持ES6需要引入）
 * version 1.7  github地址: https://github.com/Tianyu201809/sharepoint-library/tree/dev
 * 作者: apsolut China Co., Ltd.
 * 更新时间: 2020-1-29
 */


/**
 * 获取sharepoint list数据的同步函数
 * 注意填写arrayField参数时，list的显示字段和技术字段的名称要保持一致
 * @param {*string} listName 所查询列表的名称  （必填）
 * @param {*string} query   查询条件CAML语法 "<Query><Where></Where></Query>" （必填）
 * @param {*array => ["Title","ID"...]} arrayField    ["Title","ID"...]  （必填）
 * @param {string} queryNumber 想要查询的条目数   （可选）
 */
function getListDataSync(listName, query, arrayField, queryNumber) {
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
    CAMLRowLimit: isNaN(queryNumber) ? '' : String(parseInt(queryNumber)),
    completefunc: function (xData, Status) {
      if (Status === 'success') {
        if ($(xData.responseXML).SPFilterNode("z:row").length > 0) {
          $(xData.responseXML).SPFilterNode("z:row").each(function (i, val) {
            for (var j = 0; j < arrayField.length; j++) {
              var key = String(arrayField[j]);
              data[i] ? data[i] : data[i] = {};
              data[i][key] = $(this).attr("ows_" + key + "") || "";
            }
          });
        }
        //没有数据则 返回data = []
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
 * @param {*string} listName 所查询列表的名称  （必填）
 * @param {*string} query   查询条件CAML语法 "<Query><Where></Where></Query>" （必填）
 * @param {*array => ["Title","ID"...]} arrayField  （必填）
 * @param {string} queryNumber 想要查询的条目数   （可选）
 */
function getListDataAsync(listName, query, arrayField, queryNumber) {
  return new Promise(function (resolve, reject) {
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
      CAMLRowLimit: isNaN(queryNumber) ? '' : String(parseInt(queryNumber)),
      completefunc: function (xData, Status) {
        if (Status === 'success') {
          if ($(xData.responseXML).SPFilterNode("z:row").length > 0) {
            var data = [];
            $(xData.responseXML).SPFilterNode("z:row").each(function (i, val) {
              for (var j = 0; j < arrayField.length; j++) {
                var key = String(arrayField[j]);
                data[i] ? data[i] : data[i] = {};
                data[i][key] = $(this).attr("ows_" + key + "") || "";
              }
            });
            resolve(data);
          } else {
            // var err = {};
            // err['response'] = xData.responseText;
            // err['status'] = "error";
            resolve([]);
          }
        } else {
          var err = {};
          err['response'] = xData.responseText;
          err['status'] = "error";
          reject(err)
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
          err['status'] = "error";
          reject(err);
        }
      }
    });
  })
}


/**
 * 检查所传入的数组是否含有ID
 * 如果存在ID值，则返回ID
 * 不存在ID 返回一个空字符串
 */
function checkArrayhasIDField(array) {
  var flag = "";
  $.each(array, function (i, item) {
    var type = Object.prototype.toString.call(item);
    if (type === '[object Array]' && item[0] === 'ID' && (item[1] && item[1] != 'undefined')) {
      //说明该数组中包含ID字段
      flag = item[1];
      return false;
    }
  })
  return flag;
}

/**
 * 批量添加/更新表单数据
 * 如果传入存在[ID,'11']这种
 * @param {*} listName  表单名称
 * @param {*} arraylistItem  [[[],[],[]],[[],[],[]]] 传入的二维数组参数
 * 异步函数
 */
function mulInsertListDataAsync(listName, arraylistItems) {
  return new Promise(function (resolve, reject) {
    (function loop(index) {
      var f = checkArrayhasIDField(arraylistItems[index]);
      if (!f) {
        //create
        $().SPServices({
          operation: 'UpdateListItems',
          async: true,
          batchCmd: 'New',
          listName: listName,
          valuepairs: arraylistItems[index],
          completefunc: function (xData, Status) {
            if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
              var approvalLineID = $(xData.responseXML).SPFilterNode("z:row").attr("ows_ID");
              console.log("approval line id:" + approvalLineID);
              if (index < arraylistItems.length - 1) {
                console.log(index);
                index = index + 1;
                loop(index)
              } else {
                console.log('insert data success')
                resolve(true);
              }
            } else {
              var errorText = $(xData.responseXML).SPFilterNode("ErrorText")[0].textContent;
              reject(errorText);
            }
          }
        });
      } else {
        //update
        var approverLineID = _getItemsID(arraylistItems[index]);
        $().SPServices({
          operation: 'UpdateListItems',
          async: true,
          //batchCmd: 'New',
          ID: approverLineID,
          listName: listName,
          valuepairs: arraylistItems[index],
          completefunc: function (xData, Status) {
            if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
              if (index < arraylistItems.length - 1) {
                index = index + 1;
                loop(index)
              } else {
                console.log('update data success')
                resolve(true);
              }
            } else {
              var errorText = $(xData.responseXML).SPFilterNode("ErrorText")[0].textContent;
              reject(errorText);
            }
          }
        });
      }
    })(0);
  })
}

/**
 * 批量添加/更新表单数据
 * 如果传入存在[ID,'11']这种
 * @param {*} listName  表单名称
 * @param {*} arraylistItem  [[[],[],[]],[[],[],[]]] 传入的二维数组参数
 * 同步函数
 */
function mulInsertListDataSync(listName, arraylistItems) {
  var result = [];
  for (var index = 0; arraylistItems < array.length; index++) {
    var f = checkArrayhasIDField(arraylistItems[index]);
    if (!f) {
      //create
      $().SPServices({
        operation: 'UpdateListItems',
        async: false,
        batchCmd: 'New',
        listName: listName,
        valuepairs: arraylistItems[index],
        completefunc: function (xData, Status) {
          if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
            result.push(true)
          } else {
            var errorText = $(xData.responseXML).SPFilterNode("ErrorText")[0].textContent;
            console.log(errorText);
            result.push(false);
          }
        }
      });
    } else {
      //update
      var approverLineID = _getItemsID(arraylistItems[index]);
      $().SPServices({
        operation: 'UpdateListItems',
        async: false,
        //batchCmd: 'New',
        ID: approverLineID,
        listName: listName,
        valuepairs: arraylistItems[index],
        completefunc: function (xData, Status) {
          if (Status === "success" && $(xData.responseXML).find("ErrorCode").text() === "0x00000000") {
            result.push(true)
          } else {
            var errorText = $(xData.responseXML).SPFilterNode("ErrorText")[0].textContent;
            console.log(errorText);
            result.push(false)
          }
        }
      });
    }
  }
  var _result = true;
  if (result.length > 0) {
    result.forEach(function (item) {
      if (!item) {
        _result = false;
      }
    })
  }
  return _result;
}

/**
 * 获取二维数组中的id值
 * [["ID",12],["company","DGRC"]]
 */
function _getItemsID(array) {
  var approvalLineID = "";
  $.each(array, function (i, item) {
    var type = Object.prototype.toString.call(item);
    if (type === '[object Array]' && item[0] === 'ID' && item[1]) {
      //说明该数组中包含ID字段
      approvalLineID = item[1]
      return false;
    }
  })
  return approvalLineID;
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
    completefunc: function (xData, Status) {
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
  return new Promise(function (resolve, reject) {
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
      completefunc: function (xData, Status) {
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
Date.prototype.format = function (fmt) { //author: meizz   
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
 * 获取当前时间
 * ISO格式
 */

function getCurrentDateTime() {
  return ConvertDateISO(new Date());
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
    completefunc: function (xData, Status) {
      if ($(xData.responseXML).SPFilterNode("Group").length > 0) {
        $(xData.responseXML).SPFilterNode("Group").each(function () {
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
          resolve([]); //获取权限失败
        }

      }
    });

  })
}

/**
 * 检查当前登录人是是否存在某个权限或某些权限
 * @param {*String / array} userGroup 所查询的权限
 * 当参数为array时，根据传入的参数，返回对应的boolean值
 */
function userHasGroupSync(userGroup) {
  //首先检查userGroup类型
  if (userGroup && typeof userGroup === 'string') {
    var userContainGroup = getUserGroupsSync();
    if (userContainGroup.indexOf(userGroup) > -1) {
      return true
    }
  } else if (typeof userGroup === 'object' && Object.prototype.toString.call(userGroup) === '[object Array]' && userGroup) {
    var arr = [];
    for (var i = 0; i < userGroup.length; i++) {
      if (userContainGroup.indexOf(userGroup[i]) > -1) {
        arr.push(true);
      } else {
        arr.push(false);
      }
    }
    return arr;
  }
}


/**
 * 检查当前登录人是是否存在某个权限或某些权限
 * @param {*String / array} userGroup 所查询的权限
 * 当参数为array时，根据传入的参数，返回对应的boolean值
 */
function userHasGroupAsync(userGroup) {
  return new Promise(function (resolve, reject) {
    getUserGroupsAsync().then(function (userContainGroup) {
      //首先检查userGroup类型
      if (userGroup && typeof userGroup === 'string') {
        if (userContainGroup.indexOf(userGroup) > -1) {
          resolve(true)
        }
      } else if (typeof userGroup === 'object' && Object.prototype.toString.call(userGroup) === '[object Array]' && userGroup) {
        var arr = [];
        for (var i = 0; i < userGroup.length; i++) {
          if (userContainGroup.indexOf(userGroup[i]) > -1) {
            arr.push(true);
          } else {
            arr.push(false);
          }
        }
        resolve(arr);
      }
    }).catch(function (e) {
      reject(e)
    })
  })

}






/**
 * 对象克隆方法（深度克隆）
 * 注意：该方法不能克隆对象中存在的函数
 */
function cloneObj(obj) {
  return JSON.parse(JSON.stringify(obj))
}

/**
 * 获取数组中最大的数据
 * 如果传入参数不是数组，或者传入的数组中存在非数字的元素，返回NaN
 * @param {Array} arr 数组集合
 */
function getMaxNumFromArray(arr) {
  if (Object.prototype.toString.call(arr).indexOf('Array') === -1) {
    return NaN;
  }
  return Math.max.apply(Math, arr);
}




/**
 * 
 * @param {*String} formatStr 日期转换格式模板如:  YYYY-MM-DD HH:mm:ss 
 * 该函数返回所期待的日期字符串格式
 */
Date.prototype.Format = function (formatStr) {
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

/**
 * 
 * 生成唯一编号，一般用于Title字段
 * 该方法返回一个唯一的字符串
 */
function GenerateNumber(str) {
  str ? str : str = "F";
  var newdate = new Date();
  var newmonth = (newdate.getMonth() + 1);
  var curNumber = newdate.Format('YYYY') + (newmonth < 10 ? ('0' + newmonth) : newmonth) + newdate.Format('DD') + "-" + newdate.Format('HHmmSS');
  var title = str + curNumber;
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
function deleteItemsInListAsync(listname, array) {
  return new Promise(function (resolve, reject) {
    var _array = [];
    for (var i = 0; i < array.length; i++) {
      (function (index) {
        _array[index] = new Promise(function (resolve, reject) {
          $().SPServices({
            operation: 'UpdateListItems',
            async: true,
            batchCmd: 'Delete', //可以包含的参数: New, Update, Delete, Moderate
            listName: listname,
            ID: array[index],
            completefunc: function (xData, Status) {
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
    Promise.all(_array.map(function (p) {
      return p.catch(function (e) {
        return e;
      })
    }))
      .then(function (result) {
        console.log(result);
        console.log('delete data success');
        resolve(true);
      })
      .catch(function (err) {
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

/**
 * 数组去重方法
 */

function unique(arr) {
  return arr.filter(function (item, index, arr) {
    //当前元素，在原始数组中的第一个索引==当前索引值，否则返回当前元素
    return arr.indexOf(item, 0) === index;
  });
}


/**
* 输出两个数组中不同的元素
* @param { Array }arr1 
* @param { Array }arr2 
*/
function getArrDifference(arr1, arr2) {
  return arr1.concat(arr2).filter(function (v, i, arr) {
    return arr.indexOf(v) === arr.lastIndexOf(v)
  })
}

/**
 * 输出两个数组中相同的元素
 * @param { *Array } arr1 
 * @param { *Array } arr2 
 */

function getArrEqual(arr1, arr2) {
  var newArr = [];
  for (var i = 0; i < arr2.length; i++) {
    for (var j = 0; j < arr1.length; j++) {
      if (arr1[j] === arr2[i]) {
        newArr.push(arr1[j]);
      }
    }
  }
  return newArr;
}


/**
 * 对象合并polyfill
 * 为兼容IE浏览器使用Object.assign()方法
 */

function zyEs6AssignPolyfill() {
  if (!Object.assign) {
    Object.defineProperty(Object, "assign", {
      enumerable: false,
      configurable: true,
      writable: true,
      value: function (target, firstSource) {
        "use strict";
        if (target === undefined || target === null) throw new TypeError("Cannot convert first argument to object");
        var to = Object(target);
        for (var i = 1; i < arguments.length; i++) {
          var nextSource = arguments[i];
          if (nextSource === undefined || nextSource === null) continue;
          var keysArray = Object.keys(Object(nextSource));
          for (var nextIndex = 0, len = keysArray.length; nextIndex < len; nextIndex++) {
            var nextKey = keysArray[nextIndex];
            var desc = Object.getOwnPropertyDescriptor(nextSource, nextKey);
            if (desc !== undefined && desc.enumerable) to[nextKey] = nextSource[nextKey];
          }
        }
        return to;
      }
    });
  }
}

/**
 * 判断当前运行环境是否为IE浏览器
 * 是  返回true
 * 不是  返回fasle
 */
function isIE() {
  if (!!window.ActiveXObject || "ActiveXObject" in window) {
    return true;
  } else {
    return false;
  }
}

/**
 * 日期字符串截取
 * @param {*} str 
 */
function dateFormat(str) {
  if (!str || typeof str != 'string') {
    return "";
  }
  return str.split(" ")[0];
}

/**
 * ********************************
 * 将对象转化成二维数组
 * @param {*Object}
 * eg {a:1,b:2} => [[a,1],[b,2]]
 * ********************************
 */
function object2Array(obj) {
  var array = [];
  for (var key in obj) {
    var _arr = [];
    _arr[0] = key;
    _arr[1] = obj[key] || "";
    array.push(_arr);
  }
  return array;
}

/**
 * ********************************
 * 数组规范化处理
 * 去除数组中 Boolean值为false的数据
 * ********************************
 */
function removeIllegalArrayElement(array) {
  return array.filter(function (item) {
    return item;
  })
}

/**
 * 获取当前登录账户的详细信息
 * emial / tel / cost center等
 * @param {*} login 
 */
function getUserProfilebyLoginAsync(login) {
  //这里传递的是用户名称，如：apac\tianyuz
  return new Promise(function (resolve, reject) {
    if (!login) {
      login = $().SPServices.SPGetCurrentUser({
        fieldName: "Name"
      });
    }
    $().SPServices({
      operation: 'GetUserProfileByName',
      async: true,
      accountName: login,
      completefunc: function (xData, Status) {
        if (Status === 'success') {
          var user = {};
          $(xData.responseXML).SPFilterNode("PropertyData").each(function () {
            user[$(this).find("Name").text()] = $(this).find("Value").text();
          });
          ;
          user.login = user.AccountName || "";
          user.full_name = user.PreferredName || "";
          user.email = user.WorkEmail || "";
          user.department = user.Department || "";
          user.telephone = user.WorkPhone || "";
          user.dcxcostcenter = user.dcxCostCenter || "";
          user.userName = user.PreferredName || "";
          try {
            var dn = user["SPS-DistinguishedName"];
            var sstr = "";
            if (~dn.indexOf("GlobalResources")) {
              sstr = "GlobalResources,OU=";
            }
            else {
              sstr = "Users,OU=";
            }
            user.companycode = dn.split(sstr)[1].split(',')[0];
            resolve(user)
          }
          catch (e) {
            user.companycode = "";
            resolve(e)
          }
        } else {
          reject(xData)
        }
      }
    });
  })
}


/**
 * *********************************************
 * SharePoint batch webservice封装批量更新方法
 * 2021-1-29更新
 * *********************************************
 */

/**
 * ********************************
 * @param {*String} listName list列表名称
 * @param {*Object} itemsData 传递参数
 * @param {*Object Array} arrayField 需要返回查看的参数
 * 
 * 参数格式：
 * itemsData = [{    
 * //  ID:"1",
 * //  Title: "123",
 * //  RequestNumber: "F11111"
 * },{
 * //  ID:"2",
 * //  Title: "1234",
 * //  RequestNumber: "F111112"
 * }]
 * 
 * arrayField = [‘ID’, 'Title', 'OtherFieldName...']
 * 
 * 
 * ********************************
 */
function updateListItemsBatchAsync(listName, itemsData, arrayField) {
  var config = mappingParamsMethod(itemsData);
  var soapEnv = generateBatchString(listName, config);
  var url = _spPageContextInfo.webServerRelativeUrl + "/_vti_bin/lists.asmx";
  return new Promise(function (resolve, reject) {
    $.ajax({
      url: url,
      beforeSend: function (xhr) {
        xhr.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems");
      },
      contentType: "text/xml; charset=utf-8",
      type: "POST",
      dataType: "xml",
      data: soapEnv,
      complete: function (xData, Status) {
        debugger;
        if (Status === 'success') {
          if ($(xData.responseXML).SPFilterNode("Results").length > 0) {
            var arr = {};
            arr.status = 'success';
            arr.result = [];
            $(xData.responseXML).SPFilterNode("Result").each(function (i, val) {
              var data = {};
              var action = $(val).attr('ID').split(',')[1];
              var ErrorCode = $(this).find("ErrorCode").text();
              data.action = action;
              data.errorCode = ErrorCode;
              $(val).SPFilterNode("z:row").each(function (j, io) {
                if (arrayField && arrayField.length > 0) {
                  for (var k = 0; k < arrayField.length; k++) {
                    //将需要显示的字段放置到返回对象中
                    data[arrayField[k]] = $(io).attr("ows_" + arrayField[k]);
                  }
                } else {
                  data.ID = $(io).attr("ows_ID");
                  data.Title = $(io).attr("ows_Title");
                }

              });
              arr.result.push(data);
            })
            console.log(arr);
            resolve(arr);
          } else {
            resolve([]);
          }
        } else {
          reject('call webservice _vti_bin/lists.asmx error');
        }

      },
    });
  })

/**
* 
* @param {*Array} config  [{},{}]
* //转换之后参数格式如下：
* // config = [{
* //     method: "New",
* //     updateFields: {
* //         Title: "1",
* //         Date_Column: "1",
* //         Date_Time_Column: "1"
* //     },
* // }];
* 
*/
  //内部参数转换
  function mappingParamsMethod(config) {
    return config.map(function (item) {
      var obj = {};
      obj['method'] = 'New';
      obj['updateFields'] = item;
      for (var key in item) {
        if (
          key === 'ID' &&
          item[key] &&
          typeof item[key] !== 'undefined'
        ) {
          obj['method'] = 'Update';
          break;
        }
      }
      return obj;
    })
  }
  //生成soap请求报文
  function generateBatchString(listName, config) {
    if (!listName) {
      console.log('generateBatchString函数未传递listName');
      return;
    }

    if (JSON.stringify(config) == "{}") {
      console.log("config为空");
      return;
    }

    var onErrorAction = "Continue";
    var rootNode_start = "<Batch OnError=\"" + onErrorAction + "\"  >";
    var rootNode_end = "</Batch>";
    var methodNode_body = "";
    config.forEach(function (item, i) {
      var method = item.method;
      var methodId = i + 1;
      var methodNode_start = "<Method ID=\"" + methodId + "\" Cmd=\"" + method + "\">";
      var methodNode_end = "</Method>";
      for (var key in item.updateFields) {
        var FiledResutl = '';
        if (key === 'ID') {
          if (item.updateFields[key] && typeof item.updateFields[key] !== 'undefined') {
            var FieldNode_start = "<Field Name=\"" + key + "\">" + item.updateFields[key];
            var FieldNode_end = "</Field>";
            var FieldNode_body = FieldNode_start + FieldNode_end;
          } else {
            var FieldNode_start = "<Field Name=\"" + key + "\">" + item.updateFields[key];
            var FieldNode_end = "</Field>";
            var FieldNode_body = FieldNode_start + FieldNode_end;
          }
        }
        var FieldNode_start = "<Field Name=\"" + key + "\">" + item.updateFields[key];
        var FieldNode_end = "</Field>";
        var FieldNode_body = FieldNode_start + FieldNode_end;
        FiledResutl += FieldNode_body;
        methodNode_start += FiledResutl;
      }
      methodNode_start += methodNode_end;
      methodNode_body += methodNode_start;
    });
    rootNode_start += methodNode_body;
    var batch = rootNode_start + rootNode_end;
    var soapEnv = "<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'> \
         <soap:Body> \
        <UpdateListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'> \
        <listName>" + listName + "</listName> \
        <updates>" + batch + "</updates> \
        </UpdateListItems> \
        </soap:Body> \
        </soap:Envelope>";
    return soapEnv;
  }
}



/**
 * ************************************
 * 批量更新list数据 同步代码
 * ************************************
 */
function updateListItemsBatchSync(listName, itemsData, arrayField) {
  var config = mappingParamsMethod(itemsData);
  var soapEnv = generateBatchString(listName, config);
  var url = _spPageContextInfo.webServerRelativeUrl + "/_vti_bin/lists.asmx";
  var arr = {};
  arr.result = [];
  $.ajax({
    url: url,
    async: false,
    beforeSend: function (xhr) {
      xhr.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems");
    },
    contentType: "text/xml; charset=utf-8",
    type: "POST",
    dataType: "xml",
    data: soapEnv,
    complete: function (xData, Status) {
      if (Status === 'success') {
        if ($(xData.responseXML).SPFilterNode("Results").length > 0) {
          arr.status = 'success';
          $(xData.responseXML).SPFilterNode("Result").each(function (i, val) {
            var data = {};
            var action = $(val).attr('ID').split(',')[1];
            var ErrorCode = $(this).find("ErrorCode").text();
            data.action = action;
            data.errorCode = ErrorCode;
            $(val).SPFilterNode("z:row").each(function (j, io) {
              if (arrayField && arrayField.length > 0) {
                for (var k = 0; k < arrayField.length; k++) {
                  //将需要显示的字段放置到返回对象中
                  data[arrayField[k]] = $(io).attr("ows_" + arrayField[k]);
                }
              } else {
                data.ID = $(io).attr("ows_ID");
                data.Title = $(io).attr("ows_Title");
              }

            });
            arr.result.push(data);
          })
        }
      } else {
        arr.status = 'error';
        arr.result = [];
        arr.errorMessage = 'call method updateListItemsBatchSync fails, please check paramters';
      }
    },
  });

  return arr;
  //内部参数转换
  function mappingParamsMethod(config) {
    return config.map(function (item) {
      var obj = {};
      obj['method'] = 'New';
      obj['updateFields'] = item;
      for (var key in item) {
        if (
          key === 'ID' &&
          item[key] &&
          typeof item[key] !== 'undefined'
        ) {
          obj['method'] = 'Update';
          break;
        }
      }
      return obj;
    })
  }
  //生成soap请求报文
  function generateBatchString(listName, config) {
    if (!listName) {
      console.log('generateBatchString函数未传递listName');
      return;
    }

    if (JSON.stringify(config) == "{}") {
      console.log("config为空");
      return;
    }

    var onErrorAction = "Continue";
    var rootNode_start = "<Batch OnError=\"" + onErrorAction + "\"  >";
    var rootNode_end = "</Batch>";
    var methodNode_body = "";
    config.forEach(function (item, i) {
      var method = item.method;
      var methodId = i + 1;
      var methodNode_start = "<Method ID=\"" + methodId + "\" Cmd=\"" + method + "\">";
      var methodNode_end = "</Method>";
      for (var key in item.updateFields) {
        var FiledResutl = '';
        if (key === 'ID') {
          if (item.updateFields[key] && typeof item.updateFields[key] !== 'undefined') {
            var FieldNode_start = "<Field Name=\"" + key + "\">" + item.updateFields[key];
            var FieldNode_end = "</Field>";
            var FieldNode_body = FieldNode_start + FieldNode_end;
          } else {
            var FieldNode_start = "<Field Name=\"" + key + "\">" + item.updateFields[key];
            var FieldNode_end = "</Field>";
            var FieldNode_body = FieldNode_start + FieldNode_end;
          }
        }
        var FieldNode_start = "<Field Name=\"" + key + "\">" + item.updateFields[key];
        var FieldNode_end = "</Field>";
        var FieldNode_body = FieldNode_start + FieldNode_end;
        FiledResutl += FieldNode_body;
        methodNode_start += FiledResutl;
      }
      methodNode_start += methodNode_end;
      methodNode_body += methodNode_start;
    });
    rootNode_start += methodNode_body;
    var batch = rootNode_start + rootNode_end;
    var soapEnv = "<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'> \
           <soap:Body> \
          <UpdateListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'> \
          <listName>" + listName + "</listName> \
          <updates>" + batch + "</updates> \
          </UpdateListItems> \
          </soap:Body> \
          </soap:Envelope>";
    return soapEnv;
  }

}


