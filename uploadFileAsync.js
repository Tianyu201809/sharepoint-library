/**
 * 重构方法：
 * 异步上传附件，防止浏览器卡死
 * 封装附件上传异步模式
 * 
 * v1.0 初始版本
 * 未测试版，可能无法直接使用
 */


/**
 * 
 * @param {(enent obj)}  * kendoUI的组件的事件对象 或者 传递file控件的id值 
 */
function onUploadFiles(e) {
    //默认传入对象是kendoUI的时间对象
    if (typeof e === 'object') {
        //此时默认为kendoui的事件对象
        var files = e.files;
        _uploadFilesCommonAsync(files, listName, listItemID).then(function(message) {
            console.log(message);
        }).catch(function(e) {
            console.log(e)
        })
    } else if (typeof e === 'string') {
        //获取上传控件id，根据id获取目前上传的附件
        var files = $("#" + e)[0].files[0]; //获取所上传的文件数组
        var length = files.length; //文件数组长度
        if (length === 0) return; //说明没有上传新的文件
        _uploadFilesCommonAsync(files, listName, listItemID).then(function(message) {
            console.log(message);
        }).catch(function(e) {
            console.log(e)
        })
    } else {
        //不能上传
        return false;
    }
}
/**
 * 
 * @param files 在页面上上传的多个文件
 * @param listName  目标form的名称
 * @param listItemID form item id
 */
function _uploadFilesCommonAsync(files, listName, listItemID) {
    return new Promise(function(resolve, reject) {
        //如果必填参数没有填写，返回false
        if (!files || !listName || !listItemID) {
            return Promise.reject(false);
        }
        Promise.all(files.map(function(file) {
            return new Promise(function(resolve1, reject1) {
                var reader = new FileReader();
                file = file.rawFile;
                var fileName = file.name;
                var obj = {};
                //当文件读取成功之后
                reader.onloadend = function(event) {
                    obj.result = event.target.result;
                    obj.fileName = fileName;
                    resolve1(obj);
                };
                //当文件读取失败之后
                reader.onerror = function(event) {
                    obj.result = event.target.result;
                    obj.fileName = fileName;
                    reject1(obj);
                };
                //读取文件的blob内容
                reader.readAsArrayBuffer(file);

            }).then(function(data) {
                //判断，并执行上传
                return new Promise(function(resolve2, reject2) {
                    //SP对象是SharePoint环境变量（全局）, 在SharePoint环境下存在
                    if (!window.SP.Base64EncodedByteArray) {
                        //如果没有加载sharepoint相应类库，先加载类库方法，然后执行上传逻辑
                        SP.SOD.executeFunc("sp.js", "SP.ClientContext", function() {
                            var contentData = transformBlob(data.result);
                            uploadFileToSPServer(listName, listItemID, data.fileName, contentData).then(function(b) {
                                if (b) {
                                    //当前文件上传成功
                                    var obj = {};
                                    var fileName = data.fileName;
                                    var status = 'success';
                                    obj.fileName = fileName;
                                    obj.status = status;
                                    resolve2(obj);
                                } else {
                                    //当前文件上传失败
                                    var obj = {};
                                    var fileName = data.fileName;
                                    var status = 'error';
                                    obj.fileName = fileName;
                                    obj.status = status;
                                    resolve(obj);
                                }
                            })
                        });
                    } else {
                        //已经加载了内置类库，执行上传逻辑
                        var contentData = transformBlob(data.result);
                        uploadFileToSPServer(listName, listItemID, data.fileName, contentData).then(function(b) {
                            if (b) {
                                //当前文件上传成功
                                resolve2(true)
                            } else {
                                //当前文件上传失败
                                reject2(false)
                            }
                        })
                    }
                })
            })
        })).then(function(result) {
            //此时所有附件上传完成
            console.log(result); //result里同时存在成功上传的附件和失败上传的附件信息
            resolve(result);
        }).catch(function(e) {
            //有的附件没有上传完成
            console.log(e);
            reject('文件上传失败');
        })

    })
}


/**
 * 
 * @param buffer reader对象读取完成的文件
 */
function transformBlob(buffer) {
    var bytes = new Uint8Array(buffer);
    var content = new SP.Base64EncodedByteArray();
    for (var i = 0; i < bytes.length; i++) {
        content.append(bytes[i]);
    }
    return content;
}

/**
 * 
 * @param listName 列表名称
 * @param listItemID  表单ID编号
 * @param contentData  附件数据（buffer）
 */
function uploadFileToSPServer(listName, listItemID, fileName, contentData) {
    return new Promise(function(resolve, reject) {
        $().SPServices({
            operation: "AddAttachment",
            listName: listName,
            async: true,
            listItemID: listItemID,
            fileName: fileName,
            attachment: contentData.toBase64String(),
            completefunc: function(xData, Status) {
                if (Status != 'success') {
                    resolve(true);
                } else {
                    reject(false);
                }
            }
        });
    })
}