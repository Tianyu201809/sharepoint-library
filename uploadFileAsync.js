/**
 * 重构方法：
 * 异步上传附件，防止浏览器卡死
 * 封装附件上传异步模式
 */


/**
 * 
 * @param {(enent obj)} kendoUI的时间对象 
 */
function onUploadFiles(e) {
    //默认传入对象是kendoUI的时间对象
    if (typeof e === 'object') {
        var files = e.files;
        _uploadFilesCommonAsync(files, listName, listItemID).then(function(message) {
            console.log(message);
        }).catch(function(e) {
            console.log(e)
        })
    } else if (typeof e === 'string') {
        //获取上传控件id，根据id获取目前上传的附件
        var files = $("#" + e)[0].files;
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
                    }
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
                    if (!SP || !SP.Base64EncodedByteArray) {
                        //如果没有加载sharepoint相应类库，先加载类库方法，然后执行上传逻辑
                        SP.SOD.executeFunc("sp.js", 'SP.ClientContext', function() {
                            var contentData = transformBlob(data.result);
                            uploadFileToSPServer(listName, listItemID, contentData, function(b) {
                                if (b) {
                                    //当前文件上传成功
                                    resolve2(true)
                                } else {
                                    //当前文件上传失败
                                    reject2(false)
                                }
                            })
                        });
                    } else {
                        //已经加载了内置类库，执行上传逻辑
                        var contentData = transformBlob(data.result);
                        uploadFileToSPServer(listName, listItemID, window._fileName, contentData, function(b) {
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
            resolve('所有文件上传成功')
        }).catch(function(e) {
            //有的附件没有上传完成
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
function uploadFileToSPServer(listName, listItemID, fileName, contentData, callback) {
    $().SPServices({
        operation: "AddAttachment",
        listName: listName,
        async: true,
        listItemID: listItemID,
        fileName: fileName,
        attachment: contentData.toBase64String(),
        completefunc: function(xData, Status) {
            if (Status != 'success') {
                callback(true)
            } else {
                callback(false)
            }
        }
    });
}