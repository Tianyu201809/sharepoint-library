/**
 * 重构方法：
 * 异步上传附件，防止浏览器卡死
 * 封装附件上传异步模式
 * 
 * v1.1 修改版本
 * 未测试版，可能无法直接使用
 */


/**
 * 
 * @param {(event obj)}  * kendoUI的组件的事件对象 
 * 注意：如果没有使用kendoui类库，那平时使用时传递file控件的id值 
 */
function onUploadFiles(e) {
    //默认传入对象是kendoUI的事件对象
    if (typeof e === 'object') {
        //此时默认为kendoui的事件对象
        var files = e.files;
        var listName, listItemID;
        /**
         * 此处请设置 listName, listItemID的参数值
         */


        _uploadFilesCommonAsync(files, listName, listItemID).then(function (message) {
        /**
         * 此处请设置 listName, listItemID的参数值
         */
            console.log(message);
        }).catch(function (e) {
        /**
         * 此处请设置 listName, listItemID的参数值
         */
            console.log(e)
        })
    } else if (typeof e === 'string') {
        //获取上传控件id，根据id获取目前上传的附件
        var files = $("#" + e)[0].files; //获取所上传的文件类数组对象
        var listName, listItemID;
        /**
         * 此处请设置 listName, listItemID的参数值
         */

        var length = files.length; //文件数组长度
        if (length === 0) return; //说明没有上传新的文件

        //执行上传方法
        _uploadFilesCommonAsync(files, listName, listItemID).then(function (message) {
            /**
             * 此处可以处理返回信息弹窗或是取消遮罩层的操作
             */
            console.log(message); //打印信息
        }).catch(function (e) {
            /**
             * 此处可以处理返回信息弹窗或是取消遮罩层的操作
             */
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
 * @param listName  sharepoint list的名称
 * @param listItemID list item id
 */
function _uploadFilesCommonAsync(files, listName, listItemID) {
    return new Promise(function (resolve, reject) {
        //如果必填参数没有填写，返回false
        if (!files || !listName || !listItemID) {
            return Promise.reject(false);
        }
        var f = isArray(files); //files是数组返回true，不是数组返回false 
        if (!f) {
            //不是数组
            files = Array.prototype.slice.call(files); //将类数组对象转化为数组
        }
        Promise.all(files.map(function (file) {
            return new Promise(function (resolve1, reject1) {
                var reader = new FileReader();
                /**
                 * 如果使用了KendoUI插件
                 * 则需要从file中的rawFile提取文件属性
                 */
                file.rawFile ? file = file.rawFile : null;
                var fileName = file.name;
                var obj = {};
                //当文件读取成功之后
                reader.onloadend = function (event) {
                    obj.result = event.target.result;
                    obj.fileName = fileName;
                    resolve1(obj);
                };
                //当文件读取失败之后
                reader.onerror = function (event) {
                    obj.result = event.target.result;
                    obj.fileName = fileName;
                    reject1(obj);
                };
                //读取文件的blob内容
                reader.readAsArrayBuffer(file);

            })
        }))
            .then(function (filesCollect) {
                //filesCollect.result是解析完成file文件集合
                //len是所上传的文件数量
                var len = filesCollect.length;
                //所上传的文件名集合容器
                var filesNameContainer = [];
                return new Promise(function (resolve) {
                    if (!window.SP.Base64EncodedByteArray) {
                        //如果没有加载sharepoint相应类库，先加载类库方法，然后执行上传逻辑
                        SP.SOD.executeFunc("sp.js", "SP.ClientContext", function () {
                            //声明一个递归函数, 使用递归的方式进行附件上传
                            (function loopCallBack(index) {
                                var contentData = transformBlob(filesCollect[index].result);
                                contentData = contentData.toBase64String();
                                $().SPServices({
                                    operation: "AddAttachment",
                                    listName: listName,
                                    async: true,
                                    listItemID: listItemID,
                                    fileName: fileName,
                                    attachment: contentData,
                                    completefunc: function (xData, Status) {
                                        if (Status == 'success') {
                                            filesNameContainer.push(filesCollect[index].fileName); //将上传成功的文件的文件名放入容器中
                                            if (index < len - 1) {
                                                //条件成立，还需要继续上传下一个文件
                                                index = index + 1;
                                                loopCallBack(index);
                                            } else {
                                                //后面没有新的文件要上传了, 将所有上传成功的文件的文件名输出
                                                resolve(filesNameContainer);
                                            }

                                        } else {
                                            //没有成功的不放入容器数组中
                                            if (index < len - 1) {
                                                //条件成立，还需要继续上传下一个文件
                                                index = index + 1;
                                                loopCallBack(index);
                                            } else {
                                                //后面没有新的文件要上传了, 将所有上传成功的文件的文件名输出
                                                resolve(filesNameContainer)
                                            }
                                        }
                                    }
                                });
                            })(0)
                        })
                    } else {
                        //声明一个递归函数, 使用递归的方式进行附件上传
                        (function loopCallBack(index) {
                            var contentData = transformBlob(filesCollect[index].result);//转化每一个文件为blob格式
                            contentData = contentData.toBase64String();
                            $().SPServices({
                                operation: "AddAttachment",
                                listName: listName,
                                async: true,
                                listItemID: listItemID,
                                fileName: fileName,
                                attachment: contentData,
                                completefunc: function (xData, Status) {
                                    if (Status == 'success') {
                                        //将上传成功的文件的文件名放入容器中
                                        var obj = {};
                                        obj.status = 'success';
                                        obj.fileName = filesCollect[index].fileName;
                                        filesNameContainer.push(obj);
                                        if (index < len - 1) {
                                            //条件成立，还需要继续上传下一个文件
                                            index = index + 1;
                                            loopCallBack(index);
                                        } else {
                                            //后面没有新的文件要上传了, 将所有上传成功的文件的文件名/状态 输出
                                            resolve(filesNameContainer);
                                        }
                                    } else {
                                        //没有成功的不放入容器数组中
                                        var obj = {};
                                        obj.status = 'error';
                                        obj.fileName = filesCollect[index].fileName;
                                        filesNameContainer.push(obj);
                                        if (index < len - 1) {
                                            //条件成立，还需要继续上传下一个文件
                                            index = index + 1;
                                            loopCallBack(index);
                                        } else {
                                            //后面没有新的文件要上传了, 将所有上传的文件的文件名/状态 输出
                                            resolve(filesNameContainer)
                                        }
                                    }
                                }
                            });
                        })(0)
                    }
                })
            })
            .then(function (result) {
                //此时所有附件上传完成
                console.log(result); //result为对象数组，每个元素是每个文件上传的状态
                resolve(result);
            })
            .catch(function (e) {
                //有的附件没有上传完成
                console.log(e);
                reject('文件上传失败');
            })

    })
}


/**
 * 
 * @param {buffer} reader对象读取完成的文件
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
 * 未使用
 * @param listName 列表名称
 * @param listItemID  表单ID编号
 * @param contentData  附件数据（buffer）
 */


// function uploadFileToSPServer(listName, listItemID, fileName, contentData) {
//     return new Promise(function (resolve, reject) {
//         $().SPServices({
//             operation: "AddAttachment",
//             listName: listName,
//             async: true,
//             listItemID: listItemID,
//             fileName: fileName,
//             attachment: contentData.toBase64String(),
//             completefunc: function (xData, Status) {
//                 if (Status != 'success') {
//                     resolve(true);
//                 } else {
//                     reject(false);
//                 }
//             }
//         });
//     })
// }

/**
 * 判断一个变量是否为数组
 * 是数组返回true
 * 非数组返回false
 * @param {*} obj 
 */
function isArray(obj) {
    return Object.prototype.toString.call(obj) === '[object Array]';
}