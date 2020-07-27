/**
 * 说明：
 * 函数防抖与节流功能的实现
 * 防抖函数概述：当一个函数在某一场景下需要被频繁触发时（比如搜索框中，输入关键字，出现下拉提示信息），我们想要减少监听函数
 * 调用的频率，这时，我们就需要使用到防抖函数了
 * 
 * 原理：防抖函数最终返回一个新的函数，这个函数有判断功能，如果当前函数还处于计时器计时状态，那么
 * 就不执行函数
 * 
 */


/**
 * 
 * @param {function} *func 需要被防抖处理的执行函数的名称
 * @param {number} *delay  延时时间（默认为500毫秒）
 * 使用方式：debounce('你的函数名称', 延迟时间)  debounce('你的函数名称', 500)
 */

function debounce(func, delay) {
    var timer = null;
    var that = this;
    delay ? delay : 500;
    return function() {
        if (timer) {
            clearTimeout(timer);
        }
        timer = setTimeout(function() {
            func.apply(that, args)
        }, delay);
    }
}