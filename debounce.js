/**
 * 说明：
 * 所谓防抖，就是指触发事件后在 n 秒内函数只能执行一次，
 * 如果在 n 秒内又触发了事件，则会重新计算函数执行时间
 * 
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
    delay = delay || 500;
    return function() {
        if (timer) {
            clearTimeout(timer);
        }
        timer = setTimeout(function() {
            func.apply(that, args)
        }, delay);
    }
}

/**
 * 说明：
 * 所谓节流，就是指连续触发事件但是在 n 秒中只执行一次函数
 * 
 * @param {function} *func 需要节流处理的函数名称 
 * @param {number} *wait 延时时间
 */
function throttle(func, wait) {
    let previous = 0;
    return function() {
        let now = Date.now();
        let that = this;
        let args = arguments;
        if (now - previous > wait) {
            func.apply(that, args);
            previous = now;
        }
    }
}