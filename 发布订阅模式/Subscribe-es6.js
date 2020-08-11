/**
 * 发布订阅模式
 * 基于ES6开发的 发布订阅库
 * 版本v1.0
 * 作者：Tianyu Zhang
 * 日期：2020-08-11
 */

let _subscript = (function () {
	//Sub:发布订阅类
	class Sub {
		constructor() {
			//创建一个事件池，用来存储后期需要执行的方法
			this.$pond = []
		}
		//向事件池中追加方法
		add(func) {
			if (typeof func !== 'function') {
				return
			} else {
				//数组的some方法，验证数组中是否有新的
				let flag = this.$pond.some((item) => {
					return item === func
				})
				!flag ? this.$pond.push(func) : null
			}
		}
		//=> 从事件池中移除方法
		remove(func) {
			let $pond = this.$pond
			for (let i = 0; i < $pond.length; i++) {
				let item = $pond[i]
				if (item === func) {
					//$pond.splice(i, 1)  //导致数组塌陷问题，我们移除不能真的移除，将他改成null
					$pond[i] = null
					break
				}
			}
		}
		//=>通知事件池中的方法，按照顺序执行
		fire(...args) {
			let $pond = this.$pond
			for (let i = 0; i < $pond.length; i++) {
				let item = $pond[i]
				if (typeof item !== 'function') {
                    //此时需要删除
                    $pond.splice(i, 1);
                    //因为上面删除了数据，所以导致数组的索引都变了（每个元素的索引都 -1 了），
                    //需要将i减一
                    i--;
                    continue;
				}
				item.call(this, ...args)
			}
		}
	}
	//暴露出去给外面使用
	return function () {
		return new Sub()
	}
})()
