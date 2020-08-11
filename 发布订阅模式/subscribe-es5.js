/**
 * 发布订阅模式
 * 基于ES6开发的 发布订阅库
 * 版本v1.0
 * 作者：Tianyu Zhang
 * 日期：2020-08-11
 */
'use strict'
var _subscript = (function () {
	function _instanceof(left, right) {
		if (
			right != null &&
			typeof Symbol !== 'undefined' &&
			right[Symbol.hasInstance]
		) {
			return !!right[Symbol.hasInstance](left)
		} else {
			return left instanceof right
		}
	}

	function _classCallCheck(instance, Constructor) {
		if (!_instanceof(instance, Constructor)) {
			throw new TypeError('Cannot call a class as a function')
		}
	}

	function _defineProperties(target, props) {
		for (var i = 0; i < props.length; i++) {
			var descriptor = props[i]
			descriptor.enumerable = descriptor.enumerable || false
			descriptor.configurable = true
			if ('value' in descriptor) descriptor.writable = true
			Object.defineProperty(target, descriptor.key, descriptor)
		}
	}

	function _createClass(Constructor, protoProps, staticProps) {
		if (protoProps) _defineProperties(Constructor.prototype, protoProps)
		if (staticProps) _defineProperties(Constructor, staticProps)
		return Constructor
	}

	//Sub:发布订阅类
	var Sub = /*#__PURE__*/ (function () {
		function Sub() {
			_classCallCheck(this, Sub)

			//创建一个事件池，用来存储后期需要执行的方法
			this.$pond = []
		} //向事件池中追加方法

		_createClass(Sub, [
			{
				key: 'add',
				value: function add(func) {
					if (typeof func !== 'function') {
						return
					} else {
						//数组的some方法，验证数组中是否有新的
						var flag = this.$pond.some(function (item) {
							return item === func
						})
						!flag ? this.$pond.push(func) : null
					}
				}, //=> 从事件池中移除方法
			},
			{
				key: 'remove',
				value: function remove(func) {
					var $pond = this.$pond

					for (var i = 0; i < $pond.length; i++) {
						var item = $pond[i]

						if (item === func) {
							//$pond.splice(i, 1)  //导致数组塌陷问题，我们移除不能真的移除，将他改成null
							$pond[i] = null
							break
						}
					}
				}, //=>通知事件池中的方法，按照顺序执行
			},
			{
				key: 'fire',
				value: function fire() {
					var $pond = this.$pond

					for (
						var _len = arguments.length,
							args = new Array(_len),
							_key = 0;
						_key < _len;
						_key++
					) {
						args[_key] = arguments[_key]
					}

					for (var i = 0; i < $pond.length; i++) {
						var item = $pond[i]

						if (typeof item !== 'function') {
							//此时需要删除
							$pond.splice(i, 1) //因为上面删除了数据，所以导致数组的索引都变了（每个元素的索引都 -1 了），
							//需要将i减一
							i--
							continue
						}

						item.call.apply(item, [this].concat(args))
					}
				},
			},
		])

		return Sub
	})() //暴露出去给外面使用

	return function () {
		return new Sub()
	}
})()
