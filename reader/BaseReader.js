/**
 * 读word，execl文件，拿到数据
 * 基类Reader
 * 基本方法1：读文件
 * 基本方法2：获取读取文件后的数据
 * 根据type实例化不同类型的工具，进行读取
 * 
 */

class BaseReader {
    constructor ({ type }) {
        // 文件类型
        this.type = type
    }

    // 读，异步操作
    read () {

    }

    // 获取数据
    getData () {
    }
}

module.exports = BaseReader