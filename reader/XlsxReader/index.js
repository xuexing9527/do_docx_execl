const BaseReader = require('../BaseReader')
const fs = require('fs')
const ExcelJS = require('exceljs');
const { timeStamp } = require('console');
 
class XlsxReader extends BaseReader {

    constructor (config = { path: '' }) {
        super({ type: 'xlsx', path: config.path })
        const { path } = config
        this.path = path
        this.workbook = new ExcelJS.Workbook()
        this.worksheetList = []
        this.worksheetDataList = []
    }


    // 支持多级目录，暂只支持平级目录，不存储不同级目录文件名
     getFileNames (callback) {
        // console.log(this.path)
        const fsReadDir = (path) => {
            fs.readdir(path, (err, files) => {
                // 此处容错处理
                // if (err) throw err
                if (err) console.log(err)
                files.forEach(item => {
                    const itemPath = `${path}/${item}`
                    // console.log(itemPath)
                    fs.stat(itemPath, (err, stats) => {
                        if (err) {
                            console.error(err)
                            return
                        }

                        // 检查是否是文件
                        if (stats.isFile()) {
                            // console.log(`${itemPath} 是一个文件。`)
                            this.fileList.push({ path: itemPath, name: item })
                            // this.read(itemPath)
                            callback && callback(itemPath)
                            // console.log(this.fileList)
                        }

                        // 检查是否是文件夹
                        if (stats.isDirectory()) {
                            // console.log(`${itemPath} 是一个文件夹。`)
                            fsReadDir(itemPath)
                        }
                    });

                })
            })

        }

        fsReadDir(this.path)

    }

    async read() {
        await this.workbook.xlsx.readFile(this.path)

        this.workbook.eachSheet((worksheet, sheetId) => {
            const arr = []
            worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                arr.push(row.values)
                // // 输出每行的数据
                // console.log(`Row ${rowNumber}: ${row.values}`)
            })
            this.worksheetDataList.push(arr)
        })

        // // 获取第一个工作表
        // const worksheet = this.workbook.getWorksheet(1);
        // // 遍历工作表的行
        // worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        //     // 输出每行的数据
        //     console.log(`Row ${rowNumber}: ${row.values}`)
        // })
    }

    getData () {
        return this.worksheetDataList
    }
}

module.exports = XlsxReader