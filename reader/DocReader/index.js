const BaseReader = require('../BaseReader')
const admZip = require('adm-zip')
const fs = require('fs')
 
class DocReader extends BaseReader {

    constructor (config = { path: '', outputPath: '' }) {
        super({ type: 'doc', path: config.path })
        this.outputPath = config.outputPath
        this.fileList = []
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

    read (path) {

        this.getFileNames((path) => {
            const zip = new admZip(path);
            // // 此处应判断文件类型为docx，做兼容处理，以防报错
            if ((path.indexOf('docx') > 0)) {
                // 解压docx文件到目录 outputPath
                zip.extractAllTo(this.outputPath, true);
                // 提取内容
                let contentXml = zip.readAsText("word/document.xml");
                // console.log(contentXml)
            }

        })
        // fs.readdir(this.path, (err, files) => {
        //     if (err) throw err
        //     files.forEach(item => {
        //         // console.log(item)
        //         console.log(`${this.path}/${item}`)
        //         fs.readdir(`${this.path}/${item}`, (err, files) => {
        //             // 此处容错处理
        //             // if (err) throw err
        //             if (err) console.log(err)
        //             console.log(`${this.path}/${item}/${files}`)

        //         })
        //     })
        //     // writeXlsx(className, dataArr, classDataWrite)
        // })


    }
}

module.exports = DocReader