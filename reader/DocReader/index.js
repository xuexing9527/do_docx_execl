const BaseReader = require('../BaseReader')
const docx = require('docx')
const fs = require('fs')
 
class DocReader extends BaseReader {

    constructor (config = { path: '' }) {
        super({ type: 'docx', path: config.path })
        // this.fileList = []
        // this.existingDoc = this.existingDoc || {}
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

        // this.getFileNames((path) => {
        //     const zip = new admZip(path);
        //     // // 此处应判断文件类型为docx，做兼容处理，以防报错
        //     if ((path.indexOf('docx') > 0)) {
        //         // 解压docx文件到目录 outputPath
        //         zip.extractAllTo(this.outputPath, true);
        //         // 提取内容
        //         let contentXml = zip.readAsText("word/document.xml");
        //         // console.log(contentXml)
        //     }

        // })


        // 读取现有的 Word 文档
        const { Document, Packer, Paragraph } = docx
        this.existingDoc = new Document(fs.readFileSync(path))
        return this.existingDoc

        // // 进行修改，添加新的段落
        // const newParagraph = new Paragraph("This is a new paragraph.")

        // existingDoc.addSection({
        //     properties: {},
        //     children: [newParagraph]
        // });

        // // 保存修改后的文档
        // Packer.toBuffer(existingDoc).then((buffer) => {
        //     // 将 buffer 写入文件或进行其他处理
        //     fs.writeFileSync(`modified.docx`, buffer);
        // });

    }
}

module.exports = DocReader