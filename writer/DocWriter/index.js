const BaseWriter = require('../BaseWriter')
const docx = require('docx')
const fs = require('fs')
 
class DocWriter extends BaseWriter {

    constructor (config = { path: '' }) {
        const { path } = config
        super({ type: 'docx', path })
        // 读取现有的 Word 文档
        const { Document } = docx
        console.log(path)
        console.log(fs.readFileSync(path))
        this.existingDoc = new Document(fs.readFileSync(path))
    }

    // 插入内容
    insertContent(txt) {
        const { Paragraph } = docx
        console.log(this.existingDoc)
        // 进行修改，添加新的段落
        const newParagraph = new Paragraph(txt);

        this.existingDoc.addSection({
            properties: {},
            children: [newParagraph]
        });

    }

    // 保存
    save () {
        const { Packer } = docx
        // 保存修改后的文档到新文件
        Packer.toBuffer(this.existingDoc).then((buffer) => {
            // 将 buffer 写入新文件
            fs.writeFileSync("modified.docx", buffer);
        });
    }


}

module.exports = DocWriter