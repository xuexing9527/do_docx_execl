const BaseWriter = require('../BaseWriter')
const fs = require('fs')
const JSZip = require("jszip");

class DocWriter extends BaseWriter {
    constructor(config = { path: '', output: '' }) {
        super({ type: 'docx' })
        const { path, output } = config
        this.path = path
        this.output = output || `${this.path}-modified.docx`
        this.zip = null
        this.documentXml = ''
        this.modifiedContent = ''
    }

    async read() {
        try {
            // 读取现有 Word 文档
            const fileBuffer = fs.readFileSync(this.path);
            // 使用 JSZip 解压缩文档
            this.zip = await JSZip.loadAsync(fileBuffer);
            this.documentXml = await this.zip.file("word/document.xml").async("text");
        } catch (error) {
            console.error("Error manipulating docx:", error);
        }
    }

    async insertContent (txt) {
        try {
            // 此处每次都read，避免多次插入被覆盖
            await this.read()
            // 此处可定义多模版
            const contentToInsert = `<w:p><w:r><w:t>${txt}</w:t><br/></w:r></w:p>`
            // 在文档末尾插入要添加的内容，可以定义位置，暂定开头，结尾
            this.modifiedContent = this.documentXml.replace(/(<\/w:body><\/w:document>)/, `${contentToInsert}$1`);
            await this.save()
        } catch (error) {
            console.error("Error manipulating docx:", error);
        }
    }

    async save(outputPath = this.output) {
        // 使用 JSZip 更新文档内容
        this.zip.file("word/document.xml", this.modifiedContent);
        const updatedBuffer = await this.zip.generateAsync({ type: "nodebuffer" });
        // 将修改后的内容保存为新文档
        fs.writeFileSync(outputPath, updatedBuffer);
        console.log("Content inserted successfully.");
    }
}

module.exports = DocWriter