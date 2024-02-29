const BaseReader = require('../BaseReader')
const fs = require('fs')
const JSZip = require("jszip");
 
class DocReader extends BaseReader {

    constructor (config = { path: '' }) {
        super({ type: 'docx', path: config.path })
        const { path } = config
        this.path = path
        // 必要时候，会传递zip
        this.zip = null
        this.documentXml = ''
    }

    async read () {
        try {
            // 读取现有 Word 文档
            const fileBuffer = fs.readFileSync(this.path);
            // 使用 JSZip 解压缩文档
            this.zip = await JSZip.loadAsync(fileBuffer);
            this.documentXml = await this.zip.file("word/document.xml").async("text");
        } catch (error) {
            console.error("Error manipulating docx path:", this.path)
            console.error("Error manipulating docx:", error);
        }
    }

    getXml () {
        return this.documentXml
    }
}

module.exports = DocReader