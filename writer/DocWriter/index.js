const BaseWriter = require('../BaseWriter')
const fs = require('fs')
const JSZip = require("jszip");

class DocWriter extends BaseWriter {
    constructor(config = { output: '', documentXml: '' }) {
        super({ type: 'docx' })
        const { output } = config
        this.output = output || `${this.path}-modified.docx`
        this.documentXml = this.documentXml
    }

    /**
     * 
     * 写入，要么写入reader读取后的 zip 文件中，要么写入新的文档。
     * 为了保证格式的统一，建议写入 zip 原文件中，并打包 file 应该为 传入的 zip 中的file
     * 写的内容 docxXml  写到哪 zip  存到哪 outputPath
     * @param {*} docxXml
     * @param {*} zip
     * @param {*} outputPath 
     */
    async write (docxXml, zip, outputPath = this.output) {
        // 使用 JSZip 更新文档内容

        // 创建一个新文档，或者接收reader读取的文档 【此处如不用初始文档，是否会影响原格式？】
        const xml = docxXml !== undefined ? docxXml : this.documentXml
        zip.file("word/document.xml", xml);
        const updatedBuffer = await zip.generateAsync({ type: "nodebuffer" });
        // 将修改后的内容保存为新文档
        fs.writeFileSync(outputPath, updatedBuffer);
        console.log("Content inserted successfully.");
    }
}

module.exports = DocWriter