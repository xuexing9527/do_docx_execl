const BaseWriter = require('../BaseWriter')
const fs = require('fs')
const officegen = require("officegen")
const xml2js = require("xml2js");


class DocWriter extends BaseWriter {

    constructor(config = { path: '' }) {
        const { path } = config
        super({ type: 'docx' })
        this.path = path
        this.initinal()
    }

    initinal() {
        // todo 初始化读取doc文件，或者创建
        this.read()
    }

    read() {
        const fs = require("fs");
        const JSZip = require("jszip");

        const existingDocPath = this.path;
        const outputDocPath = "modified-document.docx";
        const contentToInsert = "丰年：This is the content I want to insert at the beginning.";


        // 读取现有 Word 文档
        const fileBuffer = fs.readFileSync(existingDocPath);

        // 使用 JSZip 解压缩文档
        async function modifyDocx() {
            try {
                const zip = await JSZip.loadAsync(fileBuffer);

                // 获取 document.xml 文件内容
                const documentXml = await zip.file("word/document.xml").async("text");

                // 找到 <w:body> 的位置
                const bodyStart = documentXml.indexOf("<w:body>");
                console.log(bodyStart)

                // 在文档开头插入要添加的内容，保留原始内容
                const modifiedContent = `${documentXml.slice(0, bodyStart)}${contentToInsert}${documentXml.slice(bodyStart)}`;
                console.log(modifiedContent.indexOf('丰年'))

                // 使用 JSZip 更新文档内容
                zip.file("word/document.xml", modifiedContent);

                // 将修改后的内容保存为新文档
                const updatedBuffer = await zip.generateAsync({ type: "nodebuffer" });

                fs.writeFileSync(outputDocPath, updatedBuffer);

                console.log("Content inserted successfully.");
            } catch (error) {
                console.error("Error manipulating docx:", error);
            }
        }

        // 调用异步函数
        modifyDocx();

        // // 使用 JSZip 解压缩文档
        // async function modifyDocx() {
        //     try {
        //         const zip = await JSZip.loadAsync(fileBuffer);

        //         // 获取 document.xml 文件内容
        //         const documentXml = await zip.file("word/document.xml").async("text");
        //         console.log(documentXml)

        //         // 在文档开头插入要添加的内容
        //         const modifiedContent = `<w:p>${contentToInsert}</w:p>${documentXml}`;

        //         // 使用 JSZip 更新文档内容
        //         zip.file("word/document.xml", modifiedContent);

        //         // 将修改后的内容保存为新文档
        //         const updatedBuffer = await zip.generateAsync({ type: "nodebuffer" });

        //         fs.writeFileSync(outputDocPath, updatedBuffer);

        //         console.log("Content inserted successfully.");
        //     } catch (error) {
        //         console.error("Error manipulating docx:", error);
        //     }
        // }

        // // 调用异步函数
        // modifyDocx();

    }

    insertContent(txt) {
    }

    save(savePath = this.path) {
    }
}

module.exports = DocWriter