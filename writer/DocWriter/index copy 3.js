const BaseWriter = require('../BaseWriter')
const fs = require('fs')
const mammoth = require("mammoth");
const chardet = require("chardet");

 
class DocWriter extends BaseWriter {

    constructor (config = { path: '' }) {
        const { path } = config
        super({ type: 'docx' })
        this.path = path
        this.initinal()
    }

    initinal () {
        // todo 初始化读取doc文件，或者创建
        this.read()
    }

    read() {


        // 准备要插入的内容
        const contentToInsert = "This is the content I want to insert at the beginning.";

        // 读取现有的 Word 文档
        const existingDocPath = this.path;
        const outputDocPath = "modified-document.docx";

        // 使用 chardet 检测文档编码
        const fileBuffer = fs.readFileSync(existingDocPath);
        const detectedEncoding = chardet.detect(fileBuffer);
        console.log("Detected encoding:", detectedEncoding);

        // 将 fileBuffer 转换为 Uint8Array
        const fileUint8Array = new Uint8Array(Buffer.from(fileBuffer));

        mammoth.extractRawText({ buffer: fileBuffer })

            .then(result => {
                // 在文档头部插入要添加的内容
                const modifiedContent = contentToInsert + "\n" + result.value;

                // 使用 mammoth 尝试转换为 HTML（以保留格式）
                return mammoth.extract({ buffer: fileBuffer });
            })
            .then(result => {
                // 将新插入的内容合并到原始 HTML 中
                const modifiedHtml = result.value + `<p>${contentToInsert}</p>`;

                // 保存新文档
                fs.writeFileSync(outputDocPath, modifiedHtml, { encoding: "utf-8" });

                console.log("Content inserted successfully.");
            })
            .catch(error => {
                console.error("Error extracting raw text:", error);
            });



            // .then(result => {
            //     // 在文档头部插入要添加的内容
            //     const modifiedContent = contentToInsert + "\n" + result.value;

            //     // 使用 transformDocument 保留原文档的格式
            //     return mammoth.transformDocument({ buffer: fileBuffer }, { transform: mammoth.transforms.paragraph(modifiedContent) });
            // })
            // .then(transformedResult => {
            //     // 将修改后的内容保存为新文档
            //     fs.writeFileSync(outputDocPath, transformedResult.value, { encoding: "binary" });

            //     console.log("Content inserted successfully.");
            // })
            // .catch(error => {
            //     console.error("Error extracting raw text:", error);
            // });


            // .then(result => {
            //     // 在文档头部插入要添加的内容
            //     const modifiedContent = contentToInsert + "\n" + result.value;
            //     console.log(result)

            //     // 将修改后的内容保存为新文档
            //     fs.writeFileSync(outputDocPath, modifiedContent, { encoding: "utf-8" });

            //     console.log("Content inserted successfully.");
            // })
            // .catch(error => {
            //     console.error("Error extracting raw text:", error);
            // });

        // // 读取现有文档的内容
        // fs.readFile(existingDocPath, "windows-1252", (err, data) => {
        //     if (err) {
        //         throw err;
        //     }

        //     // 在文档头部插入要添加的内容
        //     const modifiedContent = contentToInsert + "\n" + data;

        //     // 将修改后的内容保存为新文档
        //     fs.writeFile(outputDocPath, modifiedContent, "windows-1252", err => {
        //         if (err) {
        //             throw err;
        //         }
        //         console.log("Content inserted successfully.");
        //     });
        // });

    }

    insertContent(txt) {
    }

    save(savePath = this.path) {
    }
}

module.exports = DocWriter