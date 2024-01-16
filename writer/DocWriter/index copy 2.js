const BaseWriter = require('../BaseWriter')
const unzipper = require("unzipper");
const fs = require('fs')
 
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

        fs.createReadStream(existingDocPath)
            .pipe(unzipper.Parse())
            .on("entry", entry => {
                if (entry.path === "word/document.xml") {
                    let data = "";
                    entry.on("data", chunk => {
                        data += chunk.toString();
                    });

                    entry.on("end", () => {
                        // 在文档头部插入要添加的内容
                        data = contentToInsert + "\n" + data;
                        // data = contentToInsert;
                        console.log(data)

                        // 将修改后的内容重新写入文档
                        fs.writeFileSync(outputDocPath, data);

                        console.log("Content inserted successfully.");
                    });
                } else {
                    // 保留原始文档的其余内容
                    // const entryStream = entry.pipe(fs.createWriteStream(`temp/${entry.path}`));
                    // entryStream.on("finish", () => {
                    //     entry.autodrain();
                    // });
                    entry.autodrain();
                }
            });



    }

    insertContent(txt) {
    }

    save(savePath = this.path) {
    }
}

module.exports = DocWriter