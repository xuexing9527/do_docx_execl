const BaseWriter = require('../BaseWriter')
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater')
const fs = require('fs')
const path = require("path");
 
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

    read () {
        const content = fs.readFileSync(this.path, "binary");
        // // console.log(content)
        // this.doc.loadZip(content)
        const zip = new PizZip(content);

        // This will parse the template, and will throw an error if the template is
        // invalid, for example, if the template is "{user" (no closing tag)
        this.doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });
    }

    insertContent(txt) {
        txt = txt || 'This is a new paragraph'
        this.doc.render({
            first_name: "John",
            last_name: "Doe",
            phone: "0652455478",
            description: "New Website",
            txt
        })

    }

    save(savePath = this.path) {
        // Get the zip document and generate it as a nodebuffer
        const buf = this.doc.getZip().generate({
            type: "nodebuffer",
            // compression: DEFLATE adds a compression step.
            // For a 50MB output document, expect 500ms additional CPU time
            compression: "DEFLATE",
        });

        // buf is a nodejs Buffer, you can either write it to a
        // file or res.send it with express for example.
        console.log(__dirname)
        fs.writeFileSync(path.resolve(__dirname, "text.docx"), buf);
    }
}

module.exports = DocWriter