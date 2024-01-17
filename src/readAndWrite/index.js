const DocWriter = require('../../writer/DocWriter')
const DocReader = require('../../reader/DocReader')
const docxReader = new DocReader({ path: './data/docEngineeringCostClassTwo/test.docx' })
const docxWriter = new DocWriter({ output: './data/docEngineeringCostClassTwo/test-modified.docx' })

async function main () {
    // 读
    await docxReader.read()
    // 获取 xml
    const docxXml = docxReader.getXml()

    const handleXml = (docxXml) => {
        const txt = `${new Date()}：我要插入这一行。`
        const contentToInsert = `<w:p><w:r><w:t>${txt}</w:t><br/></w:r></w:p>`
        // 在docxXml 末尾插入要添加的内容 contentToInsert
        return docxXml.replace(/(<\/w:body><\/w:document>)/, `${contentToInsert}$1`)
    }
    const modifiedXml = handleXml(docxXml)

    // 写
    docxWriter.write(modifiedXml, docxReader.zip)
}

main()