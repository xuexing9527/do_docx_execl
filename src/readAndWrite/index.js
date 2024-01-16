const DocWriter = require('../../writer/DocWriter')
const docxWriter = new DocWriter({ path: './data/docEngineeringCostClassTwo/test.docx', output: './data/docEngineeringCostClassTwo/test-modified.docx' })

async function main () {
    await docxWriter.insertContent(`指导教师1：丰年2678    This is the content I want to insert at the beginning. ${new Date().getTime()}`)
    await docxWriter.insertContent(`指导教师2：丰年2678    This is the content I want to insert at the beginning. ${new Date().getTime()}`)
}

main()