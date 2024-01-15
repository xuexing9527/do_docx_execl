// const DocReader = require('./../../reader/DocReader')

// const docReader = new DocReader({ path: './data/docEngineeringCostClassTwo', outputPath: './data/docEngineeringCostClassTwo/output'})

// docReader.read()

const DocWriter = require('../../writer/DocWriter')

const docxWriter = new DocWriter({ path: './data/docEngineeringCostClassTwo/062200230201常文丽实验报告合并版.docx' })

console.log(docxWriter)

// docxWriter.insertContent('2024年1月15日00:25:53')
// docxWriter.save()