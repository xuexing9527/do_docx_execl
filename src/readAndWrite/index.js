const DocReader = require('./../../reader/DocReader')

const docReader = new DocReader({ path: './data/docEngineeringCostClassTwo', outputPath: './data/docEngineeringCostClassTwo/output'})

docReader.read()