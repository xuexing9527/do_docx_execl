const config = {
    outputPath: '../output',
    docxsPath: '../docxs',
    reg: {
        includeTargetReg: /<w:t>-=<\/w:t>.*?>(\d+?)<.*?<w:t>=-<\/w:t>/g,
        targetReg: /<w:t>(\d+?)<\/w:t>/,
    },

}
module.exports = config