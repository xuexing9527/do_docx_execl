const admZip = require('adm-zip')
const fs = require('fs')
const config = require('./config')

// 读取内容
function readWord(name, errArr, className, dataArr) {
    // 此处应判断文件类型为docx，做兼容处理，以防报错
    const zip = new admZip(`${config.docxsPath}/${className}/${name}`);
    // 解压docx文件到目录 outputPath
    const outputPath = `${config.outputPath}/docxs/${className}/${name}`
    zip.extractAllTo(outputPath, true);
    // 提取内容
    let contentXml = zip.readAsText("word/document.xml");
    // 正则匹配文字
    let matchWT = contentXml.match(config.reg.includeTargetReg);
    if (!matchWT || matchWT?.length < 8) {
        errArr.push(`小于8个成绩的人：${name}, 当前为${matchWT?.map(item => item.match(config.reg.targetReg)[1])}`)
    } else {
        let arr = [];
        arr.push(name)
        arr = arr.concat(matchWT?.map(item => item.match(/<w:t>(\d+?)<\/w:t>/)[1]))
        dataArr.push(arr)
    }
}

function fsRead(className, classDataWrite) {
    fs.readdir(`./docxs/${className}`, (err, files) => {
        if (err) throw err
        const errArr = []
        const dataArr = []
        files.forEach(item => {
            readWord(item, errArr, className, dataArr)
        })
        writeXlsx(className, dataArr, classDataWrite)
    })
}

function writeXlsx(className, data, classDataWrite) {
    const xlsx = require('node-xlsx').default;
    // 目标处理数据, 依照传入的数据顺序
    const targetData = [];
    classDataWrite.forEach((item, index) => {
        targetData.push([item])
    })
    targetData.forEach((item, index) => {
        data.forEach(i2 => {
            if (i2[0].indexOf(item[0]) >= 0) {
                targetData[index] = targetData[index].concat(i2)
            }
        })
    })
    const buffer = xlsx.build([{ name: 'test', data: targetData }])
    fs.writeFileSync(`${config.outputPath}/xlsx/${className}.xlsx`, buffer)
}

// 格式化 xlsx 数据 书写顺序
const data1 = ['22300200101', '22300200102', '22300200103', '22300200104', '22300200105', '22300200106', '22300200107', '22300200108', '22300200109', '22300200110', '22300200111', '22300200112', '22300200113', '22300200114', '22300200115', '22300200116', '22300200117', '22300200118', '22300200119', '22300200120', '22300200121', '22300200122', '22300200123', '22300200124', '22300200125', '22300200126', '22300200127', '22300200128', '22300200129', '22300200130', '22300200131', '22300200132', '22300200133', '22300200134', '22300200135', '22300200136', '22300200137', '22300200138', '22300200139', '22300200140', '22300200141', '22300200142', '22300200143', '22300200144', '22300200145', '22300200146', '22300200147', '22300200148', '22300200149', '22300200150', '22300200151', '22300200152', '22300200153', '22300200154', '22300200155', '22300180315', '22300180552']
const data2 = [ '22300200201', '22300200202', '22300200203', '22300200205', '22300200206', '22300200207', '22300200208', '22300200209', '22300200210', '22300200211', '22300200212', '22300200213', '22300200214', '22300200215', '22300200216', '22300200217', '22300200218', '22300200219', '22300200220', '22300200221', '22300200222', '22300200223', '22300200224', '22300200225', '22300200226', '22300200227', '22300200228', '22300200229', '22300200230', '22300200231', '22300200232', '22300200233', '22300200234', '22300200235', '22300200236', '22300200237', '22300200238', '22300200239', '22300200240', '22300200241', '22300200242', '22300200243', '22300200244', '22300200245', '22300200246', '22300200247', '22300200248', '22300200249', '22300200250', '22300200251', '22300200252', '22300200253', '22300200254', '22300200255', '12100200133', '12100200236', '12200200121']

fsRead('1班', data1)
// fsRead('2班', data2)