const XlsxReader = require('../../reader/XlsxReader')
const DocReader = require('../../reader/DocReader')
const DocWriter = require('../../writer/DocWriter')

// 这里或加 __path__
const xlsxReader = new XlsxReader({ path: './src/readAndWrite/testData/test.xlsx' })
const docxReader = new DocReader({ path: './src/readAndWrite/testData/test.docx' })
const docxWriter = new DocWriter({ output: './src/readAndWrite/testData/test-modified.docx' })

async function main () {
    // 读xlsx
    await xlsxReader.read()
    // 处理成目标数据
    const getXlsxData = () => {
        // 取第2个表
        const targetSheet = xlsxReader.getData()[1]
        // 截取目标值
        const targetData = targetSheet.slice(6, 61)
        // 表头
        const keyNames = [
            '学号', '姓名', 'Windows基本操作', '手抄报', '个人简历', '常用函数', '论文开题报告模板', 'Internet基础知识', '平均分', '文字编辑', '表格',
            '文字排版', '创建、编辑、格式化表格', '数据排序、筛选等', '图表、分析数据', '幻灯片制作', '平均分', '课堂表现', '平时成绩', '机考成绩', '总成绩'
        ]

        const purgeData = targetData.map(item => {
            const row = {}
            keyNames.map((it, index) => {
                row[it] = item[index + 2]?.result || item[index + 2]
            })
            return row
        })
        console.log(JSON.stringify(purgeData))
        return purgeData
    }

    getXlsxData()

    // 读docx
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