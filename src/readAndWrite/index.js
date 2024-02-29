const XlsxReader = require('../../reader/XlsxReader')
const DocReader = require('../../reader/DocReader')
const DocWriter = require('../../writer/DocWriter')

// 这里或加 __path__
const xlsxReader = new XlsxReader({ path: './testData/1.xlsx' })
const pathArr = []

// consum = pathArr[1].slice(17).slice(0, 12)
// const docxReader = new DocReader({ path: './testData/test.docx' })
// const docxReader = new DocReader({ path: pathArr[1]})
// const docxWriter = new DocWriter({ output: `./testData/modified/${pathArr[1].slice(17).slice(0, pathArr[1].slice(17).indexOf('/'))}.docx` })

async function main () {
    // 读xlsx
    await xlsxReader.read()
    // 处理成目标数据
    const getXlsxData = () => {
        // 取第2个表
        const targetSheet = xlsxReader.getData()[1]
        // 截取目标值(表格的第几行到第几行)
        const targetData = targetSheet.slice(6, 61)
        // 表头
        // const keyNames = [
        //     '学号', '姓名', 'Windows基本操作', '手抄报', '个人简历', '常用函数', '论文开题报告模板', 'Internet基础知识', '平均分', '文字编辑', '表格',
        //     '文字排版', '创建、编辑、格式化表格', '数据排序、筛选等', '图表、分析数据', '幻灯片制作', '平均分', '课堂表现', '平时成绩', '机考成绩', '总成绩'
        // ]
        const keyNames = [
            '学号', '姓名', 'Windows基本操作', '手抄报', '个人简历', '常用函数', '论文开题报告模板', 'Internet基础知识']
        

        const purgeData = targetData.map(item => {
            const row = {}
            keyNames.map((it, index) => {
                // 遍历 keyNames，给 keyNames 找到对应的位置的 列值
                row[it] = item[index + 2]?.result || item[index + 2]
            })
            return row
        })
        // console.log(JSON.stringify(purgeData))
        return purgeData
    }

    // const xlsContent = JSON.stringify(getXlsxData()) 
    const xlsContent = getXlsxData()
    
    console.log(xlsContent)

    // // 读docx
    // await docxReader.read()
    // // 获取 xml
    // const docxXml = docxReader.getXml()

    const handleXml = (docxXml, num) => {
        // const txt = `${new Date()}：刘鑫玥`
        let modifiledXml = ''
        xlsContent.map(item => {
            if (item['学号'] === num) {
                // const txt = JSON.stringify(item).match(/{(.*)}/)[1]
                const txt = `学号: ${item['学号']}<br/>` + 
                    `姓名: ${item['姓名']}<br/>` +
                    `Windows基本操作: ${item['Windows基本操作']}<br/>` +
                    `手抄报: ${item['手抄报']}<br/>` +
                    `个人简历: ${item['个人简历']}<br/>` +
                    `常用函数: ${item['常用函数']}<br/>` +
                    `论文开题报告模板: ${item['论文开题报告模板']}<br/>` +
                    `Internet基础知识: ${item['Internet基础知识']}<br/>`
                
                const contentToInsert = `<w:p><w:r><w:t>${txt}</w:t><br/></w:r></w:p>`
                // 在docxXml 末尾插入要添加的内容 contentToInsert
                modifiledXml = docxXml.replace(/(<\/w:body><\/w:document>)/, `${contentToInsert}$1`)
                return docxXml.replace(/(<\/w:body><\/w:document>)/, `${contentToInsert}$1`)
            }
        })

        console.log(num)
        return modifiledXml
    }
    
    pathArr.map(async item => {
        console.log(item)
        const docxReader = new DocReader({ path: item })
        const docxWriter = new DocWriter({ output: `./testData/modified/${item.slice(17).slice(0, item.slice(17).indexOf('/'))}.docx` })
        const num = item.slice(17).slice(0, 12)
         // 读docx
        await docxReader.read()
        // 获取 xml
        const docxXml = docxReader.getXml()
       
        const modifiedXml = handleXml(docxXml, num)

        // 写
        docxWriter.write(modifiedXml, docxReader.zip)
    })

    // const modifiedXml = handleXml(docxXml)

    // // 写
    // docxWriter.write(modifiedXml, docxReader.zip)
}

main()