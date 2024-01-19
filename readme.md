## 设计初衷
- 批量读写docx和xlsx文件（未经深入调研，只做初步工具。如有心，再考虑后续改造）。

## 项目结构
- reader
    - 读取部分，暂支持docx，xlsx读取，分别为：DocReader，XlsxReader。
    - FileReader，当前文档添加时未完成，本意打算读取文件夹方便，并做文件类型清洗。抽离DocReader或XlsxReader中的文件容错处理

- writer
    - DocWriter，项目初期需求只需批量写docx文件故先完成。
        - 目标支持文档前后简单插入
        - 以拼接字符串的形式，拼接xml文件。（因前期缺乏docx文件中xml格式标签嵌套研究，嵌套错误导致内容不可见，故采取拼接字符串的粗糙形式，未采用DOM操作）

- src  
src中为业务逻辑
    - readAndWrite
        - 读xlsx文件中的目标数据，批量往docx文件末尾插入处理后的目标数据

## 其它事项(待调整)
- 读取数据文件夹未列明
- 因撰写文档日，执行文件各目录未做调整，执行命令暂时是到readAndWrite下执行：  
    ````js
        cd src/readAndWrite
        node index.js
    ````