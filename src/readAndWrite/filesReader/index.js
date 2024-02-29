const filesReader = require('./../../../reader/FileReader')

const reader = new filesReader({ path: './../testData' })
const arr = []
reader.getFileNames((data) => {
    console.log(data)
    arr.push(data)
    console.log(arr)
})

