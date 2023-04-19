var filereader = require('./filereader.js')

filereader.extract('\\Users\\Fernando\\Documents\\prueba.docx').then(function(res, err) {
    if (err) {
        console.log(err)
    }
    console.log(res)
})