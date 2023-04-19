var app = require('connect')();
var http = require('http');
var serverPort = 8085;

var docx = require('./docx.js')

http.createServer(app).listen(serverPort, function () {
    console.log('Your server is listening on port %d (http://localhost:%d)', serverPort, serverPort);
  });

docx.extract('\\Users\\Fernando\\Documents\\prueba.docx').then(function(res, err) {
    if (err) {
        console.log(err)
    }
    console.log(res)
})