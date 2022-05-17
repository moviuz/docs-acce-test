const StreamZip = require('node-stream-zip');

module.exports = {

    open: function(filePath) {
        return new Promise(
            function(resolve, reject) {
                const zip = new StreamZip({
                    file: filePath,
                    storeEntries: true
                })

                zip.on('ready', () => {
                    var chunks = []
                    var content = ''
                    zip.stream('word/document.xml', (err, stream) => {
                        if (err) {
                            reject(err)
                        }
                        stream.on('data', function(chunk) {
                            chunks.push(chunk)
                        })
                        stream.on('end', function() {
                            content = Buffer.concat(chunks)
                            zip.close()
                            resolve(content.toString())
                        })
                    })
                })
            }
        )
    },

    extract: function(filePath) {
        return new Promise(
            function(resolve, reject) {
                module.exports.open(filePath).then(function (res, err) {
                    if (err) { 
                        reject(err) 
                    }
                    console.log("****************************************")
                    var body = ''
                    var components = res.toString()
                   // var components = res.toString().split('descr=')
                   if (components == '<wp:docPr '){
                       console.log('+++++++++++++++HAY UN MACTH +++++++++++++++++')
                   }
                   ver imgs = filePath.get
                   console.log(components[2]+'TESIS')
                    for(var i=0;i<components.length;i++) {
                        //var tags = components[i].split('>')
                        //var content = tags[1].replace(/<.*$/,"/n")
                       // body += content+' '
                    }

                    console.log(components)

                    resolve(components)
                })
            }
        )
    }

}

return module.exports

