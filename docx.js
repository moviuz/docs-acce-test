const StreamZip = require('node-stream-zip');
const jsdom = require("jsdom");
const { JSDOM } = jsdom;
const { DOMParser, XMLSerializer  } = require('xmldom')
const util = require('util')
const xpath = require('xpath');
const _ = require('underscore')
const JsonFilter = require('@barreljs/json-filter')
var JSONPath = require('advanced-json-path');
const txml = require('txml');
const { first } = require('lodash');

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
                const validations = {
                    image : 'Invalid',
                    tables: 'Invalid',
                    url: 'Invalid'

                }
                module.exports.open(filePath).then(function (res, err) {
                    if (err) { 
                        reject(err) 
                    }
                    console.log("****************************************")
                    var body = ''
                    var components = res.toString()
                    var convert = require('xml-js')
                    var keys = ['w:p']
                   var result1 = convert.xml2json(components, {compact: true, spaces:4})
                   var result2 = convert.xml2js(components, {compact: false, spaces:4})
                   var exmap = new DOMParser().parseFromString(components)
                   //var DOMParser = new (require('xmldom')).DOMParser;
                   //var document = DOMParser.parseFromString(components);
                    let totalImg = exmap.getElementsByTagName('pic:cNvPr')
                  
                  // var nodesByName = document.getElementsByTagName('pic:cNvPr')
                  // console.log(g+'**************')
                   console.log(totalImg.length + " image array ")
                   //***MEOTDO PARA SABER SI IMAGENES TIENES DESCRIPCIÓN SI UNA NO CUENTA CON DESCRIPCION SE PARA EL METODO*****
                
                   for (let imgAux = 0 ; imgAux <= totalImg.length; imgAux0++ ){
                    if ( totalImg[imgAux].getAttribute("descr").length = 0 ) break
                   }


                  // console.log(typeof (g))
                   console.log(totalImg[0].getAttribute("descr"))
                   console.log(totalImg[0].getAttribute("descr").length + " img con texto alternativo") // aqui compruebo que las imagenes cuenten con texto alternativo

                   let hipervinculos = exmap.getElementsByTagName('w:hyperlink') //obtengo el array de hipervinculos
                   //console.log(hipervinculos[0].childNodes+'hipervinculos')
                    console.log(hipervinculos.length + " string array")
                   var body = ''
                   var components2 = hipervinculos[1].toString().split('<w:t') // obtengo el string de los hipervinculos
                   var components3 = hipervinculos[1].toString().split('<w:t')

                   for(var i=0;i<components2.length;i++) {
                       var tags = components2[i].split('>')
                       var content = tags[1].replace(/<.*$/,"")
                       body += content+' '
                   }
                   console.log('hipervinculo 1 : ' + body )
                   // si testRegex es FALSE , es accesible el enlace 
                   var regex = new RegExp("^(http[s]?:\\/\\/(www\\.)?|ftp:\\/\\/(www\\.)?|www\\.){1}([0-9A-Za-z-\\.@:%_\+~#=]+)+((\\.[a-zA-Z]{2,3})+)(/(.)*)?(\\?(.)*)?");
                   let testRegex = regex.test(body.trim())
                   console.log('resultado regex= ' + testRegex) //comparo si es una url y si lo es , no es accsible el enlace

                    console.log(body.trim()+ 'body')
                    console.log(typeof body)
                    console.log('unstring')
                  let x = exmap.documentElement.getElementsByTagName('pic:cNvPr')
                  var nodes = xpath.select('pic:cNvPr', components)
                  var path = '*..pic:cNvPr';
                  var result = JSONPath(result1, path)
                  var dom =  txml.getElementsByClassName(components, 'pic:cNvPr')
                 var tags = ['pic:cNvPr']
                ///**** Tabajo con Tablas 
               
                  let tables =  exmap.documentElement.getElementsByTagName('w:tbl') // Obtengo el numero de tablas del documento
                   console.log ( tables.length + 'total de tablas ')
                     //console.log(tables[0].nodeName = 'w:t')
                   //1let data = [...tables[0].rows].map(t => [...t.children].map(u => u.innerText))
                   // let y = tables[0].rows.length
                   //console.log('tamano de rows 0 ' + y )
                 // console.log(tables[0].getElementsByTagName('w:t') +  'prueba de tc')
                  let dataRows = tables[1].getElementsByTagName('w:t')
                  let totalcolums = tables[1].getElementsByTagName('w:tc')
                  let doubleRows = tables[1].getElementsByTagName('w:gridSpan')

                  console.log(dataRows.length  + ' Cantidad de texto en columnas')  // imprimo la cantidad de columndas con datos de texto
                  console.log(totalcolums.length  + 'Cantidad de Columnas') 
                  console.log(doubleRows.length  + ' Cantidad de row dobless ')
                  let atributtosTabla  = tables[1].getAttribute('w:firstRow')
                  console.log(atributtosTabla  + 'tipo de valor ')
                  console.log('************')
                  //console.log(tblFirstRow[1].getAttribute('w:firstRow'))
                  console.log( typeof atributtosTabla)
                  let tblFirstRow  = exmap.getElementsByTagName('w:tblLook')
                  
                  console.log(tblFirstRow[0].getAttribute('w:firstRow'))
                  console.log(tblFirstRow.length + '****')
                   if (tblFirstRow[1]  === '0'){
                       console.log('es tabla sin encabezados')
                       tblLook
                   }
                  // si dataRows es menor a totalColums quiere decir que falta datos en la tabla

                  // Obteniedo titulo  y descripcion de tabla
                  let tituloTables = tables[0].getElementsByTagName('w:tblCaption')
                  console.log(tituloTables.length + 'h1')
                  let tituloTables2 = tables[1].getElementsByTagName('w:tblCaption')
                  console.log(tituloTables2.length + 'h1')

                  let descripcionTablas  = tables[0].getElementsByTagName('w:tblDescription')

                  let descripcionTablas2  = tables[1].getElementsByTagName('w:tblDescription')
                  console.log( 'TablaComplta: ' +descripcionTablas.length  + ' Tablaincompleta: ' + descripcionTablas2.length )
                 // console.log(tables[0].getElementsByTagName('w:tblDescription'))
                 const filter = {
                     type: 'pic:cNvPr'
                 }
                 var resl = JsonFilter(result2,filter)
                let c = resl.all
                console.log(c)
                   try{
                }catch{
                    console.log(err)
                }
                  // console.log(result1)
                   var object = JSON.parse(result1)
                })
            }
        )
    },
   prueba: function(obj, keys) {
        if( keys.length ) {
          key = keys.shift();
          obj = obj[key];
          return prueba(obj, keys);
        } else {
          return obj;
        }
      }

}



