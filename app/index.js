const fs = require('fs');
const carbone = require('./../lib/index');
var data = {
    firstname : 'John',
    lastname : 'Doe',
    image:'*Image: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',
    table: '*Table:{"title":"Receiver Settings","header":["col1","col2","col3"],"data":[["data1","data2","data3"]]}'
  };

  var options = {
    convertTo : 'pdf' //can be docx, txt, ...
  };

  carbone.render('./examples/simpleImage.odt', data, {}, function(err, result){
    if (err) return console.log(err);
    fs.writeFileSync('result.odt', result);
    // process.exit(); // to kill automatically LibreOffice workers
  });