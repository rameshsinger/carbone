const express = require('express')
const fs = require('fs');
const carbone = require('./lib/index');
var docxXMLEditor = require('./lib/DocxXMLEditor');
const app = express()
const port = 3000
app.use(express.json({limit: '150mb'}));
app.get('/', (req, res) => {
  res.send('Report Generatr Running')
})
var options = {
    convertTo : 'docx' //can be docx, txt, ...
  };
app.post('/report', function (req, res) {  
      let templatePath=req.body.templatePath;
      let destinationPath=req.body.destinationPath;
      let reportData=req.body.reportData;
      carbone.render(templatePath, reportData, options, async function(err,result ){
        if (err) {
          console.log(err);
            res.send('Failure')
            return;
        }
        fs.writeFileSync(destinationPath, result);
        console.log("File Generated successfully:"+destinationPath)
        await docxXMLEditor.updateTotalPageNumber(destinationPath,1);
        res.send('SUCCESS')
        // process.exit(); // to kill automatically LibreOffice workers
      });
});
app.listen(port, () => {
  console.log(`app listening on port ${port}`)
})