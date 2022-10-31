const express = require('express')
const fs = require('fs');
const carbone = require('./lib/index');
const app = express()
const port = 3000
app.use(express.json());
app.get('/', (req, res) => {
  res.send('Report Generatr Running')
})
var options = {
    convertTo : 'pdf' //can be docx, txt, ...
  };
app.post('/report', function (req, res) {  
    console.log(JSON.stringify(req.body));
      let templatePath=req.body.templatePath;
      let destinationPath=req.body.destinationPath;
      let reportData=req.body.reportData;
      carbone.render(templatePath, reportData, {}, function(err, result){
        if (err) {
            console.log(err);
            res.send('Failure')
            return;
        }
        fs.writeFileSync(destinationPath, result);
        console.log("File Generated successfully:"+destinationPath)
        res.send('SUCCESS')
        process.exit(); // to kill automatically LibreOffice workers
      });
});
app.listen(port, () => {
  console.log(`app listening on port ${port}`)
})