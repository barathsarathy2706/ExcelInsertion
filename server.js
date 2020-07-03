console.log('Log in')
const express = require('express');
const bodyParser= require('body-parser')
const app = express();
var xl = require('excel4node');

app.use(bodyParser.urlencoded({ extended: true }))

app.listen(3000, function() {
  console.log('listening on 3000')
})
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html')
  console.log('dir name :' + __dirname)
})

// Require library
var xl = require('excel4node');
 
// Create a new instance of a Workbook class
var wb = new xl.Workbook();
 
// Add Worksheets to the workbook
var ws = wb.addWorksheet('Sheet 1');
var ws2 = wb.addWorksheet('Sheet 2');
 
// Create a reusable style
var style = wb.createStyle({
  font: {
    color: '#FF0800',
    size: 12,
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -',
});
 



app.post('/quotes', (req, res) => {
	// Set value of cell A1 to 100 as a number type styled with paramaters of style
ws.cell(1, 1)
  .string(req.body.name)
  .style(style);
  
ws.cell(1, 2)
  .string(req.body.quote)
  .style(style);
 
wb.write(req.body.name+'.xlsx');
  console.log(req.body)
   res.redirect('/')
})
