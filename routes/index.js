var express = require('express');
var router = express.Router();

const XLSX = require("xlsx");

/* GET home page. */
router.get('/', function(req, res, next) {
    res.render('index', { title: 'zzz' });
});

router.get("/excel", function(req, res, next) {
    let workbook = XLSX.readFile("./data/happy.xlsx")
    let worksheet = workbook.Sheets["log"]
        // console.log(XLSX.utils.sheet_to_json(worksheet))
    res.render('index', { object: XLSX.utils.sheet_to_json(worksheet) });
});

router.get("/insert", function(req, res, next) {
    let workbook = XLSX.readFile("./data/happy.xlsx")
    let worksheet = workbook.Sheets["log"]

    var columns = ['A', 'B', 'C', 'D'];
    var newIndex = parseInt(worksheet['!ref'].split(':')[1].slice(1)) + 1;
    worksheet['!ref'] = 'A1:D' + newIndex;

    for (var i = 0; i < columns.length; i++) {
        if (i == 0) {
            worksheet[columns[i] + newIndex] = {
                t: 'n',
                v: newIndex - 1
            }
        } else {
            worksheet[columns[i] + newIndex] = {
                t: 'n',
                v: 100
            }
        }
    }
    // write to new file
    XLSX.writeFile(workbook, './data/happy.xlsx');
    res.redirect('/excel')
});

module.exports = router;