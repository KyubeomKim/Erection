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
        // let cell = worksheet['B7'].v;
        // console.log(cell)

    // modify value in D4
    // console.log(worksheet);
    console.log(worksheet['B2'].v);
    // worksheet['B7'].t = 's';
    // worksheet['C1'].t = 's';
    worksheet['C1'].v = 'Idiot.';

    // write to new file
    XLSX.writeFile(workbook, './data/happy.xlsx');
    res.render('index', { title: "success!" });
});

module.exports = router;