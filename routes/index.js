var express = require('express');
var router = express.Router();

const XLSX = require("xlsx");
const XLSX_CALC = require('xlsx-calc');
var formulajs = require("formulajs");
XLSX_CALC.import_functions(formulajs)

/* GET home page. */
router.get('/', function(req, res, next) {
    res.render('index', { title: 'zzz' });
});

router.get("/dashboard", function(req, res, next) {
    let workbook = XLSX.readFile("./data/happy.xlsx")
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    console.log(XLSX.utils.sheet_to_json(worksheetDashboard))

    res.render('dashboard', { params: XLSX.utils.sheet_to_json(worksheetDashboard) });
});

router.get("/log", function(req, res, next) {
    let workbook = XLSX.readFile("./data/happy.xlsx")
    let worksheetLog = workbook.Sheets["log"]
        // console.log(XLSX.utils.sheet_to_json(worksheetLog))
    res.render('index', { object: XLSX.utils.sheet_to_json(worksheetLog) });
});

router.get("/insert", function(req, res, next) {
    let workbook = XLSX.readFile("./data/happy.xlsx")
    let worksheetLog = workbook.Sheets["log"]
    let worksheetDashboard = workbook.Sheets["Dashboard"]

    var columns = ['A', 'B', 'C', 'D'];
    var newIndex = parseInt(worksheetLog['!ref'].split(':')[1].slice(1)) + 1;
    worksheetLog['!ref'] = 'A1:D' + newIndex;

    for (var i = 0; i < columns.length; i++) {
        if (i == 0) {
            worksheetLog[columns[i] + newIndex] = {
                t: 'n',
                v: newIndex - 1
            }
        } else {
            worksheetLog[columns[i] + newIndex] = {
                t: 'n',
                v: 100
            }
            worksheetDashboard['B' + (i + 1)].v += 100
            worksheetDashboard['B7'].v += 100
        }
    }
    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/happy.xlsx');
    res.redirect('/log')
});

module.exports = router;