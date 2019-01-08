var express = require('express');
var router = express.Router();

const XLSX = require("xlsx");
const XLSX_CALC = require('xlsx-calc');
var formulajs = require("formulajs");
var filename = "happy.xlsx";
XLSX_CALC.import_functions(formulajs)

/* GET home page. */
router.get('/', function(req, res, next) {
    res.render('result', { message: 'zzz' });
});

router.get("/dashboard", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"];

    var totalProfitList = []
    for (var i = 2; i < 5; i++) {
        totalProfitList.push(worksheetTotal["B" + i].v + worksheetDashboard["E" + i].v);
    }
    // console.log(XLSX.utils.sheet_to_json(worksheetDashboard))

    res.render('dashboard', { params: XLSX.utils.sheet_to_json(worksheetDashboard), totalProfitList: totalProfitList });
});

router.get("/log", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetLog = workbook.Sheets["log"]
        // console.log(XLSX.utils.sheet_to_json(worksheetLog))
    res.render('log', { object: XLSX.utils.sheet_to_json(worksheetLog) });
});

router.get("/calculate", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"];
    var params = [];

    for (var i = 2; i < 5; i++) {
        var obj = {}
        obj["name"] = worksheetDashboard["A" + i].v;
        obj["money"] = worksheetDashboard["D" + i].v;
        obj["difference"] = worksheetTotal["C" + i].v;
        params.push(obj);
    }
    res.render('calculate', { params: params });
});

router.post("/insert", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
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
                v: req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)])
            }
            worksheetDashboard['B' + (i + 1)].v += req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)])
            worksheetDashboard['B7'].v += req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)])
        }
    }
    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/happy.xlsx');
    res.redirect('/log')
});

router.post("/totalupdate", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]

    worksheetDashboard['B7'].v = req.body['total'] == '' ? 0 : parseInt(req.body['total'])
    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/happy.xlsx');
    res.redirect('/dashboard')
});

router.post("/calculate", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"];

    for (var i = 2; i < 5; i++) {
        worksheetTotal["B" + i].v += worksheetDashboard["E" + i].v
        worksheetTotal["C" + i].v += worksheetDashboard["D" + i].v - req.body['player' + (i - 2)] == '' ? 0 : parseInt(req.body['player' + (i - 2)]);
        worksheetDashboard["B" + i].v = 0
    }
    worksheetDashboard['B7'].v = 0

    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/happy.xlsx');
    res.redirect('/calculate')
});

module.exports = router;