var express = require('express');
var router = express.Router();

const fs = require('fs');
const XLSX = require("xlsx");
const XLSX_CALC = require('xlsx-calc');
var formulajs = require("formulajs");
var defaultFilename = "init.xlsx";
var filename = "data.xlsx";
XLSX_CALC.import_functions(formulajs)

function calculateCommissionProfit() {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"];

    commissionProfitList = [];

    console.log(-((worksheetTotal["C3"].v + worksheetDashboard["E3"].v) * (worksheetDashboard["B6"].v)))
    // 규범
    commissionProfitList.push((worksheetTotal["C3"].v + worksheetDashboard["E3"].v)*0.25*0.5)
    // 짱수
    commissionProfitList.push(-((worksheetTotal["C3"].v + worksheetDashboard["E3"].v) * (worksheetDashboard["B6"].v)))
    // 성수
    commissionProfitList.push((worksheetTotal["C3"].v + worksheetDashboard["E3"].v)*0.25*0.5)
    
    console.log(commissionProfitList)
    return commissionProfitList;
}

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
        totalProfitList.push(worksheetTotal["C" + i].v + worksheetDashboard["E" + i].v);
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
    var calculateCommissionList = calculateCommissionProfit();

    for (var i = 2; i < 5; i++) {
        var obj = {}
        obj["name"] = worksheetDashboard["A" + i].v;
        obj["money"] = worksheetTotal["B" + i].v + calculateCommissionList[i-2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v;
        obj["difference"] = worksheetTotal["D" + i].v;
        params.push(obj);
    }
    res.render('calculate', { params: params });
});

router.post("/insert", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetLog = workbook.Sheets["log"]
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"]

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
                v: (req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)]))
            }
            worksheetTotal['C' + (i+1)].v += worksheetDashboard['E' + (i + 1)].v
            worksheetDashboard['B' + (i + 1)].v = worksheetDashboard['D' + (i + 1)].v
            worksheetDashboard['B' + (i + 1)].v += (req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)]))
            worksheetTotal['B' + (i + 1)].v += (req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)]))
            worksheetDashboard['B7'].v += (req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)]))
        }
    }
    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/' + filename);
    res.redirect('/log')
});

router.post("/totalupdate", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]

    worksheetDashboard['B7'].v = (req.body['total'] == '' ? 0 : parseInt(req.body['total']))
    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/' + filename);
    res.redirect('/dashboard')
});

router.post("/commissionupdate", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]

    worksheetDashboard['B6'].v = (req.body['commission'] == '' ? 0 : parseFloat(req.body['commission']))
    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/' + filename);
    res.redirect('/dashboard')
});

router.post("/calculate", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"];

    var calculateCommissionList = calculateCommissionProfit();


    for (var i = 2; i < 5; i++) {
        worksheetTotal["D" + i].v += worksheetTotal["B" + i].v + calculateCommissionList[i-2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v - (req.body['player' + (i - 2)] == '' ? 0 : parseInt(req.body['player' + (i - 2)]));
        worksheetTotal["C" + i].v += worksheetDashboard["E" + i].v
        worksheetDashboard["B" + i].v = 0
    }
    worksheetDashboard['B7'].v = 0

    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/' + filename);
    res.redirect('/calculate')
});

router.get("/api/dashboard", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"];

    // var params = XLSX.utils.sheet_to_json(worksheetDashboard)
    var params = []
    var calculateCommissionList = calculateCommissionProfit();
    // for (var i = 2; i < 5; i++) {
    //     params[i - 2]['totalProfit'] = worksheetTotal["C" + i].v + worksheetDashboard["E" + i].v
    // }

    for (var i = 2; i < 5; i++) {
        var obj ={}
        obj["seed"] = parseFloat(worksheetTotal["B" + i].v.toFixed(2))
        obj["rate"] = parseFloat(worksheetDashboard["C" + i].v.toFixed(2))
        obj["balance"] = parseFloat((worksheetTotal["B" + i].v + calculateCommissionList[i-2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v).toFixed(2))
        obj["profit"] = parseFloat((obj["balance"] - obj["seed"]).toFixed(2))
        params.push(obj)
    }

    res.json(params);
});

router.get("/api/calculate", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"];
    var calculateCommissionList = calculateCommissionProfit();
    var params = []

    for (var i = 2; i < 5; i++) {
        var obj = {}
        obj["name"] = worksheetDashboard["A" + i].v;
        obj["money"] = worksheetTotal["B" + i].v + calculateCommissionList[i-2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v;
        obj["difference"] = worksheetTotal["D" + i].v;
        params.push(obj);
    }
    res.json(params);
});

router.get("/api/datalist", function(req, res, next) {
    var files = [];
    fs.readdirSync("./data/").forEach(file => {
        if (file.split(".")[1] == "xlsx") {
            files.push(file)
        }
    })
    res.json(files);
});

router.post("/api/insert", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetLog = workbook.Sheets["log"]
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"]

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
                v: (req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)]))
            }

            worksheetTotal['C' + (i+1)].v += worksheetDashboard['E' + (i + 1)].v
            worksheetDashboard['B' + (i + 1)].v = worksheetDashboard['D' + (i + 1)].v
            worksheetDashboard['B' + (i + 1)].v += (req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)]))
            worksheetTotal['B' + (i + 1)].v += (req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)]))
            worksheetDashboard['B7'].v += (req.body['player' + (i)] == '' ? 0 : parseInt(req.body['player' + (i)]))
        }
    }
    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/' + filename);
    res.json({
        result: "success"
    })
});

router.post("/api/totalupdate", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]

    worksheetDashboard['B7'].v = (req.body['total'] == '' ? 0 : parseInt(req.body['total']))
    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/' + filename);
    res.json({
        result: "success"
    })
});

router.post("/api/commissionupdate", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]

    worksheetDashboard['B6'].v = (req.body['commission'] == '' ? 0 : parseFloat(req.body['commission']))
    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/' + filename);
    res.json({
        result: "success"
    })
});

router.post("/api/calculate", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"];
    var calculateCommissionList = calculateCommissionProfit();

    for (var i = 2; i < 5; i++) {
        worksheetTotal["D" + i].v += worksheetTotal["B" + i].v + calculateCommissionList[i-2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v - (req.body['player' + (i - 2)] == '' ? 0 : parseInt(req.body['player' + (i - 2)]));
        worksheetTotal["C" + i].v += worksheetDashboard["E" + i].v
        worksheetDashboard["B" + i].v = 0
    }
    worksheetDashboard['B7'].v = 0

    XLSX_CALC(workbook)
        // write to new file
    XLSX.writeFile(workbook, './data/' + filename);
    res.json({
        result: "success"
    })
});

router.post("/api/filename", function(req, res, next) {
    filename = req.body['filename']
    res.json({ filename: filename });
});

router.post("/api/initdata", function(req, res, next) {
    filename = req.body['filename']
    let workbook = XLSX.readFile("./data/" + defaultFilename)
    XLSX_CALC(workbook)
    XLSX.writeFile(workbook, './data/' + filename);
    res.json({ filename: filename });
});


module.exports = router;