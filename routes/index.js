var express = require('express');
var router = express.Router();

const fs = require('fs');
const XLSX = require("xlsx");
const XLSX_CALC = require('xlsx-calc');
var formulajs = require("formulajs");
var defaultFilename = "init.xlsx";
var filename = "";
XLSX_CALC.import_functions(formulajs)

function calculateCommissionProfit() {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"];

    commissionProfitList = [];

    console.log(-((worksheetTotal["C3"].v + worksheetDashboard["E3"].v) * (worksheetDashboard["B6"].v)))
        // 규범
    commissionProfitList.push((worksheetTotal["C3"].v + worksheetDashboard["E3"].v) * (worksheetDashboard["B6"].v) * 0.5)
        // 짱수
    commissionProfitList.push(-((worksheetTotal["C3"].v + worksheetDashboard["E3"].v) * (worksheetDashboard["B6"].v)))
        // 성수
    commissionProfitList.push((worksheetTotal["C3"].v + worksheetDashboard["E3"].v) * (worksheetDashboard["B6"].v) * 0.5)

    console.log(commissionProfitList)
    return commissionProfitList;
}

/* GET home page. */
router.get('/', function(req, res, next) {
    res.render('result', { message: 'zzz' });
});

router.get("/dashboard", function(req, res, next) {
    if (filename == "") {
        res.json({ result: "파일을 먼저 등록 해 주세요." })
    } else {
        let workbook = XLSX.readFile("./data/" + filename)
        let worksheetDashboard = workbook.Sheets["Dashboard"]
        let worksheetTotal = workbook.Sheets["total"];

        var totalProfitList = []
        for (var i = 2; i < 5; i++) {
            totalProfitList.push(worksheetTotal["C" + i].v + worksheetDashboard["E" + i].v);
        }
        // console.log(XLSX.utils.sheet_to_json(worksheetDashboard))

        res.render('dashboard', { params: XLSX.utils.sheet_to_json(worksheetDashboard), totalProfitList: totalProfitList });
    }
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
        obj["money"] = worksheetTotal["B" + i].v + calculateCommissionList[i - 2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v;
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
            worksheetTotal['C' + (i + 1)].v += worksheetDashboard['E' + (i + 1)].v
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
    var files = [];
    fs.readdirSync("./data/").forEach(file => {
        if (file.split(".")[1] == "xlsx") {
            files.push(file)
        }
    })
    if (req.body['filename'] != undefined && req.body['filename'] != "" && files.find(file => file === req.body['filename'] + ".xlsx") == undefined) {
        let workbook = XLSX.readFile("./data/" + filename)
        let worksheetDashboard = workbook.Sheets["Dashboard"]
        let worksheetTotal = workbook.Sheets["total"];

        var calculateCommissionList = calculateCommissionProfit();


        for (var i = 2; i < 5; i++) {
            worksheetTotal["D" + i].v += worksheetTotal["B" + i].v + calculateCommissionList[i - 2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v - (req.body['player' + (i - 2)] == '' ? 0 : parseFloat(req.body['player' + (i - 2)]));
            worksheetTotal["C" + i].v = calculateCommissionList[i - 2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v
            worksheetDashboard["B" + i].v = 0
        }
        worksheetDashboard['B7'].v = 0

        XLSX_CALC(workbook)
            // write to new file
        XLSX.writeFile(workbook, './data/' + filename);

        filename = req.body['filename'] + ".xlsx"
        workbook = XLSX.readFile("./data/" + defaultFilename)
        XLSX_CALC(workbook)
        XLSX.writeFile(workbook, './data/' + filename);
        // res.json({ filename: filename });

        res.redirect('/calculate')
    } else {
        // alert("파일명이 입력되지 않았거나 중복, 혹은 올바르지 않은 값 입니다.")
        // res.redirect('/calculate')
        res.json({ result: "파일명이 입력되지 않았거나 중복, 혹은 올바르지 않은 값 입니다." });
    }
});

router.get("/api/checkinit", function(req, res, next) {
    if (filename == "") {
        res.json({
            result: false,
            message: "진행 이력이 없습니다. 새로 파일을 작성 해 주세요."
        })
    } else {
        res.json({
            result: true,
            message: "success"
        })
    }
})

router.get("/api/dashboard", function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetDashboard = workbook.Sheets["Dashboard"]
    let worksheetTotal = workbook.Sheets["total"];

    var files = [];
    fs.readdirSync("./data/").forEach(file => {
        if (file.split(".")[1] == "xlsx") {
            if (file != filename) {
                files.push(file)
            }
        }
    })

    var params = []
    var calculateCommissionList = calculateCommissionProfit();
    for (var i = 2; i < 5; i++) {
        var obj = {}
        obj["seed"] = parseFloat(worksheetTotal["B" + i].v.toFixed(2))
        obj["rate"] = parseFloat(worksheetDashboard["C" + i].v.toFixed(2))
        obj["balance"] = parseFloat((worksheetTotal["B" + i].v + calculateCommissionList[i - 2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v).toFixed(2))
        obj["profit"] = parseFloat((obj["balance"] - obj["seed"]).toFixed(2))
        obj["totalProfit"] = obj["balance"] - obj["seed"]
        console.log(files)
        files.forEach(file => {
            let workbook = XLSX.readFile("./data/" + file)
            let worksheetTotal = workbook.Sheets["total"];
            obj["totalProfit"] += worksheetTotal["C" + i].v
        })
        obj["totalProfit"] = parseFloat(obj["totalProfit"].toFixed(2))
        params.push(obj)
    }
    params.push({ "seed": worksheetDashboard["B6"].v })
    params.push({ "seed": worksheetDashboard["B7"].v })

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
        obj["money"] = worksheetTotal["B" + i].v + calculateCommissionList[i - 2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v;
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

            worksheetTotal['C' + (i + 1)].v += worksheetDashboard['E' + (i + 1)].v
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
        result: true,
        message: "success"
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
        result: true,
        message: "success"
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
        result: true,
        message: "success"
    })
});

router.post("/api/calculate", function(req, res, next) {
    var files = [];
    fs.readdirSync("./data/").forEach(file => {
        if (file.split(".")[1] == "xlsx") {
            files.push(file)
        }
    })
    if (req.body['filename'] != undefined && req.body['filename'] != "" && files.find(file => file === req.body['filename'] + ".xlsx") == undefined) {
        let workbook = XLSX.readFile("./data/" + filename)
        let worksheetDashboard = workbook.Sheets["Dashboard"]
        let worksheetTotal = workbook.Sheets["total"];

        var newFilename = req.body['filename'] + ".xlsx"
        let newWorkbook = XLSX.readFile("./data/" + defaultFilename)
        let newWorksheetTotal = newWorkbook.Sheets["total"];

        var calculateCommissionList = calculateCommissionProfit();


        for (var i = 2; i < 5; i++) {
            worksheetTotal["D" + i].v += worksheetTotal["B" + i].v + calculateCommissionList[i - 2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v - (req.body['player' + (i - 2)] == '' ? 0 : parseFloat(req.body['player' + (i - 2)]));
            newWorksheetTotal["D" + i].v = worksheetTotal["D" + i].v
            worksheetTotal["C" + i].v = calculateCommissionList[i - 2] + worksheetDashboard["E" + i].v + worksheetTotal["C" + i].v
            worksheetDashboard["B" + i].v = 0
        }
        worksheetDashboard['B7'].v = 0

        XLSX_CALC(workbook)
        XLSX.writeFile(workbook, './data/' + filename);

        // workbook = XLSX.readFile("./data/" + defaultFilename)
        XLSX_CALC(newWorkbook)
        XLSX.writeFile(newWorkbook, './data/' + newFilename);
        filename = newFilename
        res.json({ result: true, message: "success" });
    } else {
        res.json({ result: false, message: "파일명이 입력되지 않았거나 중복, 혹은 올바르지 않은 값 입니다." });
    }


});

router.post("/api/filename", function(req, res, next) {
    filename = req.body['filename'] + ".xlsx"
    res.json({ result: true, message: "success" });
});

router.post("/api/initdata", function(req, res, next) {
    filename = req.body['filename'] + ".xlsx"
    let workbook = XLSX.readFile("./data/" + defaultFilename)
    XLSX_CALC(workbook)
    XLSX.writeFile(workbook, './data/' + filename);
    res.json({ result: true, message: "success" });
});

router.get("/api/reset", function(req, res, next) {
    var files = [];
    fs.readdirSync("./data/").forEach(file => {
        if (file.split(".")[1] == "xlsx" && file != "init.xlsx") {
            fs.unlinkSync("./data/" + file)
        }
    })
    filename = ""
    console.log("filename: " + filename)
    res.json({ result: true, message: "success" })
})

router.get("/api/report", function(req, res, next) {
    var reportData = [];
    var calculateCommissionList = calculateCommissionProfit();
    console.log(fs.readdirSync("./data/"));
    fs.readdirSync("./data/").forEach(file => {
        if (file != "init.xlsx" && file.split(".")[1] == "xlsx") {
            var filedate = file.split(" ")[0]
            if (reportData.length == 0 || reportData[reportData.length - 1]["date"] != filedate) {
                var obj = {
                    "date": filedate.substring(3, 7),
                    "fileList": [
                        file
                    ]
                }
                reportData.push(obj)
            } else {
                reportData[reportData.length - 1]["fileList"].push(file)
            }
        }
    })

    reportData.forEach(function(data, idx, array) {
        data["totalSeed"] = {
            "player0": 0,
            "player1": 0,
            "player2": 0
        }
        data["totalProfit"] = {
            "player0": 0,
            "player1": 0,
            "player2": 0
        }
        data["fileList"].forEach(function(file, idx2, array2) {
            console.log(file)
            let workbook = XLSX.readFile("./data/" + file)
            let worksheetTotal = workbook.Sheets["total"];
            for (var i = 0; i < 3; i++) {
                data["totalSeed"]["player" + i] += worksheetTotal["B" + (i + 2)].v
                if (filename == file) {
                    let worksheetDashboard = workbook.Sheets["Dashboard"];
                    data["totalProfit"]["player" + i] += parseFloat((calculateCommissionList[i] + worksheetDashboard["E" + (i + 2)].v + worksheetTotal["C" + (i + 2)].v).toFixed(2))
                } else {
                    data["totalProfit"]["player" + i] += worksheetTotal["C" + (i + 2)].v.toFixed(2)
                }
            }
        })
    })
    res.json(reportData)
})


module.exports = router;