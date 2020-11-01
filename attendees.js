'use strict'

const Excel = require('exceljs');
const fs = require('fs');
const $ = require('jquery');

const checkedString = "x";

function getFormattedDateString() {
    var d = new Date(),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) {
        month = '0' + month;
    }

    if (day.length < 2) {
        day = '0' + day;
    }

    return [year, month, day].join('-');
}

function getFileName() {
    var dateString = getFormattedDateString();
    return dateString + '.xlsx';
}

function loadNames(loaded) {

    fs.readFile('names.txt', 'utf8', function (err, data) {
        if (err) {
            return console.log(err);
        }

        var names = data
            .split('\r\n')
            .filter(function (el) { return /\S/.test(el); });

        loaded(names);
    });
}

function load(loaded) {
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(getFileName())
        .then(function () {
            var worksheet = workbook.getWorksheet(getFormattedDateString());
            worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
                if (rowNumber != 1) {
                    let name = row.values[1];
                    let checked = row.values[2] === checkedString;
                    addAttendee(name, checked);
                }
            });

            loaded();
        });
}

function save() {

    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet(getFormattedDateString());
    worksheet.columns = [
        { header: "Name", key: "name", width: "50" },
        { header: "Signed In", key: "checked" }
    ];
    worksheet.getRow(1).font = { bold: true };

    $("#attendeeTable tr").each(function (i, row) {
        let name = $(".name", row).text();
        let checked = $(".check", row).prop("checked");
        worksheet.addRow({
            name: name,
            checked: checked ? checkedString : ""
        });
    });

    workbook.xlsx.writeFile(getFileName());
    console.log("File saved");
}

function init() {
    if (fs.existsSync(getFileName())) {
        load(function () {
            $(".check").on("change", save);
        });
    }
    else {
        loadNames(function (names) {
            names.forEach(function (name, i) {
                addAttendee(name, false);
            });
            $(".check").on("change", save);
            save();
        });
    }
}

function addAttendee(name, checked) {
    $("#attendeeTable")
        .append("<tr><td><input type='checkbox'" + (checked ? " checked" : "") + " class='check'></td><td class='name'>" + name + "</td></tr>");
}

$(function () {
    init();
});
