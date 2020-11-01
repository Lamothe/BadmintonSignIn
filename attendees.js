'use strict'

const Excel = require('exceljs');
const fs = require('fs');
const $ = require('jquery');

function formatDate() {
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

function save() {
    let dateString = formatDate();
    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet(dateString);
    worksheet.columns = [
        { header: "Name", key: "name", width: "50" },
        { header: "Signed In", key: "checked" }
    ];
    worksheet.getRow(1).font = { bold: true };

    loadNames(function (names) {
        names.forEach(function (name, i) {
            worksheet.addRow({
                name: name,
                checked: 'x'
            });
        });

        workbook.xlsx.writeFile(dateString + '.xlsx');

        console.log("File saved");
    });
}

function init() {
    loadNames(function (names) {
        names.forEach(function (name, i) {
            addAttendee(name);
        });
        $(".check").on("change", save);
        save();
    });
}

function addAttendee(name) {
    $("#attendeeTable")
        .append("<tr><td><input type='checkbox' class='check'></td><td class='name'>" + name + "</td></tr>");
}

$(function () {
    init();
});
