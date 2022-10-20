/*Sort sheet based on CGPA (DONE)
Each subject will have 50 seats
assign electives to each student
create an excel sheet of those electives*/

const parser = require('simple-excel-to-json')
const doc = parser.parseXls2Json('./Assignment.xlsx')[0];

// INSERTION SORT CODE TO SORT ARRAY OF OBJECTS
for (i = 1; i < doc.length; i++) {
    let key = doc[i].CGPA;
    let temp = doc[i];
    j = i - 1;
    while (j >= 0 && doc[j].CGPA < key) {
        doc[j + 1] = doc[j];
        j--;
    }
    doc[j + 1] = temp;
}
const json2xls = require("json2xls")
const fs = require("fs")
let count_fwt = 50;
let count_uiux = 50;
let count_erp = 50;
let count_its = 50;
const alloted_seats = doc.map((student) => {
    if (student.OPTION_1 == "Fundamentals of Web Technologies" && count_fwt > 0) {
        student.ALLOTED = "FWT";
        count_fwt--;
    }
    else if (student.OPTION_1 == "Internet, Technology and Society" && count_its > 0) {
        student.ALLOTED = "ITS";
        count_its--;
    }
    else if (student.OPTION_1 == "Enterprise Resource Planning" && count_erp > 0) {
        student.ALLOTED = "ERP";
        count_erp--;
    }
    else if (student.OPTION_1 == "User Interface/User Experience (UI/UX) Design" && count_uiux > 0) {
        student.ALLOTED = "UIUX";
        count_uiux--;
    }
    else if (student.OPTION_2 == "Fundamentals of Web Technologies" && count_fwt > 0) {
        student.ALLOTED = "FWT";
        count_fwt--;
    }
    else if (student.OPTION_2 == "Internet, Technology and Society" && count_its > 0) {
        student.ALLOTED = "ITS";
        count_its--;
    }
    else if (student.OPTION_2 == "Enterprise Resource Planning" && count_erp > 0) {
        student.ALLOTED = "ERP";
        count_erp--;
    }
    else if (student.OPTION_2 == "User Interface/User Experience (UI/UX) Design" && count_uiux > 0) {
        student.ALLOTED = "UIUX";
        count_uiux--;
    }
    else if (student.OPTION_3 == "Fundamentals of Web Technologies" && count_fwt > 0) {
        student.ALLOTED = "FWT";
        count_fwt--;
    }
    else if (student.OPTION_3 == "Internet, Technology and Society" && count_its > 0) {
        student.ALLOTED = "ITS";
        count_its--;
    }
    else if (student.OPTION_3 == "Enterprise Resource Planning" && count_erp > 0) {
        student.ALLOTED = "ERP";
        count_erp--;
    }
    else if (student.OPTION_3 == "User Interface/User Experience (UI/UX) Design" && count_uiux > 0) {
        student.ALLOTED = "UIUX";
        count_uiux--;
    }
    else if (student.OPTION_4 == "Fundamentals of Web Technologies" && count_fwt > 0) {
        student.ALLOTED = "FWT";
        count_fwt--;
    }
    else if (student.OPTION_4 == "Internet, Technology and Society" && count_its > 0) {
        student.ALLOTED = "ITS";
        count_its--;
    }
    else if (student.OPTION_4 == "Enterprise Resource Planning" && count_erp > 0) {
        student.ALLOTED = "ERP";
        count_erp--;
    }
    else if (student.OPTION_4 == "User Interface/User Experience (UI/UX) Design" && count_uiux > 0) {
        student.ALLOTED = "UIUX";
        count_uiux--;
    }
    return student;
})
console.log(doc);
const filtered_fwt = alloted_seats.filter(student => student.ALLOTED == "FWT");
const filtered_erp = alloted_seats.filter(student => student.ALLOTED == "ERP");
const filtered_its = alloted_seats.filter(student => student.ALLOTED == "ITS");
const filtered_uiux = alloted_seats.filter(student => student.ALLOTED == "UIUX");

const excelDocument_fwt = json2xls(filtered_fwt);
const excelDocument_erp = json2xls(filtered_erp);
const excelDocument_its = json2xls(filtered_its);
const excelDocument_uiux = json2xls(filtered_uiux);

fs.writeFileSync("FWT.xlsx", excelDocument_fwt, "binary");
fs.writeFileSync("ITS.xlsx", excelDocument_its, "binary");
fs.writeFileSync("ERP.xlsx", excelDocument_erp, "binary");
fs.writeFileSync("UIUX.xlsx", excelDocument_uiux, "binary");

// console.log(filteredDocument)
