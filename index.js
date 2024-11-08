const fs = require('fs').promises;
const XLSX = require('xlsx');
const path = require('path');

const backwallDataPath = path.join(__dirname, 'YSTRAW.xlsx');
const newBackwallPath = path.join(__dirname, 'NEW_DATA_SHEET.xlsx');

const backwallDataWorkbook = XLSX.readFile(backwallDataPath);
const newBackwallWorkbook = XLSX.readFile(newBackwallPath);


const sheetNamed1 = backwallDataWorkbook.SheetNames[0];
const sheetNamed2 = newBackwallWorkbook.SheetNames[0];


const sheet1 = backwallDataWorkbook.Sheets[sheetNamed1];
const sheet2 = newBackwallWorkbook.Sheets[sheetNamed2];

let backwallDataSheet = XLSX.utils.sheet_to_json(sheet1);
let newBackwallSheet = XLSX.utils.sheet_to_json(sheet2);

let newArr = [];
let matchedArr = [];
let notMatchedArr = [];
let notMDeviceId = [];

var wb = XLSX.utils.book_new();

const convertAndWriteXLSX_AHD = async () => {
    console.log("Actual Length :: ", backwallDataSheet.length, newBackwallSheet.length);

    backwallDataSheet?.forEach((elmData) => {
        const filteredData = newBackwallSheet.find((i) => i['Device ID'] == elmData['Ystraw ID']);
        // console.log(filteredData);

        if (filteredData) {
            // console.log(`Old Data TL N : ${elmData['TL Name']}, TL M : ${elmData['TL Mobile No']}`);
            // console.log(`New Data TL N : ${filteredData['TL Name']}, TL M : ${filteredData['TL Contact No']}`);
            // console.log(filteredData['Store Name'], filteredData['TL Name'], filteredData['TL Contact No'], filteredData[`AE Name`], filteredData[`AE Mobile NO`]);

            elmData['Outlet Address'] = filteredData['Address'];
            elmData['City'] = filteredData['City']?.toUpperCase();

            // console.log(elmData);
            newArr.push(elmData);
            // matchedArr.push(filteredData);
            // console.log(`${filteredData['Device Id']} == ${elmData['Device ID']}`);
            // console.log("Outlet Name Matched", filteredData, `${elmData['Reference Code']} :: ${filteredData[`Reference Code (Dhanush I'd)`]}`);
        } else {
            newArr.push(elmData);
            // console.log(elmData);
            // notMatchedArr.push(elmData);
        }
    });

    console.log(`updatedArr :: ${newArr.length} , baseFileArr :: ${backwallDataSheet.length}`);
    // console.log(`matchedArr :: ${matchedArr.length} , notMatchedArr :: ${notMatchedArr.length}`);

    // console.log(notMDeviceId, notMDeviceId.length);

    // Create a new worksheet and add the data
    var ws = XLSX.utils.json_to_sheet(newArr);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'EXCEL-BACKWALL.xlsx');
    console.log('Task Done...');

    // const updatedSheet = XLSX.utils.json_to_sheet(newArr);
    // const updatedWorkbook = XLSX.utils.book_new();
    // XLSX.utils.book_append_sheet(updatedWorkbook, updatedSheet, sheetNamed1);
    // await fs.writeFile(backwallDataPath, XLSX.write(updatedWorkbook, { bookType: 'xlsx', type: 'buffer' }));
    // console.log(`Updated Rows Length :: ${matchedArr.length}`);
}

convertAndWriteXLSX_AHD();
