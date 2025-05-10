const fs = require('fs').promises;
const { Pool } = require('pg');
const XLSX = require('xlsx');
const path = require('path');

const backwallDataPath = path.join(__dirname, 'YSTRAW.xlsx');
// const newBackwallPath = path.join(__dirname, 'NEW_DATA_SHEET.xlsx');

const backwallDataWorkbook = XLSX.readFile(backwallDataPath);
// const newBackwallWorkbook = XLSX.readFile(newBackwallPath);


const sheetNamed1 = backwallDataWorkbook.SheetNames[0];
// const sheetNamed2 = newBackwallWorkbook.SheetNames[0];


const sheet1 = backwallDataWorkbook.Sheets[sheetNamed1];
// const sheet2 = newBackwallWorkbook.Sheets[sheetNamed2];

let backwallDataSheet = XLSX.utils.sheet_to_json(sheet1);
// let newBackwallSheet = XLSX.utils.sheet_to_json(sheet2);

var wb = XLSX.utils.book_new();

const pool = new Pool({
    user: "postgres",
    host: 'db.mgampbhmlnalxohuobpr.supabase.co',
    database: "postgres",
    password: 'gplVhDuxLDMeBKxs',
    port: 5432,
});

let today = new Date();
let oneDayMilliseconds = 24 * 60 * 60 * 1000;
let oneDayBefore = new Date(today.getTime() - oneDayMilliseconds);
let year = oneDayBefore.getFullYear();
let month = String(oneDayBefore.getMonth() + 1).padStart(2, "0");
let day = String(oneDayBefore.getDate()).padStart(2, "0");
let firstDate = `${year}-${month}-01`;
let current_date = `${year}-${month}-${day}`;
let dateDifference = new Date(current_date).getDate() + 1 - new Date(firstDate).getDate();

let newArr = [];
let matchedArr = [];
let notMatchedArr = [];
let notMDeviceId = [];

let response1;
let response2;


async function querydb() {
    response1 = await pool.query(`select * from ystraw_data_table  where custom_date = '${current_date}'`);
    convertAndWriteXLSX_AHD();
}

function extractDeviceId(text) {
    const match = text.match(/\b[A-Z]{2}-\d{3,4}\b/i);
    return match ? match[0] : null;
}

const convertAndWriteXLSX_AHD = async () => {
    console.log("Actual Length :: ", backwallDataSheet.length, response1.rows.length);

    backwallDataSheet?.forEach((elmData) => {
        const filteredData = response1.rows.find(d => extractDeviceId(d.display_name) == elmData['Device ID']);
        // console.log(filteredData);

        if (filteredData) {
            elmData['Display ID'] = filteredData?.display_id;
            newArr.push(elmData);
            matchedArr.push(elmData);
        } else {
            newArr.push(elmData);
            notMatchedArr.push(elmData);
        }
    });

    // console.log(`updatedArr :: ${newArr.length} , baseFileArr :: ${backwallDataSheet.length}`);
    console.log(`matchedArr :: ${matchedArr.length} , notMatchedArr :: ${notMatchedArr.length}`);

    // console.log(notMDeviceId, notMDeviceId.length);

    // Create a new worksheet and add the data
    var ws = XLSX.utils.json_to_sheet(newArr);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'NEW-LIST-YSTRAW.xlsx');
    console.log('Task Done...');

    // const updatedSheet = XLSX.utils.json_to_sheet(newArr);
    // const updatedWorkbook = XLSX.utils.book_new();
    // XLSX.utils.book_append_sheet(updatedWorkbook, updatedSheet, sheetNamed1);
    // await fs.writeFile(backwallDataPath, XLSX.write(updatedWorkbook, { bookType: 'xlsx', type: 'buffer' }));
    // console.log(`Updated Rows Length :: ${matchedArr.length}`);
}

querydb();