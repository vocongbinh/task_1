const axios = require("axios");
const XlsxPopulate = require("xlsx-populate");
const fs = require("fs");
const url = "https://go.microsoft.com/fwlink/?LinkID=521962";

async function downloadFile(url, filePath) {
  const response = await axios({
    method: "GET",
    url: url,
    responseType: "arraybuffer",
  });
  fs.writeFileSync(filePath, Buffer.from(response.data));
}

function getIndexOfSales(headerRow) {
  let salesIndex = -1;
  headerRow._cells.forEach((cel) => {
    if (cel._value === " Sales") {
      salesIndex = cel._columnNumber;
    }
  });
  return salesIndex;
}
function filterRows(sheet) {
  const headerRow = sheet._rows[1];
  const salesIndex = getIndexOfSales(headerRow)
  const rows = sheet.usedRange().value();
  let filteredRows = rows.filter((row) => row[salesIndex - 1] > 50000);
  filteredRows.unshift(rows[0]);
  return filteredRows;
}

async function createFilteredSheet(url, newFilePath) {
  try {
    const originFilePath = "origin.xlsx";
    await downloadFile(url, originFilePath);
    const workbook = await XlsxPopulate.fromFileAsync(originFilePath);
    const sheet = workbook.sheet(0);
    const filteredRows = filterRows(sheet);
    sheet.usedRange().clear();
    filteredRows.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        sheet.cell(rowIndex + 1, colIndex + 1).value(value);
      });
    });
    await workbook.toFileAsync(newFilePath);
    console.log(`Filtered sheet saved to ${newFilePath}`);
  } catch (error) {
    console.error("Error:", error);
  }
}

const newFile = "filtered_sheet.xlsx";
createFilteredSheet(url, newFile);
