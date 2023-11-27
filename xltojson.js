const XLSX = require('xlsx');
const fs = require('fs');

function excelToJson(excelFilePath, jsonFilePath) {
  try {
    // Load the Excel file
    const workbook = XLSX.readFile(excelFilePath);

    // Initialize an object to store data for all sheets
    const allSheetData = {};

    // Loop through all sheets in the workbook
    workbook.SheetNames.forEach(sheetName => {
      // Get the sheet
      const worksheet = workbook.Sheets[sheetName];

      // Convert the worksheet data to JSON
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Use the first row as keys
      const keys = jsonData[0];

      // Initialize an array to store the data rows
      const data = [];

      // Loop through the data rows, starting from the second row
      for (let i = 1; i < jsonData.length; i++) {
        const rowData = {};
        keys.forEach((key, index) => {
          rowData[key] = jsonData[i][index];
        });
        data.push(rowData);
      }

      // Store the data for this sheet using the sheet name as the key
      allSheetData[sheetName] = data;
    });

    // Write the JSON data to a file
    fs.writeFileSync(jsonFilePath, JSON.stringify(allSheetData, null, 2));

    console.log('Conversion completed. JSON file saved.');
  } catch (error) {
    console.error('Error converting Excel to JSON:', error);
  }
}

// Example usage:
const excelFile = './XLSX/verifiedConvertToJson.xlsx'; // Replace with the path to your Excel file
const jsonFile = 'output.json';   // Replace with the desired JSON output file path

excelToJson(excelFile, jsonFile);
