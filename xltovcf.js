const XLSX = require('xlsx');
const fs = require('fs');

// Function to add a plus sign to the front of the number
function addPlusToNumber(number) {
    return `+${number}`;
}

// Function to convert XLSX to VCF
function xlsxToVcf(inputFile, outputFile) {
  const workbook = XLSX.readFile(inputFile);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];

  // Define a mapping object for column names to vCard properties
  const columnMapping = {
    'First name': 'FN',
    'Last name': 'LN',
    'Phone number': 'TEL',
    // Add more mappings if needed
  };

  const data = XLSX.utils.sheet_to_json(worksheet);

  const vcards = [];

  for (const entry of data) {
    let vcard = 'BEGIN:VCARD\nVERSION:3.0\n';

    for (const columnName in entry) {
      if (columnMapping[columnName]) {
        const propertyValue = columnMapping[columnName] === 'TEL'
          ? addPlusToNumber(entry[columnName])  // Add a plus sign to the number
          : entry[columnName];
        vcard += `${columnMapping[columnName]}:${propertyValue}\n`;
      }
    }

    vcard += 'END:VCARD\n';

    vcards.push(vcard);
  }

  fs.writeFileSync(outputFile, vcards.join('\n'));

  console.log(`Conversion complete. VCF file saved as ${outputFile}`);
}

const inputFile = './XLSX/verifiedConvertToVcf.xlsx'; // Replace with your XLSX file path
const outputFile = 'Verified1858.vcf'; // Replace with the desired name for the VCF file

xlsxToVcf(inputFile, outputFile);
