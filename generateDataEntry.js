
const fs = require('fs');
const AdmZip = require('adm-zip');
const ExcelJS = require('exceljs');
const { json } = require('stream/consumers');
require('dotenv').config();

// path to the OCA bundle for examples
const directory = process.env.path;
const path = `${directory}/572eb71004e56e27e934b71a1cf400bc.zip`;

// Function to add column headers with styling to a worksheet
function addHeadersWithStyle(worksheet, attributes) {
    attributes.forEach((label, index) => {
      worksheet.getColumn(index + 1).header = label;
  
      worksheet.getCell(1, index + 1).style = {
        fill: {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFDDDDDD' },
        },
      };
    });

    worksheet.getRow(1).eachCell((cell) => {
      cell.border = {
        right: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
      };
    });
  }

async function generateDataEntry() {
    
    // Step 1: Unzip the OCA bundle
    const zip = new AdmZip(path);
    const zipEntries = zip.getEntries();
    
    // Step 2 and 3: Read JSON Files and Data Manipulation 
    const jsonData = [];
    let attributeNames;
    let attributeTypes;
    let flaggedAttributes;
    let attributeLabels;

    zipEntries.forEach((entry) => {
        if (entry.name.endsWith('.json')) { 
            const data = JSON.parse(entry.getData().toString('utf8'));
            jsonData.push(data);   
        };
    });

    // Step 4: Data Manipulation
    for (let overlay of jsonData) {
        if(overlay.type && overlay.type.endsWith('/capture_base/1.0')) {
            attributeNames = Object.keys(overlay.attributes);
            attributeTypes = Object.values(overlay.attributes);  
            flaggedAttributes = Object.values(overlay.flagged_attributes); 
        };

        if(overlay.type && overlay.type.endsWith('/label/1.0')) {
            attributeLabels = Object.values(overlay.attribute_labels);
        };
    }; 

    // Step 5: Write to Excel
    const workbook = new ExcelJS.Workbook();
    const sheet1 = workbook.addWorksheet('Schema Description');
    const sheet2 = workbook.addWorksheet('Data Entry');
    const sheet3 = workbook.addWorksheet('Schema conformant data');
    
    // Step 6: Write to Excel
    sheet1.addRow(["CB: CLASSIFICATION", "CB: Attribute Name"]);

    addHeadersWithStyle(sheet2, attributeLabels);
    addHeadersWithStyle(sheet3, attributeNames);
    
    // Save the Excel file
    await workbook.xlsx.writeFile(outputFilePath);

}

// Usage
const zipFilePath = path;
const outputFilePath = 'test.xlsx';

generateDataEntry(zipFilePath, outputFilePath)
  .then(() => console.log('Excel file created successfully'))
  .catch((error) => console.error('Error:', error));

