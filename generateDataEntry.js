
const fs = require('fs');
const AdmZip = require('adm-zip');
const ExcelJS = require('exceljs');
const { json } = require('stream/consumers');
const { start } = require('repl');
require('dotenv').config();

// Custom error-handling function
function WorkbookError(message) {
  this.name = 'WorkbookError';
  this.message = message;
  this.stack = new Error().stack;
}

WorkbookError.prototype = Object.create(Error.prototype);
WorkbookError.prototype.constructor = WorkbookError;

function generateDataEntry(path, outputFilePath) {
  // Step 1: Unzip the OCA bundle
  const zip = new AdmZip(path);
  const zipEntries = zip.getEntries();

  const jsonData = [];
  
  // Step 2: Read JSON Files
  zipEntries.forEach((entry) => {
    if (entry.name.endsWith('.json')) { 
        const data = JSON.parse(entry.getData().toString('utf8'));
        jsonData.push(data);   
    };
  });

  // Step 3: Create a new Excel workbook

  const workbook = new ExcelJS.Workbook();
 
  // Step 4: Format function
  function formatHeader1(cell) {
    cell.font = { size: 10, bold: true };
    cell.alignment = { vertical: 'top', wrapText: true };
    cell.border = { bottom: { style: 'thin' } };
  };
  
  function formatHeader2(cell) {
    cell.font = { size: 10, bold: true };
    cell.alignment = { vertical: 'top', wrapText: true };
    cell.border = { bottom: { style: 'thin' }, right: { style: 'thin' } };
  };
  
  function formatAttr1(cell) {
    cell.font = { size: 10 };
    cell.alignment = { vertical: 'top', wrapText: true };
  };
  
  function formatAttr2(cell) {
    cell.font = { size: 10 };
    cell.alignment = { vertical: 'top', wrapText: true };
    cell.border = {right: { style: 'thin' } };
  };
  
  function formatDataHeader(cell) {
    cell.font = { size: 10 };
    cell.alignment = { vertical: 'top', wrapText: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'E7E6E6' },
    };
    cell.border = {
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
  };
  
  function formatLookupHeader(cell) {
    cell.font = { size: 10, bold: true };
    cell.alignment = { vertical: 'top', wrapText: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'E7E6E6' },
    };
  };
  
  function formatLookupAttr(cell) {
    cell.font = { size: 10, bold: true };
    cell.alignment = { vertical: 'top', wrapText: true };
  };
  
  function formatLookupValue(cell) {
    cell.font = { size: 10 };
    cell.alignment = { vertical: 'top', wrapText: true };
  };

  
  const sheet1 = workbook.addWorksheet('Schema Description'); 
 
  try {
    sheet1.getRow(1).height = 35;
    sheet1.getColumn(1).width = 13;
    sheet1.getCell(1, 1).value = 'CB: Classification';
    formatHeader1(sheet1.getCell(1, 1));

    sheet1.getColumn(2).width = 17;
    sheet1.getCell(1, 2).value = 'CB: Attribute Name';
    formatHeader1(sheet1.getCell(1, 2));

    sheet1.getColumn(3).width = 12.5;
    sheet1.getCell(1, 3).value = 'CB: Attribute Type';
    formatHeader1(sheet1.getCell(1, 3));

    sheet1.getColumn(4).width = 17;
    sheet1.getCell(1, 4).value = 'CB: Flagged Attribute';
    formatHeader2(sheet1.getCell(1, 4));

  } catch (error) {
    throw new WorkbookError('.. Error in formatting sheet1 capture base header ...');
  }

  const sheet2 = workbook.addWorksheet('Data Entry');
  const sheet3 = workbook.addWorksheet('Schema conformant data');

  const attributesIndex = {};
  let attributeNames = null;

  jsonData.forEach((overlay) => {
    if (overlay.type && overlay.type.includes('/capture_base/')) {
      Object.entries(overlay.attributes).forEach(([attrName, attrType], index) => {
        const attrIndex = index + 2;
        attributesIndex[[attrName, attrType]] = attrIndex;
        sheet1.getCell(attrIndex, 1).value = overlay.classification;
        formatAttr2(sheet1.getCell(attrIndex, 1));

        if (attrIndex !== undefined) {
          sheet1.getCell(attrIndex, 2).value = attrName;
          formatAttr2(sheet1.getCell(attrIndex, 2));
        } else {
          throw new WorkbookError('.. Error check the attribute name ...');
        }

        if (attrIndex !== undefined) {
          sheet1.getCell(attrIndex, 3).value = attrType;
          formatAttr2(sheet1.getCell(attrIndex, 3)); 
        } else {
          throw new WorkbookError('.. Error check the attribute type ...');
        }

        const isFlagged = overlay.flagged_attributes.includes(attrName);
        sheet1.getCell(attrIndex, 4).value = isFlagged ? "Y" : "";
        formatAttr2(sheet1.getCell(attrIndex, 4));
      });

      // sheet 3
      attributeNames = Object.keys(overlay.attributes)
      
      try { 
          sheet3.getRow(1).values = attributeNames;
          attributeNames.forEach((attrName, index) => {
            const cell = sheet3.getCell(1, index + 1);
            formatDataHeader(cell);
          });
        } catch (error) {
          throw new WorkbookError('.. Error in formatting sheet3 data header ...');
        }

      try {
        for (let i = 0; i < attributeNames.length; i++) {
          const letter = String.fromCharCode(65 + i);
          
          for (let r = 1; r <= 1000; r++) {
            const formula = `IF(ISBLANK('Data Entry'!${letter}${r + 1}), "", 'Data Entry'!${letter}${r + 1})`;
            const cell = sheet3.getCell(r + 1, i + 1);
            cell.value = {
              formula: formula,
            };
          }
        }

        } catch (error) {
          throw new WorkbookError('.. Error in creating the formulae sheet3 data ...');
        }

        const numColumns = attributeNames.length;
        const columnWidth = 12;

        for (let i = 0; i < numColumns; i++) {
          sheet2.getColumn(i + 1).width = columnWidth;
        }
        
        for (let i = 0; i < numColumns; i++) {
          sheet3.getColumn(i + 1).width = columnWidth;
        }

    }
  });

  
  let skipped = 0;

  jsonData.forEach((overlay, i) => {
    if (overlay.type && overlay.type.includes('/character_encoding/')) {

      const startColumn = i + 4 - skipped;
      const endColumn = startColumn;

      try {

        sheet1.getColumn(startColumn).width = 15;
        sheet1.getCell(1, startColumn).value = 'OL: Character Encoding';
        formatHeader2(sheet1.getCell(1, startColumn));

        for (let row = 2; row <= attributeNames.length + 1; row++) {
          sheet1.getCell(row, startColumn).value = null;
          formatAttr2(sheet1.getCell(row, startColumn));
        }

        for (let [attrName, encoding] of Object.entries(overlay.attribute_character_encoding)) {

          if (typeof encoding == 'string') {
            console.log('attrName', attrName);

            // const rowIndex = attributesIndex[attrName];
            // console.log('rowIndex', rowIndex);
            // if (rowIndex) {
            //   sheet1.getCell(rowIndex, startColumn).value = encoding;
            // }

          }

        }

      } catch (error) {
        throw new WorkbookError('.. Error in formatting character encoding column (header and rows) ...');
      }
    }
  });
  
  return workbook;

}

// path to the OCA bundle for examples
const directory = process.env.path;
const filename = 'a5cbe768bee30be3638f434cd46d22eb.zip';
// const filename = '62434bb0350d9c5b7b8f6b4d52bfed8f.zip';
const path = `${directory}/${filename}`;
const outputFilePath = `${filename.split('.')[0]}_data_entry.xlsx`;

// Generate the workbook and handle errors
async function generateAndSaveDateEntry() {
  try {
    const generatedWorkbook = generateDataEntry(path, outputFilePath);
    await generatedWorkbook.xlsx.writeFile(outputFilePath);
    console.log(' ... Date Entry file created successfully ...');
  } catch (error) {
    console.error('Custom Error:', error.message);
    if (error instanceof WorkbookError) {
      console.error('Custom Error:', error.message);
    } else {
      console.error('Error:', error);
    }
  }
}
// Generate and save the workbook
generateAndSaveDateEntry(path, outputFilePath);

