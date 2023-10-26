
const fs = require('fs');
const AdmZip = require('adm-zip');
const ExcelJS = require('exceljs');
const { json } = require('stream/consumers');
const { start } = require('repl');
const { format } = require('path');
const { off } = require('process');
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
            const formula = `IF(ISBLANK('Data Entry'!$${letter}$${r + 1}), "", 'Data Entry'!$${letter}$${r + 1})`;

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
  let lang = [];;
  const lookupEntries = {}

  jsonData.forEach((overlay, i) => {
    if (overlay.type && overlay.type.includes('/character_encoding/')) {

      try {

        sheet1.getColumn(i + 4 - skipped).width = 15;
        sheet1.getCell(1, i + 4 - skipped).value = 'OL: Character Encoding';
        formatHeader2(sheet1.getCell(1, i + 4 - skipped));

        for (let row = 2; row <= attributeNames.length + 1; row++) {
          sheet1.getCell(row, i + 4 - skipped).value = null;
          formatAttr2(sheet1.getCell(row, i + 4 - skipped));
        }

        for (let [attrName, encoding] of Object.entries(overlay.attribute_character_encoding)) {

          if (typeof encoding == 'string') {
            const attrKeys = Object.keys(attributesIndex);
            const attrNameFromAttrKeys = attrKeys.map(key => key.split(',')[0]);
            const rowIndex = attrNameFromAttrKeys.indexOf(attrName) + 2;
            if (rowIndex) {
              sheet1.getCell(rowIndex, i + 4 - skipped).value = encoding;
            } 
          } 
        }

      } catch (error) {
        throw new WorkbookError('.. Error in formatting character encoding column (header and rows) ...');
      }
    } else if (overlay.type && overlay.type.includes('/cardinality/')) {
      try {

        sheet1.getColumn(i + 4 - skipped).width = 15;
        sheet1.getCell(1, i + 4 - skipped).value = 'OL: Cardinality';
        formatHeader2(sheet1.getCell(1, i + 4 - skipped));

        for (let row = 2; row <= attributeNames.length + 1; row++) {
          sheet1.getCell(row, i + 4 - skipped).value = null;
          formatAttr2(sheet1.getCell(row, i + 4 - skipped));
        }

        for (let [attrName, cardinality] of Object.entries(overlay.attribute_cardinality)) {

          const attrKeys = Object.keys(attributesIndex);
          const attrNameFromAttrKeys = attrKeys.map(key => key.split(',')[0]);
          const rowIndex = attrNameFromAttrKeys.indexOf(attrName) + 2;
          if (rowIndex) {
            sheet1.getCell(rowIndex, i + 4 - skipped).value = cardinality;
          } 
        }

      } catch (error) {
        throw new WorkbookError('.. Error in formatting cardinality column (header and rows) ...');
      }
    } else if (overlay.type && overlay.type.includes('/conformance/')) {
      try {

        sheet1.getColumn(i + 4 - skipped).width = 15;
        sheet1.getCell(1, i + 4 - skipped).value = 'OL: Conformance';
        formatHeader2(sheet1.getCell(1, i + 4 - skipped));

        for (let row = 2; row <= attributeNames.length + 1; row++) {
          sheet1.getCell(row, i + 4 - skipped).value = null;
          formatAttr2(sheet1.getCell(row, i + 4 - skipped));
        }

        for (let [attrName, conformance] of Object.entries(overlay.attribute_conformance)) {

          const attrKeys = Object.keys(attributesIndex);
          const attrNameFromAttrKeys = attrKeys.map(key => key.split(',')[0]);
          const rowIndex = attrNameFromAttrKeys.indexOf(attrName) + 2;
          if (rowIndex) {
            sheet1.getCell(rowIndex, i + 4 - skipped).value = conformance;
          } 
        }

      } catch (error) {
        throw new WorkbookError('.. Error in formatting conformance column (header and rows) ...');
      }
    } else if (overlay.type && overlay.type.includes('/conditional/')) {
      try {

        sheet1.getColumn(i + 4 - skipped).width = 15;
        sheet1.getCell(1, i + 4 - skipped).value = 'OL: Conditional [Condition]';
        formatHeader2(sheet1.getCell(1, i + 4 - skipped));

        sheet1.getColumn(i + 5 - skipped).width = 15;
        sheet1.getCell(1, i + 5 - skipped).value = 'OL: Conditional [Dependecies]';
        formatHeader2(sheet1.getCell(1, i + 5 - skipped));
        
        for (let row = 2; row <= attributeNames.length + 1; row++) {
          sheet1.getCell(row, i + 4 - skipped).value = null;
          sheet1.getCell(row, i + 5 - skipped).value = null;
          formatAttr2(sheet1.getCell(row, i + 4 - skipped));
          formatAttr2(sheet1.getCell(row, i + 5 - skipped));
        }

        for (let [attrName, condition] of Object.entries(overlay.attribute_conditions)) {

          const attrKeys = Object.keys(attributesIndex);
          const attrNameFromAttrKeys = attrKeys.map(key => key.split(',')[0]);
          const rowIndex = attrNameFromAttrKeys.indexOf(attrName) + 2;
          if (rowIndex) {
            sheet1.getCell(rowIndex, i + 4 - skipped).value = condition;
          } 
        }

        for (let [attrName, dependencies] of Object.entries(overlay.attribute_dependencies)) {

          const attrKeys = Object.keys(attributesIndex);
          const attrNameFromAttrKeys = attrKeys.map(key => key.split(',')[0]);
          const rowIndex = attrNameFromAttrKeys.indexOf(attrName) + 2;
          if (rowIndex) {
            sheet1.getCell(rowIndex, i + 5 - skipped).value = dependencies.join(",");
          } 
        }

        skipped -= 1;

      }
      catch (error) {
        throw new WorkbookError('.. Error in formatting conditional column (header and rows) ...');
      }
    } else if (overlay.type && overlay.type.includes('/format/')) {
      try {

        sheet1.getColumn(i + 4 - skipped).width = 15;
        sheet1.getCell(1, i + 4 - skipped).value = 'OL: Format';
        formatHeader2(sheet1.getCell(1, i + 4 - skipped));

        for (let row = 2; row <= attributeNames.length + 1; row++) {
          sheet1.getCell(row, i + 4 - skipped).value = null;
          formatAttr2(sheet1.getCell(row, i + 4 - skipped));
        }

        for (let [attrName, format] of Object.entries(overlay.attribute_formats)) {

          const attrKeys = Object.keys(attributesIndex);
          const attrNameFromAttrKeys = attrKeys.map(key => key.split(',')[0]);
          const rowIndex = attrNameFromAttrKeys.indexOf(attrName) + 2;
  
          if (rowIndex) {
            sheet1.getCell(rowIndex, i + 4 - skipped).value = format;
          }

          const attrTypeFromAttrKeys = attrKeys.map(key => key.split(','));
          const attrTypeObjects = attrTypeFromAttrKeys.map(([attr, type]) => ({ attr, type }));;

          for (let attrTypeObject of attrTypeObjects) {

            if (attrTypeObject.attr === attrName && attrTypeObject.type === "DateTime") {
             
              const format_attr = { numFmt: 'yyyy-mm-dd' };
              const col_i = attributesIndex[[attrName, attrTypeObject.type]] - 1;
              const letter = String.fromCharCode(65 + col_i -1);

              for (let r=1; r <= 1000; r++) {
                sheet2.getCell(r+1, col_i).value = null;
                sheet2.getCell(r+1, col_i).numFmt = format_attr.numFmt;
                sheet3.getCell(r+1, col_i).numFmt = format_attr.numFmt;
  
              }
            }
          }
        
        }

      } catch (error) {
        throw new WorkbookError('.. Error in formatting format column (header and rows) ...');
      }
    } else if (overlay.type && overlay.type.includes('/entry_code/')) {

      try {
        sheet1.getColumn(i + 4 - skipped).width = 15;
        sheet1.getCell(1, i + 4 - skipped).value = 'OL: Entry Code';
        formatHeader2(sheet1.getCell(1, i + 4 - skipped));

        for (let row = 2; row <= attributeNames.length + 1; row++) {
          sheet1.getCell(row, i + 4 - skipped).value = null;
          formatAttr2(sheet1.getCell(row, i + 4 - skipped));
        }


        for (let [attrName, entryCode] of Object.entries(overlay.attribute_entry_codes)) {

          if (Array.isArray(entryCode)) {
            const joinedCodes = entryCode.join('|');
            const attrKeys = Object.keys(attributesIndex);
            const attrNameFromAttrKeys = attrKeys.map(key => key.split(',')[0]);
            const rowIndex = attrNameFromAttrKeys.indexOf(attrName) + 2;
            if (rowIndex) {
              sheet1.getCell(rowIndex, i + 4 - skipped).value = joinedCodes;
            }
          }
  
        }

      } catch (error) {
        throw new WorkbookError('.. Error in formatting entry code column (header and rows) ...');
      }

    } else if (overlay.type && overlay.type.includes('/label/')) {
      lang.push(overlay);
      let attr_labels = null;

      const o = lang.find(overlay => overlay.language === 'en');

      if (o) {
        attr_labels = o.attribute_labels;
      } else {
        attr_labels = lang[0].attribute_labels;
      }

      if (o) {
        try {
          sheet1.getColumn(i + 4 - skipped).width = 17;
          sheet1.getCell(1, i + 4 - skipped).value = 'OL: Label';
          formatHeader2(sheet1.getCell(1, i + 4 - skipped));

          for (let row = 2; row <= attributeNames.length + 1; row++) {
            sheet1.getCell(row, i + 4 - skipped).value = null;
            formatAttr2(sheet1.getCell(row, i + 4 - skipped));
          }
          
          for (let [attrName, label] of Object.entries(attr_labels)) {

            const attrKeys = Object.keys(attributesIndex);
            const attrNameFromAttrKeys = attrKeys.map(key => key.split(',')[0]);
            const rowIndex = attrNameFromAttrKeys.indexOf(attrName) + 2;
            if (rowIndex) {
              sheet1.getCell(rowIndex, i + 4 - skipped).value = label;
            } 

            const labelValue = Object.values(attr_labels);
            sheet2.getRow(1).values = labelValue;

            labelValue.forEach((label, index) => {
              const cell = sheet2.getCell(1, index + 1);
              formatDataHeader(cell);
            });          
          }
        } catch (error) {
          throw new WorkbookError('.. Error in formatting labels code column (header and rows) ...');
        }
        lang.length = 0;
      } else {
        skipped += 1;
      }
    } else if (overlay.type && overlay.type.includes('/entry/')) {
      lang.push(overlay);
      let attr_labels = null;

      const o = lang.find(overlay => overlay.language === 'en');


      if (o) {
        attr_labels = o.attribute_entries;
      } else {
        attr_labels = lang[0].attribute_entries;
      }

      if (o) {
        try{
          sheet1.getColumn(i + 4 - skipped).width = 20;
          sheet1.getCell(1, i + 4 - skipped).value = 'OL: Entry';
          formatHeader2(sheet1.getCell(1, i + 4 - skipped));

          for (let row = 2; row <= attributeNames.length + 1; row++) {
            sheet1.getCell(row, i + 4 - skipped).value = null;
            formatAttr2(sheet1.getCell(row, i + 4 - skipped));
          }

          for (let [attrName, entries] of Object.entries(attr_labels)) {

            if (entries !== undefined && entries !== null && entries instanceof Object) {
              lookupEntries[attrName] = entries;
              const attrKeys = Object.keys(attributesIndex);
              const attrNameFromAttrKeys = attrKeys.map(key => key.split(',')[0])
              const rowIndex = attrNameFromAttrKeys.indexOf(attrName) + 2;

              const formattedEntries = [];

              for (const [key, value] of Object.entries(entries)) {
                formattedEntries.push(`${key}:${value}`);
              }

              const formattedEntryString = formattedEntries.join('|');
              if (rowIndex) {
                sheet1.getCell(rowIndex, i + 4 - skipped).value = formattedEntryString;
              }
            }
          }

        } catch (error) {
          throw new WorkbookError('.. Error in formatting entry column (header and rows) ...', error.message);
          // console.error('Error in formatting entry column (header and rows):', error.message);

        }
        lang.length = 0;
      } else {
        skipped += 1;
      }
    } else if (overlay.type && overlay.type.includes('/information/')) {
      
      lang.push(overlay);
      let attr_labels = null;

      const o = lang.find(overlay => overlay.language === 'en');

      if (o) {
        attr_labels = o.attribute_information;
      } else {
        attr_labels = lang[0].attribute_information;
      }

      if (o) {
        try {
          sheet1.getColumn(i + 4 - skipped).width = 20;
          sheet1.getCell(1, i + 4 - skipped).value = 'OL: Information';
          formatHeader2(sheet1.getCell(1, i + 4 - skipped));

          for (let row = 2; row <= attributeNames.length + 1; row++) {
            sheet1.getCell(row, i + 4 - skipped).value = null;
            formatAttr2(sheet1.getCell(row, i + 4 - skipped));
          }

          for (let [attrName, info] of Object.entries(attr_labels)) {
            const attrKeys = Object.keys(attributesIndex);
            const attrNameFromAttrKeys = attrKeys.map(key => key.split(',')[0]);
            const rowIndex = attrNameFromAttrKeys.indexOf(attrName) + 2;
            if (rowIndex) {
              sheet1.getCell(rowIndex, i + 4 - skipped).value = info;
            } 
          }

        } catch (error) {
          throw new WorkbookError('.. Error in formatting information column (header and rows) ...');
        }
        lang.length = 0;
      } else {
        skipped += 1;
      }
    }
  });

  // lookup table
  const lookUpTable = new Map();
  let lookUpStart = attributeNames.length + 6;

  sheet1.getCell(lookUpStart, 1).value = 'Lookup tables';
  formatLookupHeader(sheet1.getCell(lookUpStart, 1));

  sheet1.getCell(lookUpStart, 2).value = null;
  formatLookupHeader(sheet1.getCell(lookUpStart, 2));

  let offset = 0; // used to offset the lookup table rows

  for (const [attrName, entries] of Object.entries(lookupEntries)) {

    sheet1.getCell(lookUpStart + offset + 1, 1).value = attrName;
    formatLookupAttr(sheet1.getCell(lookUpStart + offset + 1, 1));

    lookUpTable.set(attrName, {
      start : lookUpStart + 3 + offset,
      end: lookUpStart + 2 + offset + Object.entries.length,
    });

    offset += Object.entries.length + 2;
  }

  console.log(lookUpTable);


  // let offset = 0;
  // for ([attrName, entries] of Object.entries(lookupEntries)) {
  
  //   // sheet1.getCell(lookUpStart + offset + 1, 1).value = attrName;
  //   // formatLookupAttr(sheet1.getCell(lookUpStart + offset + 1, 1));

  //   sheet1.getCell(lookUpStart, 0).value = attrName;
  //   formatLookupAttr(sheet1.getCell(lookUpStart, 0));

  //   lookUpTable.set(attrName, {
  //     start: lookUpStart + 2 + offset,
  //     end: lookUpStart + 1 + offset + entries.size,
  //   });

  //   let i = 0;
  //   for (const [k, v] of entriesMap.entries()) {
  //     sheet1.getCell(lookUpStart + 2 + offset + i, 0).value = v;
  //     sheet1.getCell(lookUpStart + 2 + offset + i, 1).value = k;
  //     i++;
  //   }
  //   offset += entriesMap.size + 2;

  //   // lookUpStart++;
  // }

  // // for ([attrName, entries] of Object.entries(lookupEntries)) {

  // //   const values = []
  // //   lookUpStart + 3 + offset,
  // //   lookUpStart + 2 + offset + Object.keys(entries).length

  // //   lookUpTable[attrName] = values;
  // // }
  




  return workbook;
}

// path to the OCA bundle for examples
const directory = process.env.path;
const filename = 'a5cbe768bee30be3638f434cd46d22eb.zip';
// const filename = 'OCA_test.zip';
// const filename = '572eb71004e56e27e934b71a1cf400bc.zip';
// const filename = '9f103493cbe64733919f00d3768e6ba5.zip';
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

