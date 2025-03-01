function main() {
  const folderId = '1CkJhpVmEgPndCVSb6ckBkBK-ijQBaHx8'; // Folder containing converted Google Sheets
  const folder = DriveApp.getFolderById(folderId);

  Logger.log('Adding Files...');
  // Convert and sort files first
  const fileList = convertAndSortFiles(folder);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Tracker') || ss.insertSheet('Tracker');

  masterSheet.clear();
  setupMasterSheet(masterSheet, fileList);

  const consolidatedData = [];

  Logger.log('Processing Files...');
  // Consolidate data for all PAL sheets in an array
  fileList.forEach(spreadsheet => {
    handleFile(spreadsheet, consolidatedData);
  });
  
  Logger.log('Sorting Data...');
  // Sort consolidated data before writing it to the master sheet
  sortConsolidatedData(consolidatedData);
  
  Logger.log('Adding Data to Sheet...');
  // Write consolidated data to the master sheet in one go
  writeConsolidatedData(masterSheet, consolidatedData);

  Logger.log('Applying Number Formatting to Data...');
  // Apply number formatting for currency and percentages
  const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0]; // Retrieve headers
  applyNumberFormatting(masterSheet, headers);
  
  Logger.log('Applying Centering and Bold Formatting to Data...');
  // Apply centering and bold formatting before conditional formatting
  applyCenteringAndBoldFormatting(masterSheet);
  
  Logger.log('Applying Conditional Formatting to Data...');
  // Apply conditional formatting after all data is written
  applyConditionalFormatting(masterSheet, headers);
  Logger.log('All done')
}

// Convert Excel files to Google Sheets and sort them in natural order (P1, P2, ..., P12)
function convertAndSortFiles(folder) {
  const files = folder.getFiles();
  const fileList = [];

  while (files.hasNext()) {
    const file = files.next();

    // If the file is already a Google Spreadsheet, just add it to the list
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      fileList.push(file);  // Add DriveApp file
      //Logger.log(`Added Google Spreadsheet directly: ${file.getName()}`);
    }
    // If the file is an Excel file, convert and add it to the list
    else if (file.getMimeType() === MimeType.MICROSOFT_EXCEL) {
      try {
        const convertedFile = convertExcelToSheet(file, folder.getId());
        if (convertedFile) {
          fileList.push(convertedFile); // Add converted DriveApp file
          //Logger.log(`Converted and added file: ${convertedFile.getName()}`);
        }
      } catch (e) {
        Logger.log(`Error converting file: ${file.getName()} - ${e.message}`);
      }
    } else {
      Logger.log(`Skipped file: ${file.getName()} due to incompatible MIME type.`);
    }
  }
  // Sort files by name depending on the condition for year-end and new cycle
  if (fileList.length === 2 && fileList.some(file => file.getName().includes('P1')) && fileList.some(file => file.getName().includes('P13'))) {
    // If only P1 and P13 are found, indicating year-end and new promo cycle, sort in reverse order
    fileList.sort((a, b) => naturalSort(b.getName(), a.getName()));
    Logger.log("Sorted in reverse natural order for year-end transition.");
  } else {
    // Sort in natural order
    fileList.sort((a, b) => naturalSort(a.getName(), b.getName()));
    Logger.log("Sorted in natural order.");
  }
  return fileList;
}

function convertExcelToSheet(file, targetFolderId) {
  try {
    // Prepare the metadata for creating the Google Sheet
    const resource = {
      name: file.getName().replace(/\.xlsx?$/, ''), // Remove .xls or .xlsx extension
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [targetFolderId]  // Specify the target folder
    };

    const blob = file.getBlob();
    const newFile = Drive.Files.create(resource, blob, {
      convert: true
    });
    file.setTrashed(true);
    return DriveApp.getFileById(newFile.id);
  } catch (e) {
    Logger.log(`Failed to convert file ${file.getName()}: ${e.message}`);
    return null;
  }
}

// Natural sort function
function naturalSort(a, b) {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
}

// Setup the master sheet dynamically based on the number of PAL files
function setupMasterSheet(sheet, fileList) {
  const headers = [
    'Type', 'BB Code', 'Description', 'Category', 'Sub Category', 'Case Size/Case PK', 'Del Out'
  ];

  // Add headers dynamically for each PAL file
  fileList.forEach((spreadsheet) => {
    const fileName = spreadsheet.getName();
    const period = fileName.match(/P\d+/); // Extract "P" followed by digits
    const periodLabel = period ? period[0] : fileName; // Default to filename if no match

    headers.push(`${periodLabel} W/SALE COST`);
    headers.push(`${periodLabel} SINGLE RETAIL`);
    headers.push(`${periodLabel} SINGLE MARGIN`);
    headers.push(`${periodLabel} PROMO`);
    headers.push(`${periodLabel} STATUS`);
  });

  sheet.appendRow(headers);
}

function handleFile(spreadsheet, consolidatedData) {
  try {
    const spreadsheetObj = SpreadsheetApp.openById(spreadsheet.getId());
    const sheetsToProcess = ['BWS PAL', 'Grocery PAL'];

    sheetsToProcess.forEach(sheetName => {
      const sheet = spreadsheetObj.getSheetByName(sheetName);
      if (sheet) {
        //Logger.log(`Processing sheet: ${sheet.getName()} from file: ${spreadsheetObj.getName()}`);
        // Add the sheet type (either 'BWS' or 'Grocery')
        const sheetType = sheetName === 'BWS PAL' ? 'BWS' : 'Grocery';
        processSheet(sheet, consolidatedData, spreadsheetObj.getName(), sheetType);
      }
    });

  } catch (e) {
    Logger.log(`Error processing file: ${spreadsheet.getName()} - ${e.message}`);
  }
}

function processSheet(sourceSheet, consolidatedData, fileName, sheetType) {
  const data = sourceSheet.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log(`Insufficient data in sheet: ${sourceSheet.getName()} to extract headers.`);
    return;
  }

  const row1 = data[0];
  const row2 = data[1];
  const combinedHeaders = row1.map((header, index) => {
    return row2[index] && row2[index].trim() ? row2[index] : header;
  });

  // Extract the header mappings, trimmed and normalized for consistency
  const headerMap = {};
  combinedHeaders.forEach((header, index) => {
    const trimmedHeader = header.trim().toUpperCase().replace(/\s+/g, ' ');
    if (trimmedHeader) {
      // Add more generic matching for the cost column
      if (trimmedHeader.includes('COST') || trimmedHeader.includes('W/SALE')) {
        headerMap['W/SALE COST'] = index;  // Broadly identify wholesale cost-related headers
      } else {
        headerMap[trimmedHeader] = index;
      }
    }
  });

  const periodMatch = fileName.match(/P\d+/);
  const period = periodMatch ? periodMatch[0] : fileName;

  let currentCategory = '';
  let currentSubCategory = '';

  for (let i = 2; i < data.length; i++) {
    const row = data[i];

    if (row[0] && row.slice(1).every(cell => !cell)) {
      currentCategory = row[0].trim();
      currentSubCategory = '';
      continue;
    }

    if (headerMap['DESCRIPTION'] !== undefined && row[headerMap['DESCRIPTION']] && !row[headerMap['BB CODE']]) {
      currentSubCategory = row[headerMap['DESCRIPTION']].trim();
      continue;
    }

    if (headerMap['BB CODE'] !== undefined && row[headerMap['BB CODE']]) {
      const bbCode = row[headerMap['BB CODE']];

      // Search for existing entry by BB Code
      let productData = null;
      for (let j = 0; j < consolidatedData.length; j++) {
        if (consolidatedData[j]['BB Code'] === bbCode) {
          productData = consolidatedData[j];
          break;
        }
      }

      if (!productData) {
        productData = {
          "BB Code": bbCode,
          "Description": row[headerMap['DESCRIPTION']] || '',
          "Category": currentCategory,
          "Sub Category": currentSubCategory,
          "Case Size/Case PK": row[headerMap['CASE SIZE']] || row[headerMap['CASE PK']] || '',
          "Del Out": headerMap['DEL OUT'] !== undefined ? row[headerMap['DEL OUT']] : '',
          "Type": sheetType // Add the 'Type' tag (BWS or Grocery)
        };
        consolidatedData.push(productData);
      }

      // Extract W/SALE COST, SINGLE RETAIL, SINGLE MARGIN, PROMO, STATUS
      productData[`${period} W/SALE COST`] = headerMap['W/SALE COST'] !== undefined ? row[headerMap['W/SALE COST']] : '';
      productData[`${period} SINGLE RETAIL`] = headerMap['SINGLE RETAIL'] !== undefined ? row[headerMap['SINGLE RETAIL']] : '';

      // Dynamically determine if SINGLE MARGIN needs adjustment
      if (headerMap['SINGLE MARGIN'] !== undefined) {
        let singleMarginValue = parseFloat(row[headerMap['SINGLE MARGIN']]);
        if (!isNaN(singleMarginValue)) {
          // If margin is greater than 1, it's likely meant to be a percentage and should be divided by 100
          if (singleMarginValue > 1) {
            singleMarginValue /= 100;
          }
          productData[`${period} SINGLE MARGIN`] = singleMarginValue;
        } else {
          productData[`${period} SINGLE MARGIN`] = row[headerMap['SINGLE MARGIN']] || '';
        }
      }

      productData[`${period} PROMO`] = headerMap['PROMO'] !== undefined ? row[headerMap['PROMO']] : '';
      productData[`${period} STATUS`] = headerMap['STATUS'] !== undefined ? row[headerMap['STATUS']] : '';
    }
  }
}

function writeConsolidatedData(masterSheet, consolidatedData) {
  // Retrieve headers dynamically from the master sheet
  const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];

  // Extend headers dynamically for cost difference and value efficiency
  const updatedHeaders = headers.flatMap((header) => {
    if (header.includes('W/SALE COST')) {
      const period = header.split(' ')[0]; // Extract period
      return [`${period} COST DIFF`, `${period} VALUE EFFICIENCY PER ITEM`, header];
    }
    return header;
  });

  // Prepare consolidated data with cost difference and value efficiency
  const dataToWrite = consolidatedData.map((productData) => {
    return updatedHeaders.map((header) => {
      if (header === 'Type') {
        return productData['Type'] || ''; // Add 'Type' tag
      }
      if (header.includes('COST DIFF')) {
        const { costDiff } = calculateCostDiff(header, productData, updatedHeaders);
        return costDiff; // Return only cost difference
      }
      if (header.includes('VALUE EFFICIENCY PER ITEM')) {
        const period = header.split(' ')[0];
        const costDiffHeader = `${period} COST DIFF`;
        const { valueEfficiency } = calculateCostDiff(costDiffHeader, productData, updatedHeaders);
        return valueEfficiency; // Return only value efficiency
      }
      return productData[header] || ''; // Use existing data for non-difference columns
    });
  });

  // Write headers and data to the master sheet
  masterSheet.clear();
  masterSheet.appendRow(updatedHeaders);
  if (dataToWrite.length > 0) {
    masterSheet
      .getRange(2, 1, dataToWrite.length, updatedHeaders.length)
      .setValues(dataToWrite);
  }
}

function calculateCostDiff(header, productData, headers) {
  const period = header.split(' ')[0]; // Extract the current period (e.g., "P2")
  const currentWholesaleHeader = `${period} W/SALE COST`;

  // Identify the previous period
  const currentIndex = headers.findIndex((h) => h === currentWholesaleHeader);
  const prevIndex = currentIndex - 7; // Adjust for column spacing

  // Skip calculation for the first promotion file or invalid prior data
  if (
    prevIndex < 8 ||
    !headers[prevIndex] ||
    !productData[headers[prevIndex]] ||
    !headers[currentIndex] ||
    !productData[currentWholesaleHeader]
  ) {
    return { costDiff: '', valueEfficiency: '' }; // Return blanks if no prior data
  }

  const prevWholesaleHeader = headers[prevIndex]; // Get the previous period's wholesale header
  const currentCost = parseFloat(productData[currentWholesaleHeader]) || 0;
  const prevCost = parseFloat(productData[prevWholesaleHeader]) || 0;

  const costDiff = currentCost - prevCost;

  // Calculate value efficiency
  const caseSize = parseFloat(productData['Case Size/Case PK']) || 0;
  const valueEfficiency =
    caseSize > 0 && costDiff !== 0 ? (costDiff / caseSize).toFixed(2) : '';

  // Return both costDiff and valueEfficiency
  return {
    costDiff: costDiff !== 0 ? costDiff.toFixed(2) : '',
    valueEfficiency,
  };
}

// Sort the consolidated data by Type, Category, Sub Category, Description, BB Code
function sortConsolidatedData(consolidatedData) {
  consolidatedData.sort((a, b) => {
    return (
      String(a['Type'] || '').localeCompare(String(b['Type'] || '')) ||
      String(a['Category'] || '').localeCompare(String(b['Category'] || '')) ||
      String(a['Sub Category'] || '').localeCompare(String(b['Sub Category'] || '')) ||
      String(a['Description'] || '').localeCompare(String(b['Description'] || '')) ||
      String(a['BB Code'] || '').localeCompare(String(b['BB Code'] || ''))
    );
  });
}

function applyNumberFormatting(sheet, headers) {
  const numRows = sheet.getLastRow();
  
  headers.forEach((header, index) => {
    const columnIndex = index + 1;

    if (header.includes('COST DIFF') || header.includes('VALUE EFFICIENCY PER ITEM') || header.includes('W/SALE COST') || header.includes('SINGLE RETAIL')) {
      // Apply currency format (£) to Cost Difference, Wholesale Cost, and Single Retail columns
      const range = sheet.getRange(2, columnIndex, numRows - 1);
      range.setNumberFormat('£#,##0.00');
    }

    if (header.includes('SINGLE MARGIN')) {
      // Apply percentage format to Single Margin columns
      const range = sheet.getRange(2, columnIndex, numRows - 1);
      range.setNumberFormat('0.00%');
    }
  });
}

function applyCenteringAndBoldFormatting(sheet) {
  const numRows = sheet.getLastRow();
  const numCols = sheet.getLastColumn();

  // Apply bold and centered formatting to headers
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // Apply center alignment to all data cells except descriptions
  const dataRange = sheet.getRange(2, 1, numRows - 1, numCols);
  dataRange.setHorizontalAlignment('center');

  // Center Description column only for headers, keep data left-aligned
  const descriptionColIndex = 3; // Assuming Description is the third column (C)
  const descriptionDataRange = sheet.getRange(2, descriptionColIndex, numRows - 1, 1);
  descriptionDataRange.setHorizontalAlignment('left');
}

function applyConditionalFormatting(sheet, headers) {
  const numRows = sheet.getLastRow();
  const numCols = headers.length;

  // Ensure there's enough data to apply formatting
  if (numRows <= 1 || numCols <= 1) {
    Logger.log("No sufficient data rows or columns found. Skipping conditional formatting.");
    return; // Exit if no valid data range
  }

  // Clear all existing conditional formatting rules from the sheet
  sheet.clearConditionalFormatRules();

  const rules = [];

  // Identify the latest promotion period for the cost difference
  let latestCostDiffIndex = -1;
  headers.forEach((header, index) => {
    if (header && header.includes('COST DIFF')) {
      latestCostDiffIndex = index; // Track the index of the latest Cost Difference
    }
  });

  // Loop through headers and apply conditional formatting based on known rules
  headers.forEach((header, index) => {
    const columnIndex = index + 1;

    if (header) { // Ensure header is not null

      // Apply conditional formatting for cost difference (green for decrease, red for increase, no color for no change)
      if (header.includes('COST DIFF') || header.includes('VALUE EFFICIENCY PER ITEM') && index > 11) {
        rules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenNumberLessThan(0)
            .setBackground("#00ff00") // Bright Green for negative difference (cost decrease)
            .setRanges([sheet.getRange(2, columnIndex, numRows - 1, 1)])
            .build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThan(0)
            .setBackground("#ff0000") // Bright Red for positive difference (cost increase)
            .setRanges([sheet.getRange(2, columnIndex, numRows - 1, 1)])
            .build()
        );
      }

      // Apply conditional formatting for wholesale cost and single retail (green for decrease, red for increase, no color for no change)
      if (header.includes('W/SALE COST') || header.includes('SINGLE RETAIL') && index > 7) {
        const previousColumnIndex = columnIndex - 7; // Previous period index in 1-based index
        if (previousColumnIndex > 7 && columnIndex <= headers.length) {
          rules.push(
            SpreadsheetApp.newConditionalFormatRule()
              .whenFormulaSatisfied(`=AND(NOT(ISBLANK(INDIRECT(ADDRESS(ROW(), ${previousColumnIndex})))), NOT(ISBLANK(INDIRECT(ADDRESS(ROW(), ${columnIndex})))), INDIRECT(ADDRESS(ROW(), ${columnIndex})) > INDIRECT(ADDRESS(ROW(), ${previousColumnIndex})))`)
              .setBackground("#ff0000") // Bright Red for increase (cost increase)
              .setRanges([sheet.getRange(2, columnIndex, numRows - 1, 1)])
              .build(),
            SpreadsheetApp.newConditionalFormatRule()
              .whenFormulaSatisfied(`=AND(NOT(ISBLANK(INDIRECT(ADDRESS(ROW(), ${previousColumnIndex})))), NOT(ISBLANK(INDIRECT(ADDRESS(ROW(), ${columnIndex})))), INDIRECT(ADDRESS(ROW(), ${columnIndex})) < INDIRECT(ADDRESS(ROW(), ${previousColumnIndex})))`)
              .setBackground("#00ff00") // Bright Green for decrease (cost decrease)
              .setRanges([sheet.getRange(2, columnIndex, numRows - 1, 1)])
              .build()
          );
        }
      }

      // Apply conditional formatting for single margin (green for increase, red for decrease, no color for no change)
      if (header.includes('SINGLE MARGIN') && index > 7) {
        const previousColumnIndex = columnIndex - 7; // Previous period index in 1-based index
        if (previousColumnIndex > 7 && columnIndex <= headers.length) {
          rules.push(
            SpreadsheetApp.newConditionalFormatRule()
              .whenFormulaSatisfied(`=AND(NOT(ISBLANK(INDIRECT(ADDRESS(ROW(), ${previousColumnIndex})))), NOT(ISBLANK(INDIRECT(ADDRESS(ROW(), ${columnIndex})))), INDIRECT(ADDRESS(ROW(), ${columnIndex})) > INDIRECT(ADDRESS(ROW(), ${previousColumnIndex})))`)
              .setBackground("#00ff00") // Bright Green for increase (margin increase)
              .setRanges([sheet.getRange(2, columnIndex, numRows - 1, 1)])
              .build(),
            SpreadsheetApp.newConditionalFormatRule()
              .whenFormulaSatisfied(`=AND(NOT(ISBLANK(INDIRECT(ADDRESS(ROW(), ${previousColumnIndex})))), NOT(ISBLANK(INDIRECT(ADDRESS(ROW(), ${columnIndex})))), INDIRECT(ADDRESS(ROW(), ${columnIndex})) < INDIRECT(ADDRESS(ROW(), ${previousColumnIndex})))`)
              .setBackground("#ff0000") // Bright Red for decrease (margin decrease)
              .setRanges([sheet.getRange(2, columnIndex, numRows - 1, 1)])
              .build()
          );
        }
      }

      // Apply conditional formatting for promo column (highlight if has value)
      if (header.includes('PROMO')) {
        rules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains('P')
            .setBackground("#00ff00") // Bright Green for promo available
            .setRanges([sheet.getRange(2, columnIndex, numRows - 1, 1)])
            .build()
        );
      }

      // Apply conditional formatting for status column
      if (header.includes('STATUS')) {
        rules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('NEW')
            .setBackground("#99CCFF") // Bright Blue for NEW
            .setRanges([sheet.getRange(2, columnIndex, numRows - 1, 1)])
            .build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('CWSL')
            .setBackground("#ff0000") // Bright Red for CWSL
            .setRanges([sheet.getRange(2, columnIndex, numRows - 1, 1)])
            .build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=AND(NOT(ISBLANK(INDIRECT(ADDRESS(ROW(), ${columnIndex})))), INDIRECT(ADDRESS(ROW(), ${columnIndex})) <> "NEW", INDIRECT(ADDRESS(ROW(), ${columnIndex})) <> "CWSL")`)
            .setBackground("#ff9933") // Bright Orange for other statuses
            .setRanges([sheet.getRange(2, columnIndex, numRows - 1, 1)])
            .build()
        );
      }



      // Conditional Formatting for Description Column (match latest cost difference)
      if (header.toLowerCase() === 'description') {
        // Find the latest cost difference column index
        const latestCostDiffIndex = headers.findIndex((h) => h.includes('COST DIFF')) + 1;

        if (latestCostDiffIndex > 0) {
          rules.push(
            SpreadsheetApp.newConditionalFormatRule()
              .whenFormulaSatisfied(`=INDIRECT(ADDRESS(ROW(), ${latestCostDiffIndex})) < 0`)
              .setBackground("#00ff00") // Bright Green for decrease in latest cost difference
              .setRanges([sheet.getRange(2, columnIndex, numRows - 1)])
              .build(),
            SpreadsheetApp.newConditionalFormatRule()
              .whenFormulaSatisfied(`=INDIRECT(ADDRESS(ROW(), ${latestCostDiffIndex})) > 0`)
              .setBackground("#ff0000") // Bright Red for increase in latest cost difference
              .setRanges([sheet.getRange(2, columnIndex, numRows - 1)])
              .build()
          );
        }
      }
    }
  });

  // Apply all the rules at once to avoid overwriting
  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
    Logger.log("Conditional formatting applied successfully.");
  } else {
    Logger.log("No valid rules to apply.");
  }
}



