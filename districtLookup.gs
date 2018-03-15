/**
 * @OnlyCurrentDoc
 */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Add missing AD and ED data', functionName: 'addMissingAdEd_'}
  ];
  spreadsheet.addMenu('NKD', menuItems);
}

/** 
 * Helper function called from the NKD Menu that does all the work.
 */
function addMissingAdEd_() {
  // Look up which columns contain the Assembly District and Election District data and address.
  var adCol = getColByName_('Assembly District');
  if (adCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "Assembly District". Please add one.');
    return;
  }
  var edCol = getColByName_('Election District')
  if (edCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "Election District". Please add one.');
    return;
  }
  var address1Col = getColByName_('Address 1');
  if (address1Col == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "Address 1". Please add one.');
    return;
  }
  var zipCol = getColByName_('Zip Code');
  // Look through the sheet and find any missing data that needs to be looked up.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  // Loop through each row in the spreadsheet (skipping the header row).
  for (var i = 1; i < values.length; i++) {
    // If there's already a value for AD & ED, continue to next row.
    if (values[i][adCol] && values[i][edCol]) {
      Logger.log("Skipped row " + i + " because it already had values");
      continue;
    }
    // If there's no value for Address1, skip this row.
    if (values[i][address1Col] == '') {
      Logger.log("Skipped row " + i + " because there was no address in column " + address1Col);
      continue;
    }
    // Put together the address.
    var address = values[i][address1Col];
    // If "Brooklyn" is not in the address, add it.
    if (address.indexOf(address) !== -1) {
      address = address + ", Brooklyn";
    }
    // If there's a zipcode value, add it to address1 for lookup.
    if (values[i][zipCol] != '') {
      address = address + ', NY ' + values[i][zipCol];
    }
    // Look up the sunlight API.
    Logger.log("Looking up addess " + address);
    var sunlightData = getSunlightData_(address);
    // Update the appropriate cells with AD & ED data.
    // Update the appropriate cells with AD & ED data.
    var adCell = dataRange.getCell(i + 1, adCol + 1);
    var edCell = dataRange.getCell(i + 1, edCol + 1);
    adCell.setValue(sunlightData.ad);
    edCell.setValue(sunlightData.ed);
  }
}

/**
 * An internal function to get the Column index with a particular name.
 * 
 * @param {String} name The column name being searched for in the first row.
 * @return {Number} The column number.
 */
function getColByName_(name) {
  var headers = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues().shift();
  var colindex = headers.indexOf(name);
  return colindex;
}

/**
 * A helper function to get all data from the Sunlight API for an address.
 * 
 * @param {String} address The street address.
 * @return {Object} The json response from the API.
 */
function getSunlightData_(address) {
  // The API to look up.
  var url = 'https://ccsunlight.org/api/v1/address/' + encodeURIComponent(address);
  // Fetch the full response.
  var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data;
}
