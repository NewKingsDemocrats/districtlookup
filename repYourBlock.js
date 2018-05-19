/**
 * @OnlyCurrentDoc
 */

// Just under 5 minutes, in ms.
var MAX_RUNNING_TIME = 290000;
// 1 minute, in ms.
var REASONABLE_TIME_TO_WAIT = 60000;

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    // {name: 'Add missing AD and ED data', functionName: 'addMissingAdEd_'},
    {name: 'Add new users from import', functionName: 'movePeopleToAD'},
    {name: 'Check for changed users from import', functionName: 'checkForChanges'},
  ];
  spreadsheet.addMenu('RYB', menuItems);
}

// Global variable to store column names.
var columnNames = (columnNames || {});
    
// Global variable to store existing IDs in master sheet
var idsInMaster = (idsInMaster || []);

/** 
 * Helper function called from the NKD Menu that does all the work.
 *
 * @return null
 */
function addMissingAdEd_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('master with corrected ad&ed');
  // Look up which columns contain the Assembly District and Election District data and address.
  var adCol = getColByName_('Assembly District', sheet);
  if (adCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "Assembly District". Please add one.');
    return;
  }
  var edCol = getColByName_('Election District', sheet)
  if (edCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "Election District". Please add one.');
    return;
  }
  var address1Col = getColByName_('primary_address1', sheet);
  if (address1Col == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "primary_address1". Please check you are running this script on the correct sheet.');
    return;
  }
  var addressCityCol = getColByName_('primary_city', sheet);
  if (addressCityCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "primary_city". Please check you are running this script on the correct sheet.');
    return;
  }
  var addressStateCol = getColByName_('primary_state', sheet);
  if (addressStateCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "primary_state". Please check you are running this script on the correct sheet.');
    return;
  }
  var zipCol = getColByName_('primary_zip', sheet);
  if (zipCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "primary_zip". Please check you are running this script on the correct sheet.');
    return;
  }
  // Look through the sheet and find any missing data that needs to be looked up.
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
      if (values[i][addressCityCol] != '') {
        address = address + ", " + values[i][addressCityCol];
      } else {
        address = address + ", Brooklyn";
      }
    }
    // If there's a state value, add it to address1 for lookup.
    if (values[i][addressStateCol] != '') {
      address = address + ', ' + values[i][addressStateCol];
    }
    // If there's a zipcode value, add it to address1 for lookup.
    if (values[i][zipCol] != '') {
      address = address + ' ' + values[i][zipCol];
    }
    // Look up the sunlight API.
    Logger.log("Looking up address " + address);
    var sunlightData = getSunlightData_(address);
    // Update the appropriate cells with AD & ED data.
    // Update the appropriate cells with AD & ED data.
    var adCell = dataRange.getCell(i + 1, adCol + 1);
    var edCell = dataRange.getCell(i + 1, edCol + 1);
    Logger.log("Sunlight AD: " + sunlightData.ad);
    if (sunlightData.ad != undefined && sunlightData.ad != 'undefined' && sunlightData.ad) {
      adCell.setValue(sunlightData.ad);
      edCell.setValue(sunlightData.ed);
    } else {
      Logger.log("Undefined result for address " + address);
    }
  }
}

/** 
 * Calls the CC Sunlight API to look up the AD & ED for a row in a sheet.
 * TODO(low priority): There's a bunch of duplicate code between here and above. Probably just delete all of above.
 *
 * @param {Number} row The row to limit the AD/ED lookup to in the Master sheet.
 * @param {Sheet} sheet The sheet to be updated.
 * @param {Boolean} forceUpdate If true, will lookup and overwrite the AD & ED even if there's already a value in those columns
 * @return null
 */
function addMissingAdEdForRow_(i, sheet, forceUpdate) {
  // Look up which columns contain the Assembly District and Election District data and address.
  var adCol = getColByName_('Assembly District', sheet);
  if (adCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "Assembly District". Please add one.');
    return;
  }
  var edCol = getColByName_('Election District', sheet)
  if (edCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "Election District". Please add one.');
    return;
  }
  var address1Col = getColByName_('primary_address1', sheet);
  if (address1Col == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "primary_address1". Please check you are running this script on the correct sheet.');
    return;
  }
  var addressCityCol = getColByName_('primary_city', sheet);
  if (addressCityCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "primary_city". Please check you are running this script on the correct sheet.');
    return;
  }
  var addressStateCol = getColByName_('primary_state', sheet);
  if (addressStateCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "primary_state". Please check you are running this script on the correct sheet.');
    return;
  }
  var zipCol = getColByName_('primary_zip', sheet);
  if (zipCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "primary_zip". Please check you are running this script on the correct sheet.');
    return;
  }
  // Look through the sheet and find any missing data that needs to be looked up.
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  // If there's already a value for AD & ED, continue to next row, unless forceUpdate is true.
  if (values[i][adCol] && values[i][edCol] && !forceUpdate) {
    Logger.log("Skipped row " + i + " because it already had values");
    return;
  }
  // If there's no value for Address1, skip this row.
  if (values[i][address1Col] == '') {
    Logger.log("Skipped row " + i + " because there was no address in column " + address1Col);
    return;
  }
  // Put together the address.
  var address = values[i][address1Col];
  // If "Brooklyn" is not in the address, add it.
  if (address.indexOf(address) !== -1) {
    if (values[i][addressCityCol] != '') {
      address = address + ", " + values[i][addressCityCol];
    } else {
      address = address + ", Brooklyn";
    }
  }
  // If there's a state value, add it to address1 for lookup.
  if (values[i][addressStateCol] != '') {
    address = address + ', ' + values[i][addressStateCol];
  }
  // If there's a zipcode value, add it to address1 for lookup.
  if (values[i][zipCol] != '') {
    address = address + ' ' + values[i][zipCol];
  }
  // Look up the sunlight API.
  Logger.log("Looking up address " + address);
  var sunlightData = getSunlightData_(address);
  // Update the appropriate cells with AD & ED data.
  Logger.log("Sunlight AD: " + sunlightData.ad);
  if (sunlightData.ad != undefined && sunlightData.ad != 'undefined' && sunlightData.ad) {
    var adCell = dataRange.getCell(i + 1, adCol + 1);
    var edCell = dataRange.getCell(i + 1, edCol + 1);
    adCell.setValue(sunlightData.ad);
    edCell.setValue(sunlightData.ed);
  } else {
    Logger.log("Undefined result for address " + address);
  }
}

/**
 * An internal function to get the Column index with a particular name.
 * 
 * @param {String} name The column name being searched for in the first row.
 * @param {Sheet} sheet The sheet being checked, e.g. SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().
 * @return {Number} The column number.
 */
function getColByName_(name, sheet) {
  var headers;
  var sheetName = sheet.getName();
  if (columnNames == null || columnNames[sheetName] == null) {
    headers = sheet.getDataRange().getValues().shift();
    columnNames[sheetName] = headers;
  }  else {
    headers = columnNames[sheetName];
  }
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

function scratchHelper() {
  var importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('master export');
  var headers = importSheet.getDataRange().getValues().shift();
  Logger.log(headers);
  //var colindex = headers.indexOf(name);
}

// Helper function called by clicking on menu item.
function movePeopleToAD() {
  loopThroughImport_('new');
}

// Helper function called by clicking on menu item.
function checkForChanges() {
  loopThroughImport_('changed');
}

/**
 * Helper to loop through import sheet and either check for new users, or check for changes to existing users.
 * @param mode Either new or update.
 */
function loopThroughImport_(mode) {
  // Record the start time to later avoid timeout.
  var startTime = (new Date()).getTime();
  
  var importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('master export');
  if (importSheet == null) {
    SpreadsheetApp.getUi().alert('Spreadsheet is missing a sheet called "master export". Please add one.');
    return;
  }
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('master with corrected ad&ed'); 
  if (masterSheet == null) {
    SpreadsheetApp.getUi().alert('Spreadsheet is missing a sheet called "master with corrected ad&ed". Please add one.');
    return;
  }
  // Get columns that we care about
  var idColImport = getColByName_('nationbuilder_id', importSheet);
  if (idColImport == -1) {
    SpreadsheetApp.getUi().alert('Import sheet is missing a column called "nationbuilder_id". Please add one.');
    return;
  }
  var idColMaster = getColByName_('nationbuilder_id', masterSheet);
  if (idColMaster == -1) {
    SpreadsheetApp.getUi().alert('Master sheet is missing a column called "nationbuilder_id". Please add one.');
    return;
  }
  var values = importSheet.getDataRange().getValues();
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var startRow = Number(scriptProperties.getProperty('start_row'));
  Logger.log('start_row for this run is ' + scriptProperties.getProperty('start_row'));
  if (!startRow) {
    startRow = 1; // Skip the header row.
  }
  
  // Loop through each row in the spreadsheet (skipping the header row).
  for (var i = startRow; i < values.length; i++) {
    // Avoid timing out
    var currTime = (new Date()).getTime();
    if (currTime - startTime >= MAX_RUNNING_TIME) {
      Logger.log('Getting close to timeout, so setting up new trigger and ending this run at i: ' + i);
      scriptProperties.setProperty('start_row', i);
      Logger.log('start_row is now ' + scriptProperties.getProperty('start_row'));
      // Make sure trigger is for correct mode.
      var scriptToTrigger = 'movePeopleToAD';
      if (mode == 'changed' ) {
        scriptToTrigger = 'checkForChanges';
      }
      ScriptApp.newTrigger(scriptToTrigger)
               .timeBased()
               .at(new Date(currTime + REASONABLE_TIME_TO_WAIT))
               .create();
      return;
    }
    // Find each one in the master sheet, and if not there add it.
    var userRowInMasterSheet = maybeGetUsersRowInCorrectSheetFromCache_(values[i][idColImport], masterSheet, idColMaster); 
    if (userRowInMasterSheet && mode == 'changed') {
      Logger.log('user ' + values[i][idColImport] + ' is already in master sheet, on row ' + userRowInMasterSheet + '.');
      updateExistingUserOnMasterSheet_(values[i][idColImport], i, userRowInMasterSheet, importSheet, masterSheet);
    } else if (!userRowInMasterSheet && mode == 'new') {
      Logger.log('user ' + values[i][idColImport] + ' is new. Adding them.');
      // If adding, then also add AD/ED, then add to AD sheet, and update tracker col
      addNewUserToMasterSheet_(i, importSheet, masterSheet);
    }
  }
  // Reset the timeout tracker now we've reached the end.
  scriptProperties.deleteProperty("start_row");
  // Clean up triggers.
  deleteAllTriggers();
  Logger.log('Script completed successfully in ' + mode + ' mode.');
}

/**
 * Updates an existing user on the Master sheet. 
 * Check if any info has changed. If anything's changed:
 *   - if address has chaged, check AD/ED change, update those
 *    - if AD change, move to new sheet, update tracker
 *   - if any other key info changes, copy change over to AD sheet
 *   - update the AD sheet to note that the info has been corrected? TODO: Figure out.
 * 
 * @param {Number} id The user's NationBuilder ID.
 * @param {Number} userPositionImportSheet The position of the user in the importSheet. 
 * @param {Number} userPositionMasterSheet The position of the user in the masterSheet.
 * @param {Sheet} importSheet The sheet imported from NationBuilder.
 * @param {Sheet} masterSheet The master sheet for RYB, where we add AD/ED, and some tracking info.
 */
function updateExistingUserOnMasterSheet_(id, userPositionImportSheet, userPositionMasterSheet, importSheet, masterSheet) {
  // Check if any info has changed. If not, continue to next row
  // If anything's changed
  //   - if address has chaged, check AD/ED change, update those
  //    - if AD change, move to new sheet, update tracker
  //   - if any other key info changes, copy change over to AD sheet 
  //   - update the AD sheet to note that the info has been corrected? TODO: Figure out.
  
  // Get columns that we care about
  var address1ColImport = getColByName_('primary_address1', importSheet);
  if (address1ColImport == -1) {
    SpreadsheetApp.getUi().alert('importSheet is missing a column called "primary_address1". Please check you are running this script on the correct sheet.');
    return;
  }
  var address1ColMaster = getColByName_('primary_address1', masterSheet);
  if (address1ColMaster == -1) {
    SpreadsheetApp.getUi().alert('masterSheet is missing a column called "primary_address1". Please check you are running this script on the correct sheet.');
    return;
  }
  var address2ColImport = getColByName_('primary_address2', importSheet);
  if (address2ColImport == -1) {
    SpreadsheetApp.getUi().alert('importSheet is missing a column called "primary_address2". Please check you are running this script on the correct sheet.');
    return;
  }
  var address2ColMaster = getColByName_('primary_address2', masterSheet);
  if (address2ColMaster == -1) {
    SpreadsheetApp.getUi().alert('masterSheet is missing a column called "primary_address2". Please check you are running this script on the correct sheet.');
    return;
  }
  var addressCityColImport = getColByName_('primary_city', importSheet);
  if (addressCityColImport == -1) {
    SpreadsheetApp.getUi().alert('importSheet is missing a column called "primary_city". Please check you are running this script on the correct sheet.');
    return;
  }
  var addressCityColMaster = getColByName_('primary_city', masterSheet);
  if (addressCityColMaster == -1) {
    SpreadsheetApp.getUi().alert('masterSheet is missing a column called "primary_city". Please check you are running this script on the correct sheet.');
    return;
  }
  var addressStateColImport = getColByName_('primary_state', importSheet);
  if (addressStateColImport == -1) {
    SpreadsheetApp.getUi().alert('importSheet is missing a column called "primary_state". Please check you are running this script on the correct sheet.');
    return;
  }
  var addressStateColMaster = getColByName_('primary_state', masterSheet);
  if (addressStateColMaster == -1) {
    SpreadsheetApp.getUi().alert('masterSheet is missing a column called "primary_state". Please check you are running this script on the correct sheet.');
    return;
  }
  var zipColImport = getColByName_('primary_zip', importSheet);
  if (zipColImport == -1) {
    SpreadsheetApp.getUi().alert('importSheet is missing a column called "primary_zip". Please check you are running this script on the correct sheet.');
    return;
  }
  var zipColMaster = getColByName_('primary_zip', masterSheet);
  if (zipColMaster == -1) {
    SpreadsheetApp.getUi().alert('masterSheet is missing a column called "primary_zip". Please check you are running this script on the correct sheet.');
    return;
  }
  var adColMaster = getColByName_('Assembly District', masterSheet);
  if (adColMaster == -1) {
    SpreadsheetApp.getUi().alert('masterSheet is missing a column called "Assembly District". Please add one.');
    return;
  }
  var edColMaster = getColByName_('Election District', masterSheet);
  if (edColMaster == -1) {
    SpreadsheetApp.getUi().alert('masterSheet is missing a column called "Election District". Please add one.');
    return;
  }
    
  var importValues = importSheet.getDataRange().getValues();
  var masterValues = masterSheet.getDataRange().getValues();
  
  var previousAd = masterValues[userPositionMasterSheet][adColMaster];
  Logger.log('previous ad: ' + previousAd);
  var previousEd = masterValues[userPositionMasterSheet][edColMaster];
  if (!isInvalidAd_(previousAd)) {
    var adSheet = getOrCreateAdSheet_(masterValues[userPositionMasterSheet][adColMaster]);
    var idColAd = getColByName_('ID', adSheet);
    var userRowInAdSheet = maybeGetUsersRowInCorrectSheet_(id, adSheet, idColAd)
  }
  // Compare address values in import and master to see if it's changed.
  if (importValues[userPositionImportSheet][address1ColImport] != masterValues[userPositionMasterSheet][address1ColMaster] ||
      importValues[userPositionImportSheet][addressCityColImport] != masterValues[userPositionMasterSheet][addressCityColMaster] ||
      importValues[userPositionImportSheet][addressStateColImport] != masterValues[userPositionMasterSheet][addressStateColMaster] ||
      importValues[userPositionImportSheet][zipColImport] != masterValues[userPositionMasterSheet][zipColMaster]) {
    // Copy new values from import to master
    var address1Cell = masterSheet.getRange(userPositionMasterSheet + 1, address1ColMaster + 1);
    address1Cell.setValue(importValues[userPositionImportSheet][address1ColImport]);
    var address2Cell = masterSheet.getRange(userPositionMasterSheet + 1, address2ColMaster + 1);
    address2Cell.setValue(importValues[userPositionImportSheet][address2ColImport]);
    var addressCityCell = masterSheet.getRange(userPositionMasterSheet + 1, addressCityColMaster + 1);
    addressCityCell.setValue(importValues[userPositionImportSheet][addressCityColImport]);
    var addressStateCell = masterSheet.getRange(userPositionMasterSheet + 1, addressStateColMaster + 1);
    addressStateCell.setValue(importValues[userPositionImportSheet][addressStateColImport]);
    var zipCell = masterSheet.getRange(userPositionMasterSheet + 1, zipColMaster + 1);
    zipCell.setValue(importValues[userPositionImportSheet][zipColImport]);
    // Update AD & ED.
    addMissingAdEdForRow_(userPositionMasterSheet, masterSheet, true);
    // Re-get master values. 
    masterValues = masterSheet.getDataRange().getValues();
    var newAd = masterValues[userPositionMasterSheet][adColMaster];
    var newEd = masterValues[userPositionMasterSheet][edColMaster];

    // If AD changes, delete row in old AD sheet and add to new AD sheet, and update the tracker column
    if (previousAd != newAd) {
      Logger.log('AD changed');
      if (!isInvalidAd_(previousAd) && userRowInAdSheet) {
        // Delete row in old AD sheet.
        Logger.log('Deleting row ' + userRowInAdSheet + ' in sheet ' + adSheet.getName() + ' because AD changed due to address change.');
        adSheet.deleteRow(userRowInAdSheet + 1);
      } else {
        Logger.log('Skipped deleting user ' + id + ' from previous AD sheet because either the sheet is invalid, or they were not on that sheet.');
      }
      // Copy data to new AD sheet and update tracker column. 
      maybeCopyUserDataToAdSheet_(userPositionMasterSheet, newAd, masterSheet);
      // The above copies all data over so we can return early.
      return;
    } else if (!isInvalidAd_(previousAd)) {
      // The address changed, but not the AD, so copy all address info over to existing row in AD sheet, including ED
      Logger.log('Address changed, but not the AD.');
      var addressColAd = getColByName_('Address', adSheet);
      var addressCell = adSheet.getRange(userRowInAdSheet + 1, addressColAd + 1);
      addressCell.setValue(masterValues[userPositionMasterSheet][address1ColMaster]);
      var edColAd = getColByName_('ED', adSheet);
      var edCell = adSheet.getRange(userRowInAdSheet + 1, edColAd + 1);
      edCell.setValue(newEd);
    }
  }
  // Compare first name, and update if necessary.
  var firstColImport = getColByName_('first_name', importSheet);
  var firstColMaster = getColByName_('first_name', masterSheet);
  if (importValues[userPositionImportSheet][firstColImport] != masterValues[userPositionMasterSheet][firstColMaster]) {
    // Update Master sheet.
    var firstCellMaster = masterSheet.getRange(userPositionMasterSheet + 1, firstColMaster + 1);
    firstCellMaster.setValue(importValues[userPositionImportSheet][firstColImport]);
    // Update AD sheet.
    if (!isInvalidAd_(previousAd)) {
      var firstColAd = getColByName_('First name', adSheet);
      var firstCellAd = adSheet.getRange(userRowInAdSheet + 1, firstColAd + 1);
      firstCellAd.setValue(importValues[userPositionImportSheet][firstColImport]);
    }
  }
  // Compare last name, and update if necessary.
  var lastColImport = getColByName_('last_name', importSheet);
  var lastColMaster = getColByName_('last_name', masterSheet);
  if (importValues[userPositionImportSheet][lastColImport] != masterValues[userPositionMasterSheet][lastColMaster]) {
    // Update Master sheet.
    var lastCellMaster = masterSheet.getRange(userPositionMasterSheet + 1, lastColMaster + 1);
    lastCellMaster.setValue(importValues[userPositionImportSheet][lastColImport]);
    // Update AD sheet.
    if (!isInvalidAd_(previousAd)) {
      var lastColAd = getColByName_('Last name', adSheet);
      var lastCellAd = adSheet.getRange(userRowInAdSheet + 1, lastColAd + 1);
      lastCellAd.setValue(importValues[userPositionImportSheet][lastColImport]);
    }
  }  
  // Compare phone number, and update if necessary.
  var phoneColImport = getColByName_('phone_number', importSheet);
  var phoneColMaster = getColByName_('phone_number', masterSheet);
  var newPhone = false;
  if (importValues[userPositionImportSheet][phoneColImport] != masterValues[userPositionMasterSheet][phoneColMaster]) {
    newPhone = true;
    // Update Master sheet.
    var phoneCellMaster = masterSheet.getRange(userPositionMasterSheet + 1, phoneColMaster + 1);
    phoneCellMaster.setValue(importValues[userPositionImportSheet][phoneColImport]);
  }
  // Compare mobile number, and update if necessary
  var mobileColImport = getColByName_('mobile_number', importSheet);
  var mobileColMaster = getColByName_('mobile_number', masterSheet);
  if (importValues[userPositionImportSheet][mobileColImport] != masterValues[userPositionMasterSheet][mobileColMaster]) {
    newPhone = true;
    // Update Master sheet.
    var mobileCellMaster = masterSheet.getRange(userPositionMasterSheet + 1, mobileColMaster + 1);
    mobileCellMaster.setValue(importValues[userPositionImportSheet][mobileColImport]);
  }
  // Some special logic for updating phone on AD sheet.
  if (newPhone && !isInvalidAd_(previousAd)) {
    var phoneValue = importValues[userPositionImportSheet][phoneColImport]
    if (phoneValue == '') {
      phoneValue = importValues[userPositionImportSheet][mobileColImport]
    }
    var phoneColAd = getColByName_('Phone', adSheet);
    var phoneCellAd = adSheet.getRange(userRowInAdSheet + 1, phoneColAd + 1);
    phoneCellAd.setValue(phoneValue);
  }
  // Compare email, and update if necessary.
  var emailColImport = getColByName_('email', importSheet);
  var emailColMaster = getColByName_('email', masterSheet);
  if (importValues[userPositionImportSheet][emailColImport] != masterValues[userPositionMasterSheet][emailColMaster]) {
    // Update Master sheet.
    var emailCellMaster = masterSheet.getRange(userPositionMasterSheet + 1, emailColMaster + 1);
    emailCellMaster.setValue(importValues[userPositionImportSheet][emailColImport]);
    // Update AD sheet.
    if (!isInvalidAd_(previousAd)) {
      var emailColAd = getColByName_('Email', adSheet);
      var emailCellAd = adSheet.getRange(userRowInAdSheet + 1, emailColAd + 1);
      emailCellAd.setValue(importValues[userPositionImportSheet][emailColImport]);
    }
  }
}

/**
 * Adds a new user to the Master sheet, updates the sheet with their AD/ED, and then copies them to their
 * AD-specific sheet. 
 * 
 * @param {Number} userPositionImportSheet The position of the user in the importSheet. 
 * @param {Sheet} importSheet The sheet imported from NationBuilder.
 * @param {Sheet} masterSheet The master sheet for RYB, where we add AD/ED, and some tracking info.
 */
function addNewUserToMasterSheet_(userPositionImportSheet, importSheet, masterSheet) {
  // First copy user to masterSheet
  var newRow = masterSheet.getLastRow(); // maybe store as global or as scriptproperty for speed?
  var userRow = importSheet.getRange(userPositionImportSheet + 1, 1, 1, importSheet.getLastColumn()); // maybe replace getLastColumn with global or scriptproperty?
  userRow.copyTo(masterSheet.getRange(newRow + 1, 1, 1, importSheet.getLastColumn()));  
  
  // Add AD/ED
  addMissingAdEdForRow_(newRow, masterSheet, false);
  var values = masterSheet.getDataRange().getValues(); // is there a faster way to do this? I just need the AD.
  var adCol = getColByName_('Assembly District', masterSheet);
  if (adCol == -1) {
    SpreadsheetApp.getUi().alert('Sheet is missing a column called "Assembly District". Please add one.');
    return;
  }

  maybeCopyUserDataToAdSheet_(newRow, values[newRow][adCol], masterSheet);
}

/** 
 * Copy the user's info from the master sheet to the appropriate AD sheet if the AD
 * is valid, then updates the tracker column in the master sheet if successful.
 *
 * @param {Number} userPosition The user's position on the Master sheet.
 * @param {Number} ad The user's Assembly District.
 * @param {Sheet} masterSheet The Master sheet to copy the user from.
 */
function maybeCopyUserDataToAdSheet_(userPosition, ad, masterSheet) {
  if (isInvalidAd_(ad)) {
    Logger.log("Row " + newRow + " has invalid AD: " + ad);
    return;
  }
  // The user hasn't been added to an AD sheet yet, so do that now.
  var adSheet = getOrCreateAdSheet_(ad);
  
  // Get all the column names.
  var idColMaster = getColByName_('nationbuilder_id', masterSheet);
  var idColAd = getColByName_('ID', adSheet);
  var lastColMaster = getColByName_('last_name', masterSheet);
  var lastColAd = getColByName_('Last name', adSheet);
  var firstColMaster = getColByName_('first_name', masterSheet);
  var firstColAd = getColByName_('First name', adSheet);
  var phoneColMaster = getColByName_('phone_number', masterSheet);	
  var phoneMobColMaster = getColByName_('mobile_number', masterSheet);
  var phoneColAd = getColByName_('Phone', adSheet);
  var emailColMaster = getColByName_('email', masterSheet);
  var emailColAd = getColByName_('Email', adSheet);
  var addressColMaster = getColByName_('primary_address1', masterSheet);
  var addressColAd = getColByName_('Address', adSheet);
  var adColMaster = getColByName_('Assembly District', masterSheet);
  var adColAd = getColByName_('AD', adSheet);
  var edColMaster = getColByName_('Election District', masterSheet);
  var edColAd = getColByName_('ED', adSheet);
  // Check that no columns are missing.
  if (idColMaster == -1 || idColAd == -1 || lastColMaster == -1 || lastColAd == -1 || firstColMaster == -1 || 
      firstColAd == -1 || phoneColMaster == -1 || phoneMobColMaster == -1 || phoneColAd == -1 || emailColMaster == -1 || 
      emailColAd == -1 || addressColMaster == -1 || addressColAd == -1 || adColMaster == -1 || adColAd == -1 || 
      edColMaster == -1 || edColAd == -1) {
    SpreadsheetApp.getUi().alert('The column headers are wrong, so user data can not be copied.');
    return;
  }
  // Find the first empty row on AD sheet, and duplicate it onto the next line to copy over the formulas and formatting.
  var newRow = adSheet.getLastRow(); 
  var templateRow = adSheet.getRange(newRow + 1, 1, 1, adSheet.getLastColumn());
  templateRow.copyTo(adSheet.getRange(newRow + 2, 1, 1, adSheet.getLastColumn()));
  newRow++

  // Now copy over all the values from Master to AD sheet.
  var masterValues = masterSheet.getDataRange().getValues();
  var idCell = adSheet.getRange(newRow, idColAd + 1);
  idCell.setValue(masterValues[userPosition][idColMaster]);
  var lastNameCell = adSheet.getRange(newRow, lastColAd + 1);
  lastNameCell.setValue(masterValues[userPosition][lastColMaster]);
  var firstCell = adSheet.getRange(newRow, firstColAd + 1);
  firstCell.setValue(masterValues[userPosition][firstColMaster]);
  // Some special case logic for phone numbers, as there are two columns in master.
  var phoneCell = adSheet.getRange(newRow, phoneColAd + 1);
  var phoneValue = masterValues[userPosition][phoneColMaster]
  if (phoneValue == '') {
    phoneValue = masterValues[userPosition][phoneMobColMaster]
  }
  phoneCell.setValue(phoneValue);
  var emailCell = adSheet.getRange(newRow, emailColAd + 1);
  emailCell.setValue(masterValues[userPosition][emailColMaster]);
  var addressCell = adSheet.getRange(newRow, addressColAd + 1);
  addressCell.setValue(masterValues[userPosition][addressColMaster]);
  var adCell = adSheet.getRange(newRow, adColAd + 1);
  adCell.setValue(masterValues[userPosition][adColMaster]);
  var edCell = adSheet.getRange(newRow, edColAd + 1);
  edCell.setValue(masterValues[userPosition][edColMaster]);
  
  // Update the master sheet to note that they've been added.
  var adTrackerCol = getColByName_('Copied to AD Sheet', masterSheet);
  if (adTrackerCol == -1) {
    SpreadsheetApp.getUi().alert('masterSheet is missing a column called "Copied to AD Sheet". Please add one.');
    return;
  }
  var trackerCell = masterSheet.getRange(userPosition + 1, adTrackerCol + 1);
  trackerCell.setValue(adSheet.getName());
}

/**
 * Check if a user ID is already in a particular sheet. 
 *
 * @param {Number} id The user's ID from the NationBuilder export.
 * @param {Sheet} sheet The sheet to check.
 * @param {Number} idCol The column in `sheet` containing the NationBuilder ID.
 * @return {?Number} The row that the user was found, or null if not found.
 */
function maybeGetUsersRowInCorrectSheet_(id, sheet, idCol) {
  var values = sheet.getDataRange().getValues();
  // Loop through each row in the spreadsheet (skipping the header row).
  for (var i = 1; i < values.length; i++) {
    if (values[i][idCol] == id) {
      return i;
    }
  }
  return;
}


/**
 * Check if a user ID is already in a particular sheet. 
 *
 * @param {Number} id The user's ID from the NationBuilder export.
 * @param {Sheet} sheet The sheet to check.
 * @param {Number} idCol The column in `sheet` containing the NationBuilder ID.
 * @return {?Number} The row that the user was found, or null if not found.
 */
function maybeGetUsersRowInCorrectSheetFromCache_(id, sheet, idCol) {
  if (idsInMaster.length == 0) {
    Logger.log('Cache miss for getting users in row.');
    var lastRow = sheet.getLastRow();
    var values = sheet.getRange('A1:A' + lastRow).getValues(); 
    for (var i = 0; i < (lastRow); i++) {
      idsInMaster.push(values[i][0]);
    }
  }
  var pos = idsInMaster.indexOf(Number(id));
  if (pos == -1) {
    return;
  }
  return pos;
}
  
/**
 * A helper to get the sheet for an AD, or create and return a new one if necessary.
 *
 * @param {Number} ad The Assembly District Number.
 * @return {Sheet}
 */
function getOrCreateAdSheet_(ad, masterSheet, templateSheet) {
  Logger.log("Trying to find ad sheet " + ad);
  if (isInvalidAd_(ad)) { return; };
  var adSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ad); 
  if (adSheet == null) {
    // Set the active sheet to the blank template.
    var templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('template AD sheet'); 
    if (templateSheet == null) {
      SpreadsheetApp.getUi().alert('Spreadsheet is missing the template AD sheet.');
      return;
    }
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(templateSheet);
    // Duplicate the active sheet. 
    adSheet = SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
    // Rename the active sheet using ad.
    SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet(ad);    
    // Set focus back on master sheet.
    var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('master with corrected ad&ed'); 
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(masterSheet);
    // Re-sort the sheet tabs.
    sortSheets();
  }
  return adSheet;
}

/**
 * A helper to check that an AD is valid.
 * This returns false for entries with no AD, or with an AD that's not in King's County.
 *
 * @param {Number} ad The Assembly District Number.
 * @return {Boolean}
 */
function isInvalidAd_(ad) {
  if (ad == 'undefined') { return false; }
  // 41-64 are Kings County https://en.wikipedia.org/wiki/New_York_State_Assembly
  if (ad < 41 || ad > 64) {
    return true;
  }
  return false;
}

/**
 * A helper for one-time use to sort sheet tabs alphabetically.
 */
function sortSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNameArray = [];
  var sheets = ss.getSheets();
   
  for (var i = 0; i < sheets.length; i++) {
    sheetNameArray.push(sheets[i].getName());
  }
  
  sheetNameArray.sort();
    
  for( var j = 0; j < sheets.length; j++ ) {
    ss.setActiveSheet(ss.getSheetByName(sheetNameArray[j]));
    ss.moveActiveSheet(j + 1);
  }
}

/**
 * A DANGEROUS helper for one-time clean up, which deletes the AD sheets.
 */
function deleteAdSheetsDANGEROUS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNameArray = [];
  var sheets = ss.getSheets();
   
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().length < 3) {
      // ss.deleteSheet(sheets[i]);
      Logger.log("deleting " + sheets[i].getName());
    }
  }
}

function resetTimeoutTracker() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("start_row", 1);
}
function setTimeoutTrackerToNumber() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("start_row", 300);
}
function getTimeoutTracker() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var startRowa = scriptProperties.getProperty('start_row');
  Logger.log(startRowa);
  var startRowb = Number(scriptProperties.getProperty('start_row'));
  Logger.log(startRowb);
  if (!startRowa) {
    Logger.log("a is null");
  }
  if (!startRowb) {
    Logger.log("b is null");
  }
}
function deleteAllTriggers() {
  // Deletes all triggers in the current project.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

}

function setTestToNumber() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("test", 1);
}  
  
function testTrigger() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var startRowa = scriptProperties.getProperty('test');
  Logger.log(startRowa);
  var startRowb = Number(scriptProperties.getProperty('test'));
  Logger.log(startRowb);
  ++startRowb;
  scriptProperties.setProperty('test', startRowb); 
  Logger.log('new value ' + scriptProperties.getProperty('test'));
  var currTime = (new Date()).getTime();
  ScriptApp.newTrigger("testTrigger")
   .timeBased()
   .at(new Date(currTime + REASONABLE_TIME_TO_WAIT))
   .create();    
}
