/*
This is an inventory system for Jamf's API's advanced search to Google Sheets.
Items that need to be changed every year:
  First you need to delete the old year from the Ui update in google sheets from the function uiChanges.
    uncomment out .deleteItem('') then add the new year to .addItem('Class of 2034', 'classOf2034')
  Now add a new function such as classOf2034 and copy the same code from an old function into the new one
    then you will need to change the advanced search id in var advancedSearchID = ""
    the search id is from the url from the advanced search when in jamf Ex:  /advancedComputerSearches.html?id=74&o=r the id number is located after the ?id=
  These will be the steps every year unless the api changes.
*/

function getBearerToken() {
  var username = ""; // Replace with your Jamf username
  var password = ""; // Replace with your Jamf password
  var jamfURL = ""; // Replace with your Jamf instance URL

  var authTokenResponse = UrlFetchApp.fetch(jamfURL + "/api/v1/auth/token", {
    "method": "post",
    "headers": {
      "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password),
      "Accept": "application/json"
    },
    "muteHttpExceptions": true
  });

  var authTokenText = authTokenResponse.getContentText();
  var authToken = JSON.parse(authTokenText).token;

  Logger.log("Bearer Token: " + authToken); // Log the bearer token

  return authToken;
}

function invalidateBearerToken(token) {
  var jamfURL = ""; // Replace with your Jamf instance URL

  UrlFetchApp.fetch(jamfURL + "/api/v1/auth/invalidate-token", {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + token
    }
  });
}


function uiChanges() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Jamf API Collection');

  // Create sub-menu for class years
  var classYearsMenu = ui.createMenu('Class Years');
  classYearsMenu.addItem('Class of 2024', 'classOf2024');
  classYearsMenu.addItem('Class of 2025', 'classOf2025');
  classYearsMenu.addItem('Class of 2026', 'classOf2026');
  classYearsMenu.addItem('Class of 2027', 'classOf2027');
  classYearsMenu.addItem('Class of 2028', 'classOf2028');
  classYearsMenu.addItem('Class of 2029', 'classOf2029');
  classYearsMenu.addItem('Class of 2030', 'classOf2030');
  classYearsMenu.addItem('Class of 2031', 'classOf2031');
  classYearsMenu.addItem('Class of 2032', 'classOf2032');
  classYearsMenu.addItem('Class of 2033', 'classOf2033');
  menu.addSubMenu(classYearsMenu);

  // Add other menu items iPads
  var iPadgrades = ui.createMenu('iPad Grades');
  iPadgrades.addItem('PreK', 'preK');
  iPadgrades.addItem('Kindergarten', 'kinderGarten');
  iPadgrades.addItem('1stGrade', 'oneGrade');
  iPadgrades.addItem('2ndGrade', 'twoGrade');
  menu.addSubMenu(iPadgrades);


  menu.addSeparator();
  menu.addItem('Pull Class data through searchid iPads', 'dataPullSearchIdIpads');
  menu.addItem('Pull Class data through searchid Computers', 'dataPullSearchIdComputers');
  // Add the main menu to the UI
  menu.addToUi();
}



function changeFontAndSize() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet(); // Get the active sheet
  var range = sheet.getRange("A1:K500"); // Change this to your desired range
  var fontFamily = "Calibri";
  var fontSize = 12;
  var numberFormat = ('MM/dd/yyyy');

  range.setFontFamily(fontFamily).setFontSize(fontSize).setNumberFormat(numberFormat);
}

function getDataPopulateDataComputers(advancedSearchID) {
  var jamfURL = ""; // Replace with your Jamf instance URL
  var authToken = getBearerToken(); // Get the bearer token

  var headers = {
    "Authorization": "Bearer " + authToken
  };
  
  var options = {
    "method": "get",
    "headers": headers
  };

  var response = UrlFetchApp.fetch(jamfURL + "/JSSResource/advancedcomputersearches/id/" + advancedSearchID, options);
  var data = response.getContentText();

  // Parse the XML response
  var xmlDoc = XmlService.parse(data);
  var computerElements = xmlDoc.getRootElement().getChild("computers").getChildren("computer");

  // Get the active sheet in the currently open spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Find the last row with data in column A
  var lastRowWithData = sheet.getLastRow();

  // If there is data below the headers, clear that specific range
  if (lastRowWithData > 1) {
    sheet.getRange(2, 1, lastRowWithData - 1, 8).clear();
  }

  // Write headers to the sheet
  var headers = ["Date Checked", "Full Name", "Email Address", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department",];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  var currentDate = new Date(); // Get the current date
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var emailAddress = computer.getChildText("Email_Address")
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([currentDate, fullName, emailAddress, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
    
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  // Sort all the data to last checkin
  sheet.getRange(2, 1, rowData.length, rowData[0].length).sort(8); // 7 is the column index for "Last Check In"

  // Sort all the data to align the the left side
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setHorizontalAlignment("left").setVerticalAlignment("middle");

  Logger.log("API Response parsed and data written to Google Sheets.");

  invalidateBearerToken(authToken);
}

function classOf2024() {
  getDataPopulateDataComputers("74");
}

function classOf2025() {
  getDataPopulateDataComputers("144");
}

function classOf2026() {
  getDataPopulateDataComputers("155");
}

function classOf2027() {
  getDataPopulateDataComputers("179");
}

function classOf2028() {
  getDataPopulateDataComputers("199");
}

function classOf2029() {
  getDataPopulateDataComputers("115");
}

function classOf2030() {
  getDataPopulateDataComputers("123");
}

function classOf2031() {
  getDataPopulateDataComputers("189");
}

function classOf2032() {
  getDataPopulateDataComputers("216");
}

function classOf2033() {
  getDataPopulateDataComputers("244");
}


function getDataPopulateDataiPads(advancedSearchID) {
  var jamfURL = ""; // Replace with your Jamf instance URL
  var authToken = getBearerToken(); // Get the bearer token

  var headers = {
    "Authorization": "Bearer " + authToken
  };
  
  var options = {
    "method": "get",
    "headers": headers
  };
  var response = UrlFetchApp.fetch(jamfURL + "/JSSResource/advancedmobiledevicesearches/id/" + advancedSearchID, options);
  var data = response.getContentText();

  // Parse the XML response
  var xmlDoc = XmlService.parse(data);
  var mobileElements = xmlDoc.getRootElement().getChild("mobile_devices").getChildren("mobile_device");

  // Get the active sheet in the currently open spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Find the last row with data in column A
  var lastRowWithData = sheet.getLastRow();

  // If there is data below the headers, clear that specific range
  if (lastRowWithData > 1) {
    sheet.getRange(2, 1, lastRowWithData - 1, 8).clear();
  }

  // Write headers to the sheet
  var headers = ["Date Checked", "Full Name", "Display Name", "Serial Number", "Model", "iOS Version", "Last Inventory Update", "Username"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  var currentdate = new Date(); // Gets the current Date
  mobileElements.forEach(function(mobile) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = mobile.getChildText("Full_Name");
    var displayName = mobile.getChildText("Display_Name");
    var serialNumber = mobile.getChildText("Serial_Number");
    var model = mobile.getChildText("Model");
    var iOSVersion = mobile.getChildText("iOS_Version");
    var lastInventoryUpdate = mobile.getChildText("Last_Inventory_Update");
    var userName = mobile.getChildText("Username");

    rowData.push([currentdate, fullName, displayName, serialNumber, model, iOSVersion, lastInventoryUpdate, userName]);
  });
 
  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  // Sort all the data to last checkin
  sheet.getRange(2, 1, rowData.length, rowData[0].length).sort(7); // 7 is the column index for "Last Check In"

  // Sort all the data to align the the left side
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setHorizontalAlignment("left").setVerticalAlignment("middle");

  Logger.log("API Response parsed and data written to Google Sheets.");

  invalidateBearerToken(authToken);
}

function preK() {
  getDataPopulateDataiPads("71");
}

function kinderGarten() {
  getDataPopulateDataiPads("106");
}


function oneGrade() {
  getDataPopulateDataiPads("187");
}


function twoGrade() {
  getDataPopulateDataiPads("105");
}


function dataPullSearchIdIpads() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  var advancedSearchID = sheet.getRange("A1").getValue(); // Read the advanced search ID from A1
  getDataPopulateDataiPads(advancedSearchID);
}

function dataPullSearchIdComputers() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  var advancedSearchID = sheet.getRange("A1").getValue(); // Read the advanced search ID from A1
  getDataPopulateDataComputers(advancedSearchID);
}