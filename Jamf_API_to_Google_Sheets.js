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


function uiChanges() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Jamf API Collection')
      .addItem('Class of 2024','classOf2024')
      .addItem('Class of 2025','classOf2025')
      .addItem('Class of 2026','classOf2026')
      .addItem('Class of 2027','classOf2027')
      .addItem('Class of 2028','classOf2028')
      .addItem('Class of 2029','classOf2029')
      .addItem('Class of 2030','classOf2030')
      .addItem('Class of 2031','classOf2031')
      .addItem('Class of 2032','classOf2032')
      .addItem('Class of 2033','classOf2033')
      .addItem('PreK', 'preK')
      .addItem('Kindergarten', 'kinderGarten')
      //.deleteItem('Class of 2024')
      //.addItem('Class of 2034','classOf2034')
      .addToUi();
}


function changeFontAndSize() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet(); // Get the active sheet
  var range = sheet.getRange("A1:G200"); // Change this to your desired range
  var fontFamily = "Calibri";
  var fontSize = 12;

  range.setFontFamily(fontFamily).setFontSize(fontSize);
}


function classOf2024() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "74"; // Replace with the ID of your advanced computer search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([fullName, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}

function classOf2025() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "144"; // Replace with the ID of your advanced computer search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([fullName, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}

function classOf2026() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "155"; // Replace with the ID of your advanced computer search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([fullName, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}

function classOf2027() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "179"; // Replace with the ID of your advanced computer search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([fullName, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}

function classOf2028() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "199"; // Replace with the ID of your advanced computer search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([fullName, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}

function classOf2029() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "115"; // Replace with the ID of your advanced computer search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([fullName, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}

function classOf2030() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "123"; // Replace with the ID of your advanced computer search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([fullName, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}

function classOf2031() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "189"; // Replace with the ID of your advanced computer search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([fullName, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}

function classOf2032() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "216"; // Replace with the ID of your advanced computer search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([fullName, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}

function classOf2033() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "244"; // Replace with the ID of your advanced computer search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Computer Name", "Serial Number", "Model", "Operating System", "Last Check In", "Department"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
  computerElements.forEach(function(computer) {
    //changes the font and the size before the data populates into the google
    changeFontAndSize();

    var fullName = computer.getChildText("Full_Name");
    var computerName = computer.getChildText("Computer_Name");
    var serialNumber = computer.getChildText("Serial_Number");
    var model = computer.getChildText("Model");
    var operatingSystem = computer.getChildText("Operating_System");
    var lastCheckIn = computer.getChildText("Last_Check_in");
    var department = computer.getChildText("Department");

    rowData.push([fullName, computerName, serialNumber, model, operatingSystem, lastCheckIn, department]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}


function preK() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "71"; // Replace with the ID of your advanced mobile search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Display Name", "Serial Number", "Model", "iOS Version", "Last Inventory Update", "Username"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
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

    rowData.push([fullName, displayName, serialNumber, model, iOSVersion, lastInventoryUpdate, userName]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}


function kinderGarten() {
  var username = " "; // Replace with your Jamf username
  var password = " "; // Replace with your Jamf password
  var jamfURL = " "; // Replace with your Jamf instance URL
  var advancedSearchID = "106"; // Replace with the ID of your advanced mobile search

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
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
    sheet.getRange(2, 1, lastRowWithData - 1, 7).clear();
  }

  // Write headers to the sheet
  var headers = ["Full Name", "Display Name", "Serial Number", "Model", "iOS Version", "Last Inventory Update", "Username"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Iterate through the XML data and write it to the sheet starting below the headers
  var rowData = [];
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

    rowData.push([fullName, displayName, serialNumber, model, iOSVersion, lastInventoryUpdate, userName]);
  });

  // Write the data starting below the headers
  sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

  Logger.log("API Response parsed and data written to Google Sheets.");
}