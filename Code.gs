/*
*
* NOKORI·WARE Table of Contents Generator for Google Sheets
* December 2018
*
* This add-on will help work around the current lack of a proper Table of Contents system in Google Sheets.
*
*/

/*
* This is the core function, use this to begin running the add-on.
*/
function NOKORIWARETableOfContentsGenerator() {
  var sideBar = HtmlService
    .createHtmlOutputFromFile('Sidebar')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Table of Contents Generator');
    
  SpreadsheetApp.getUi().showSidebar(sideBar);
}

/*
* Get all sheets from this spreadsheet.
*/
function getSheets() {
  var spreadSheet = SpreadsheetApp.getActive();
  var sheets = spreadSheet.getSheets();
  
  var sheetTitles = {};
  
  for (var i = 0; i < sheets.length; i++) {
    sheetTitles[i] = sheets[i].getName();
  }
  
  return sheetTitles;
}

/*
* Will generate a list of section names for a table of contents from the A column of the given sheet.
*/
function getTableOfContentsSections(sheet) {
  var sections = [];
  
  var range = sheet.getRange(1, 1, sheet.getLastRow(), 1);
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) {
      sections.push([values[i][0], i]);
      Logger.log("Section: " + values[i][0] + " " + i);
    }
  }
  
  return sections;
}

/*
* Gets the section names of the default sheet in the Spreadsheet
*/
function getTableOfContentsSectionNames(index) {
  var spreadSheet = SpreadsheetApp.getActive();
  var sourceSheet = spreadSheet.getSheets()[index];
  var tocSheet = spreadSheet.getSheetByName(sourceSheet.getRange("A1").getValue());
  
  var names = [];
  
  if (tocSheet != null){
    var sections = getTableOfContentsSections(tocSheet);
    
    for (var i = 0; i < sections.length; i++) {
      names[i] = sections[i][0];
    }
  }
  
  return names;
}

/*
* Generates a table of contents sheet from the given source sheet index
*/
function generateTableOfContents(sourceSheetIndex) {
  var spreadSheet = SpreadsheetApp.getActive();
  var sheets = spreadSheet.getSheets();
  
  //debug("Call " + sourceSheetIndex);
  
  var sourceSheet = sheets[sourceSheetIndex];
  
  //The table of contents sheet is designated at A1 of the source sheet by name.
  var tocName = sourceSheet.getRange(1, 1).getValue();
  var tocSheet = spreadSheet.getSheetByName(tocName);
  
  if (tocSheet == null) {
    tocSheet = spreadSheet.insertSheet(tocName);
  }
  
  tocSheet.activate();
  SpreadsheetApp.flush();
  
  if (tocSheet == sourceSheet) {
    showMessage("Error", "Source sheet and Table of Contents sheet both match. Make sure they are both separate sheets.");
    return;
  }
  
  //Calculate the amount of TOC sections
  var sections = getTableOfContentsSections(sourceSheet);
  
  //If the table of contents sheet is not the source sheet, continue with the operation.
  if (tocSheet != null) {
    Logger.log("Begin Writing to " + tocSheet.getName() + " | Number of sections: " + sections.length);
    
    var newValues = [];

    for (var i = 0; i < sections.length; i++) {
      var cellNotation = "A" + (i+1);
      var cellLink = "#gid=" + sourceSheet.getSheetId() + "&range=A" + (parseInt(sections[i][1])+1);
      var cellValue = "=HYPERLINK(\"" + cellLink + "\", \"" + sections[i][0] + "\")";
  
      newValues.push([cellValue]);
      Logger.log("Insert: " + newValues[i] + " at " + cellNotation);
    }
  
     tocSheet.getRange(1, 1, sections.length, 1).setValues(newValues);
     SpreadsheetApp.flush();
  } else {
    //Show an error message and cancel the operation
    showMessage("Error", "Table of Contents Sheet does not exist.");
  }
  
  //Refresh the app once we generate the new sheet
  //main();
  
  return tocSheet;
}

function quickSelect(sourceSheetIndex, sectionIndex) {
  var spreadSheet = SpreadsheetApp.getActive();
  var sheets = spreadSheet.getSheets();
  
  var sourceSheet = sheets[sourceSheetIndex];
  var tocSheet = spreadSheet.getSheetByName(sourceSheet.getRange("A1").getValue());
  
  var selectedSectionNotation = "A" + (parseInt(sectionIndex) + 1);
  var selectedSection = tocSheet.getRange(selectedSectionNotation).getFormula();
  
  var rangeCode = "&range=";
  var sectionNotation = selectedSection.substring(selectedSection.indexOf(rangeCode) + rangeCode.length, selectedSection.indexOf("\", \""));
  
  Logger.log(selectedSectionNotation + " (" + sectionIndex + ") " + selectedSection + " | " + sectionNotation);
  
  sourceSheet.activate();
  sourceSheet.setActiveSelection(sectionNotation);
}

/*
* Shows a message with instructions on how to use the addon
*/
function showInstructions() {
  showMessage("Table of Contents Generator Instructions & About", 
              "How to use:"
              + "\n• Column A will be used to designate sections of the Table of Contents."
              + "\n• Cell A1 designates the name of the Table of Contents sheet. It will be generated automatically if necessary or overwritten if already available."
              + "\n• Once you've set up your TOC cells, use this sidebar to select the sheet to generate the Table of Contents for."
              + "\n• Click the generate button to begin. If you modify sheets while the sidebar is open, click the refresh button to update the dropdown."
              + "\n• After a Table of Contents is generated, the Quick Select drop-down can be used to quickly cycle through sections and jump to them."
              + "\n\nTable of Contents Generator for Google Sheets | NOKORI·WARE, 2018, https://www.nokoriware.com/");
}

/*
* Utility function for showing a message pop-up
*/
function showMessage(title, message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}
