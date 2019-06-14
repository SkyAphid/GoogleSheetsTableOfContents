/*
*
* NOKORI·WARE Table of Contents Generator for Google Sheets
* December 2018
*
* This add-on will help work around the current lack of a proper Table of Contents system in Google Sheets.
*
*/

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Open Sidebar', 'NOKORIWARETableOfContentsGenerator');
  menu.addItem('Generate Table of Contents From Active Sheet', 'generateTableOfContents');
  menu.addItem('Sort Active Sheet Alphabetically By Section', 'sortSheet');
  menu.addItem('Open Instruction Manual', 'showManual');
  menu.addToUi();
}

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
      //Push: Section Name, Row
      sections.push([values[i][0], i+1]);
      //Logger.log("Section: " + sections[i][0] + " | Row: " + sections[i][1]);
    }
  }
  
  return sections;
}

/*
* Gets the section names of the default sheet in the Spreadsheet to populate the Table of Contents Quick Select in the sidebar.
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
function generateTableOfContents() {
  var spreadSheet = SpreadsheetApp.getActive();

  //debug("Call " + sourceSheetIndex);
  
  //var sourceSheet = sheets[sourceSheetIndex];
  var sourceSheet = spreadSheet.getActiveSheet();
  
  //The table of contents sheet is designated at A1 of the source sheet by name.
  var tocName = sourceSheet.getRange(1, 1).getValue();
  var tocSheet = spreadSheet.getSheetByName(tocName);
  
  if (tocSheet == null) {
    //If the sheet is null, create a new one
    tocSheet = spreadSheet.insertSheet(tocName);
  } else{
    //If the sheet exists, clear it to avoid de-syncing issues
    tocSheet.clear();
  }
  
  //Reactivate the source sheet so that it doesn't jerk you over to the generated table of contents (this gets annoying)
  sourceSheet.activate();
  
  //Also flush the app so that all changes are updated immediately
  SpreadsheetApp.flush();
  
  if (tocSheet == sourceSheet) {
    showMessage("Error", "Source sheet and Table of Contents sheet both match. Make sure they are both separate sheets.");
    return;
  }
  
  //Calculate the amount of TOC sections
  var sections = getTableOfContentsSections(sourceSheet);
  
  //If the table of contents sheet is not the source sheet, continue with the operation.
  if (tocSheet != null) {
    //Logger.log("Begin Writing to " + tocSheet.getName() + " | Number of sections: " + sections.length);
    
    var newValues = [];

    for (var i = 0; i < sections.length; i++) {
      var cellNotation = "A" + (i+1);
      var cellLink = "#gid=" + sourceSheet.getSheetId() + "&range=A" + parseInt(sections[i][1]);
      var cellValue = "=HYPERLINK(\"" + cellLink + "\", \"" + sections[i][0] + "\")";
  
      newValues.push([cellValue]);
      //Logger.log("Insert: " + newValues[i] + " at " + cellNotation);
    }
  
     tocSheet.getRange(1, 1, sections.length, 1).setValues(newValues);
     SpreadsheetApp.flush();
  } else {
    //Show an error message and cancel the operation
    showMessage("Error", "Table of Contents Sheet does not exist.");
  }
  
  //Refresh the app once we generate the new sheet
  NOKORIWARETableOfContentsGenerator();
  
  return tocSheet;
}

/*
* Duplicate the given the sheet and sort it by sectionIndices.
*/
function sortSheet() {
  var spreadSheet = SpreadsheetApp.getActive();
  
  //Source sheet data
  var sourceSheet = spreadSheet.getActiveSheet();
  var sourceSheetLastRow = sourceSheet.getLastRow()+1;
  var sourceSheetLastColumn = sourceSheet.getLastColumn()+1;
  
  /*
  * Create an empty sheet we can place sorted elements into
  */
  
  var sortSheetName = sourceSheet.getName() + " - Sorted";
  
  //Delete the old sort sheet if one already exists
  var existingSortSheet = spreadSheet.getSheetByName(sortSheetName);
  if (existingSortSheet) {
    spreadSheet.deleteSheet(existingSortSheet);
  }
  
  //Create our new base sort sheet to begin working on
  var sortSheet = sourceSheet.copyTo(spreadSheet);
  sortSheet.setName(sortSheetName);
  sortSheet.clear();
  
  //Activate the sort sheet so it's in view
  sortSheet.activate();
  
  //Flush the app so that the changes are reflected immediately
  SpreadsheetApp.flush();
  
  //Refresh the app once we generate the new sheet
  NOKORIWARETableOfContentsGenerator();
  
  //Show a confirmation message
  showMessage("Sorting " + sourceSheet.getName(), "Sorting may take a moment to complete.\nA confirmation message will be shown once the new sheet is finished generating.\nPlease wait until the process is completed.");
  
  /*
  * Fetch sections and sort them alphabetically
  */
  
  var sourceSections = getTableOfContentsSections(sourceSheet);
  
  //Slide everything from A2 downwards (the Table of Contents header must always be at the top)
  var slicedSections = sourceSections.slice(1, sourceSections.length);
  slicedSections.sort();
  
  //Compile the final sorted sections array
  var sortedSections = sourceSections.slice(0, 1).concat(slicedSections);
  
  /*
  * Begin creating our new sorted sheet
  */
  
  var currentRow = 1;
  
  for (var i = 0; i < sortedSections.length; i++){
    //Logger.log(i + " Original: " + sections[i][0] + " | Sorted: " + sortedSections[i][0]);
    
    //Get the original data of the section for pasting into the new one
    var sortedSectionName = sortedSections[i][0];
    var sortedSectionRow = sortedSections[i][1];
    
    var numSectionRows = getSectionLength(sourceSheet.getLastRow()+1, sourceSections, sortedSectionName);
    var numSectionColumns = sourceSheet.getLastColumn()+1;
    
    var sourceSheetRange = sourceSheet.getRange(sortedSectionRow, 1, numSectionRows, numSectionColumns);    
    
    //Paste the range into the sorted sheet
    var sortSheetRange = sortSheet.getRange(currentRow, 1);
    sourceSheetRange.copyTo(sortSheetRange);
    
    currentRow += numSectionRows;
  }
  
  showMessage("Sorting Complete", sourceSheet.getName() + " has finished sorting.");
}

/*
* Calculate the length of the section based on its row start index and the row start index of the section after.
*/
function getSectionLength(sheetLastRow, sections, sectionName){
  for (var i = 0; i < sections.length; i++) {
    if (sections[i][0] == sectionName) {
      var sectionRow = sections[i][1]
      var nextSectionRow = (i + 1 < sections.length ? sections[i+1][1] : sheetLastRow);
      
      return (nextSectionRow - sectionRow);
    }
  }
  
  return 0;
}

/*
* Sidebar quick select functionality. Input a source sheet and section index to jump to it.
*/
function quickSelect(sourceSheetIndex, sectionIndex) {
  var spreadSheet = SpreadsheetApp.getActive();
  var sheets = spreadSheet.getSheets();
  
  var sourceSheet = sheets[sourceSheetIndex];
  var tocSheet = spreadSheet.getSheetByName(sourceSheet.getRange("A1").getValue());
  
  var selectedSectionNotation = "A" + (parseInt(sectionIndex) + 1);
  var selectedSection = tocSheet.getRange(selectedSectionNotation).getFormula();
  
  var rangeCode = "&range=";
  var sectionNotation = selectedSection.substring(selectedSection.indexOf(rangeCode) + rangeCode.length, selectedSection.indexOf("\", \""));
  
  //Logger.log(selectedSectionNotation + " (" + sectionIndex + ") " + selectedSection + " | " + sectionNotation);
  
  sourceSheet.activate();
  sourceSheet.setActiveSelection(sectionNotation);
}

/*
* Shows a message with instructions on how to use the addon
*/
function showManual() {
  showMessage("Table of Contents Generator Manual", 
              "Formatting Your Sheet For The Generator:"
              + "\n• Column A will be used to designate sections of the sheet (i.e. chapter titles and the pages they're on)."
              + "\n• Cell A1 designates the name of the generated Table of Contents sheet."
              + "\n• Once you've set up your section cells, click 'Generate Table of Contents' while the sheet is active to begin."
              + "\n• A new sheet will be generated containing the Table of Contents data, which will be used by this sidebar."
              + "\n\nUsing The Table of Contents:"
              + "\n• The 'Table of Contents Quick Select' drop-down allows for quick cycling through large sheets."
              + "\n• Select the sheet you want to cycle through from the 'Quick Select Sheet' drop-down." 
              + "\n• After selecting a sheet, if a corresponding Table of Contents is available, the 'Table of Contents Quick Select' drop-down can be used to quickly cycle through sections and jump to them."
              + "\n\nSorting Sheets"
              + "\n• You can sort your sheet alphabetically by section if you press the 'Sort' button."
              + "\n• The script duplicates the active sheet and sorts the copy in the event that you don't want to keep the changes. Otherwise, you can simply delete the original and rename the new sorted duplicate."
              + "\n\nAdditional Notes:"
              + "\n• If you create a new sheet or delete one, click 'Refresh' to update the 'Quick Select Sheet' dropdown."
              + "\n\nTable of Contents Generator for Google Sheets"
              + "\n2018-2019 :: NOKORI·WARE :: https://www.nokoriware.com/");
}

/*
* Utility function for showing a message pop-up
*/
function showMessage(title, message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}
