function maker(inType) {
  var onSheet = SpreadsheetApp.getActive(); // always defined at beginning
  var runSheet = onSheet.getActiveSheet(); // Macro Input+Run sheet
  var inputSheetName = runSheet.getRange('B2').getValue(); // the sheet with all kids data
  var weekNum = runSheet.getRange('B3').getValue();
  var date = runSheet.getRange('B4').getValue();
  var combinedQ = runSheet.getRange('B5').getValue();
  var nametagQ = runSheet.getRange('B6').getValue();
  var test = onSheet.getSheetByName(inputSheetName);
  if (!test) { // https://stackoverflow.com/questions/48974770/check-for-existence-of-a-sheet-in-google-spreadsheets
	throw new Error("Input sheet does not exist!");
  } else if (combinedQ === '') {
    throw new Error("Lunch and Allergy combination not specified!");
  } else if (nametagQ === '') {
    throw new Error("Nametag generation not specified!");
  }
  if (inType === 2) {
    if (runSheet.getRange('B7').getValue() === '') {
      throw new Error("Day of the Week cell cannot be blank!")
    }
  }
  if (!(Object.prototype.toString.call(date) === '[object Date]')) { // refer to https://stackoverflow.com/questions/643782/how-to-check-whether-an-object-is-a-date
    throw new Error("Monday date entered is not a valid date!");
  }
  var sheetNames = [];
  var dow = [];
  if (inType === 1) {
    for (var f = 0; f < 6; f++) {
      sheetNames[f] = runSheet.getRange(4, 4+f).getValue();
      dow[f] = runSheet.getRange(3, 4+f).getValue();
    }
  } else {
    sheetNames[0] = runSheet.getRange('D7').getValue();
    dow[0] = runSheet.getRange('E7').getValue();
  }
  var sourceSheet = onSheet.getSheetByName(inputSheetName);
  onSheet.setActiveSheet(sourceSheet, true);
  onSheet.moveActiveSheet(1); // sheet index is 1-based.
  sourceSheet.getFilter().sort(3, true); // sort kids by first name first (col C)
  sourceSheet.getFilter().sort(4, true); // sort kids by last name (column D)
  var filRange = sourceSheet.getFilter().getRange();
  var kids;
  for (var sheetCount = 0; sheetCount < sheetNames.length; sheetCount++) {
    if (dow[sheetCount] === 5 && sheetNames[sheetCount].indexOf('TRIP') >= 0) {
      kids = sourceSheet.getRange(filRange.getLastRow()+7,dow[sheetCount]+10).getValue(); // kids on trip cell
    } else {
      kids = sourceSheet.getRange(filRange.getLastRow()+11,dow[sheetCount]+10).getValue();
    }
    onSheet.setActiveSheet(sourceSheet,true); // ensure generate function is called on the correct sheet
    generateDayAttendance(onSheet,sourceSheet,sheetNames[sheetCount],dow[sheetCount],filRange,kids);
    onSheet.moveActiveSheet(sheetCount+1); // sorts the newly generated sheet correctly if Full Week
  }
  if (inType === 1) {
    // lunch and allergies are skipped if single day was selected
    generateLunchAllergies(onSheet,sourceSheet,combinedQ,weekNum,filRange);
  }
  if (nametagQ === 'Yes') {
    genNametags(onSheet,sourceSheet,filRange,weekNum);
  }
  onSheet.setActiveSheet(runSheet, true); // end on the run sheet
}

function makeFullWeek() {
  maker(1);
  // linked to Generate Full Week button
}

function makeSingleDay() {
  maker(2);
  // linked to Generate Single Day button, requires input in the Single Day cell
}

function generateDayAttendance(spreadsheet,inSheet,outSheetName,dayofweek,filteredRange,numKids) {
  var mainCopyRange = inSheet.getRange(filteredRange.getRow()+1,filteredRange.getColumn()+5,filteredRange.getLastRow()-1,5); // filtered range's row index AFTER the header
  var dayCopyRange = inSheet.getRange(filteredRange.getRow()+1,10+dayofweek,filteredRange.getLastRow()-1,1); // the FD/HD + lunch data is stored in this column
  var allergyRange = inSheet.getRange(filteredRange.getRow()+1,18,filteredRange.getLastRow()-1,1); // allergy range is col 18
  var contactRange = inSheet.getRange(filteredRange.getRow()+1,22,filteredRange.getLastRow()-1,1); // formatted contact range is col 18
  var hideBlanks = SpreadsheetApp.newFilterCriteria() // use for Mon to Thurs
  .setHiddenValues([''])
  .build();
  var tripFilter = SpreadsheetApp.newFilterCriteria() // use for Fri Studio
  .setHiddenValues(['','Trip'])
  .build();
  var lookForTrip = SpreadsheetApp.newFilterCriteria() // use for Fri Trip
  .whenTextContains('Trip')
  .build();
  var trip = false;
  var testing = spreadsheet.getSheetByName(outSheetName); // see if there's already a sheet with the desired output name, and then just rename it hehe
  if (testing) {
    spreadsheet.getSheetByName(outSheetName).setName('OLD COPY ' + outSheetName);
  }
  if (dayofweek < 5) {
    spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(10+dayofweek, hideBlanks); // day of week (Mon) starts at col 11 to 15.
	spreadsheet.setActiveSheet(spreadsheet.getSheetByName('TEMPLATE Daily'), true);
  } else if (outSheetName.indexOf('TRIP') >= 0) {
    // if outSheetName contains 'TRIP'
    // https://stackoverflow.com/questions/6629728/check-if-a-string-has-a-certain-piece-of-text
    spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(10+dayofweek, lookForTrip);
	spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Template TRIP'), true);
    trip = true;
  } else {
	spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(10+dayofweek, tripFilter); // Friday Studio case
	spreadsheet.setActiveSheet(spreadsheet.getSheetByName('TEMPLATE Daily'), true);
  }
  spreadsheet.duplicateActiveSheet(); // duplicates active sheet and then sets the duplicate as the active sheet
  spreadsheet.getActiveSheet().setName(outSheetName);
  spreadsheet.moveActiveSheet(1); // sheet is moved to front so even if the code errors, we can get it easily
  if (trip) {
    // pasting is slightly different for trip sheet, because no teams
	// then need to duplicate staff sheet, rename accordingly, then paste in data
    if (numKids > 25) {
	  var rowsAdd = (Math.floor((numKids-1)/5) - 4) * 5; // should always be a multiple of 5
	  spreadsheet.getActiveSheet().insertRowsBefore(30, rowsAdd);
	  spreadsheet.getRange('A3:A4').activate(); // for how to concatenate strings in JS https://www.geeksforgeeks.org/javascript-string-prototype-concat-function/
      spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A3:A'.concat(32+rowsAdd)), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); // auto-fill code, 32+rowsadd should always be a number that is 2 more than a multiple of 5 (30+2, 2 rows to compensate for header + dummy row)
    } else if (numKids < 21) {
	  var rowsDel = (4 - Math.floor((numKids-1)/5)) * 5;
	  spreadsheet.getActiveSheet().deleteRows(33 - rowsDel, rowsDel); // row that entry #30 is on (row #32) + 1 = row #33.
    }
    mainCopyRange = inSheet.getRange(filteredRange.getRow()+1,filteredRange.getColumn()+5,filteredRange.getLastRow()-1,3);
    dayCopyRange = inSheet.getRange(filteredRange.getRow()+1,10,filteredRange.getLastRow()-1,1);
    mainCopyRange.copyTo(spreadsheet.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    dayCopyRange.copyTo(spreadsheet.getRange('F2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
	allergyRange.copyTo(spreadsheet.getRange('K2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getActiveSheet().deleteRow(2);
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Template TRIP STAFF'), true);
    spreadsheet.duplicateActiveSheet();
    spreadsheet.getActiveSheet().setName(outSheetName.concat(' STAFF'));
    spreadsheet.moveActiveSheet(2); // staff sheet is put in second place
    if (numKids > 25) {
	  spreadsheet.getActiveSheet().insertRowsBefore(30, rowsAdd);
	  spreadsheet.getRange('A3:A4').activate(); // for how to concatenate strings in JS https://www.geeksforgeeks.org/javascript-string-prototype-concat-function/
      spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A3:A'.concat(32+rowsAdd)), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); // auto-fill code, 32+rowsadd should always be a number that is 2 more than a multiple of 5 (30+2, 2 rows to compensate for header + dummy row)
    } else if (numKids < 21) {
	  spreadsheet.getActiveSheet().deleteRows(33 - rowsDel, rowsDel); // row that entry #30 is on (row #32) + 1 = row #33.
    }
    mainCopyRange.copyTo(spreadsheet.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    dayCopyRange.copyTo(spreadsheet.getRange('F2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
	allergyRange.copyTo(spreadsheet.getRange('G2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    contactRange.copyTo(spreadsheet.getRange('C2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getActiveSheet().deleteRow(2);
  } else {
    if (numKids > 35) {
	  var rowsAdd = (Math.floor((numKids-1)/5) - 6) * 5; // should always be a multiple of 5
	  spreadsheet.getActiveSheet().insertRowsBefore(40, rowsAdd);
	  spreadsheet.getRange('A3:A4').activate(); // for how to concatenate strings in JS https://www.geeksforgeeks.org/javascript-string-prototype-concat-function/
      spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A3:A'.concat(42+rowsAdd)), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); // auto-fill code, 42+rowsadd should always be a number that is 2 more than a multiple of 5 (40+2, 2 rows to compensate for header + dummy row)
    } else if (numKids < 31) {
	  var rowsDel = (6 - Math.floor((numKids-1)/5)) * 5;
	  spreadsheet.getActiveSheet().deleteRows(43 - rowsDel, rowsDel); // row that entry #40 is on (row #42) + 1 = row #43.
    }
    mainCopyRange.copyTo(spreadsheet.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    dayCopyRange.copyTo(spreadsheet.getRange('G2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
	allergyRange.copyTo(spreadsheet.getRange('M2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    var dayType;
    for (var i = 2; i < numKids+3; i++) { // in a future update - this can be sped up by getting the array values of these columns and making the changes to the array, then re-writing the column. having to re-read values is slow
      var dayTypeRange = spreadsheet.getRange('G' + i);
      dayType = dayTypeRange.getValue();
      if (dayType.indexOf('L') >= 0) {
        spreadsheet.getRange('H' + i).setValue('Y');
        dayTypeRange.setValue(dayType.substring(0,dayType.length - 2));
        dayType = dayTypeRange.getValue();
      }
      if (dayType.indexOf('FD') >= 0) {
        dayTypeRange.setValue('-');
      }
    }
    spreadsheet.getActiveSheet().deleteRow(2);
    spreadsheet.getActiveSheet().hideColumns(4);
  }
  // spreadsheet.getActiveSheet().autoResizeColumns(6, 1); // auto-resize Days column but it doesn't work well so it's disabled
  inSheet.getFilter().removeColumnFilterCriteria(10+dayofweek);
}

function genLunchAller() {
  // must check if the input sheet title is valid
  var onSheet = SpreadsheetApp.getActive(); // always defined at beginning
  var runSheet = onSheet.getActiveSheet(); // Macro Input+Run sheet
  var inputSheetName = runSheet.getRange('B2').getValue(); // the sheet with all kids data
  var weekNum = runSheet.getRange('B3').getValue(); // week # specified
  var combinedQ = runSheet.getRange('B5').getValue();
  var test = onSheet.getSheetByName(inputSheetName);
  if (!test) { // https://stackoverflow.com/questions/48974770/check-for-existence-of-a-sheet-in-google-spreadsheets
	throw new Error("Input sheet does not exist!");
  } else if (combinedQ === '') {
    throw new Error("Lunch and Allergy combination not specified!");
  }
  var sourceSheet = onSheet.getSheetByName(inputSheetName);
  onSheet.setActiveSheet(sourceSheet, true);
  onSheet.moveActiveSheet(1); // sheet index is 1-based.
  sourceSheet.getFilter().sort(3, true); // sort kids by first name first (col C)
  sourceSheet.getFilter().sort(4, true); // sort kids by last name (column D)
  var filRange = sourceSheet.getFilter().getRange();
  generateLunchAllergies(onSheet,sourceSheet,combinedQ,weekNum,filRange);
}

function generateLunchAllergies(spreadsheet,inSheet,quest,week,filteredRange) { // filtered range's row index AFTER the header
  var nameCopyRange = inSheet.getRange(filteredRange.getRow()+1,6,filteredRange.getLastRow()-1,1); // names column (Last, First) (col F or 6)
  var ageCopyRange = inSheet.getRange(filteredRange.getRow()+1,8,filteredRange.getLastRow()-1,2); // age + team data (cols 8 & 9, H & I)
  var dayRange = inSheet.getRange(filteredRange.getRow()+1,10,filteredRange.getLastRow()-1,1); // Days column (which days are being attended) (col 10, J)
  var infoRange = inSheet.getRange(filteredRange.getRow()+1,11,filteredRange.getLastRow()-1,5); // holds the lunch data as "L" appended to the back of the camp type
  var allergyRange = inSheet.getRange(filteredRange.getRow()+1,18,filteredRange.getLastRow()-1,1); // allergy range is col 18 or R
  var contactRange = inSheet.getRange(filteredRange.getRow()+1,7,filteredRange.getLastRow()-1,1); // contact col (col G or 7)
  var lunchNum = inSheet.getRange(filteredRange.getLastRow()+11,16).getValue(); // lunch in column 16 or P
  var allerNum = inSheet.getRange(filteredRange.getLastRow()+11,18).getValue(); // allergies in column 18 or R
  var oneSheetName;
  var lunchRows;
  var hideBlanks = SpreadsheetApp.newFilterCriteria() // use for finding lunch ppl
  .setHiddenValues([''])
  .build();
  var hideAller = SpreadsheetApp.newFilterCriteria() // use for finding allergy ppl
  .setHiddenValues(['','-'])
  .build();
  inSheet.getFilter().setColumnFilterCriteria(16,hideBlanks); // find lunch
  if (quest === 'Yes') { // duplicate the combined sheet
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Template COMBINED'), true);
    spreadsheet.duplicateActiveSheet(); // duplicates active sheet and then sets the duplicate as the active sheet
    spreadsheet.moveActiveSheet(1); // sheet is moved to front so even if the code errors, we can get it easily
    oneSheetName = 'W'.concat(week,' Lunch+Allergies');
    lunchRows = 21;
  } else {
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Template LUNCH'), true);
    spreadsheet.duplicateActiveSheet(); // duplicates active sheet and then sets the duplicate as the active sheet
    
    oneSheetName = 'W'.concat(week,' Lunch');
    lunchRows = 26;
  }
  var testing = spreadsheet.getSheetByName(oneSheetName); // see if there's already a sheet with the desired output name, and then just rename it hehe
  if (testing) {
    spreadsheet.getSheetByName(oneSheetName).setName('OLD COPY ' + oneSheetName);
  }
  spreadsheet.getActiveSheet().setName(oneSheetName);
  var deltaRows = 0;
  if (lunchNum > (lunchRows - 6)) { // need to add rows
    deltaRows = ((Math.floor(lunchNum/5)) - (((lunchRows-1)/5)-2)) * 5; // should always be a multiple of 5
	spreadsheet.getActiveSheet().insertRowsBefore(lunchRows, deltaRows);
	spreadsheet.getRange('A3:A4').activate(); // for how to concatenate strings in JS https://www.geeksforgeeks.org/javascript-string-prototype-concat-function/
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A3:A'.concat(lunchRows+1+deltaRows)), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); // auto-fill code
  } else if (lunchNum < (lunchRows - 10)) {
    deltaRows = -1 * (((((lunchRows-1)/5)-2) - Math.floor(lunchNum/5))) * 5;
    spreadsheet.getActiveSheet().deleteRows(lunchRows+2+deltaRows, (-1 * deltaRows));
  }
  nameCopyRange.copyTo(spreadsheet.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  ageCopyRange.copyTo(spreadsheet.getRange('C2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  infoRange.copyTo(spreadsheet.getRange('E2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  allergyRange.copyTo(spreadsheet.getRange('J2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  var dayType;
  var dayTypeRange;
  for (var i = 2; i < lunchNum + 3; i++) {
    for (var day = 5; day < 10; day++) {
      dayTypeRange = spreadsheet.getActiveSheet().getRange(i, day);
      dayType = dayTypeRange.getValue();
      if (dayType.indexOf('L') >= 0) {
        dayTypeRange.setValue('x');
      } else {
        dayTypeRange.clearContent();
      }
    }
  }
  inSheet.getFilter().removeColumnFilterCriteria(16);
  var lastRowLunch = lunchRows+1+deltaRows;
  
  // start allergy section
  
  inSheet.getFilter().setColumnFilterCriteria(18,hideAller); // find allergy ppl
  var twoSheetName;
  var allerRows = 16;
  var allerRowLocation;
  var firstRowaller = lastRowLunch + 7; // where entry #1 (not the blank row) is in allergy area
  if (quest === 'Yes') { // no need to change sheets
    twoSheetName = oneSheetName;
    allerRowLocation = lastRowLunch + 20;
  } else {
    firstRowaller = 3; // reset the firstRow of allergies to A3 if it has own sheet
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Template ALLERGIES'), true);
    spreadsheet.duplicateActiveSheet(); // duplicates active sheet and then sets the duplicate as the active sheet
    twoSheetName = 'W'.concat(week,' Allergies');
    allerRowLocation = 16;
    var testing2 = spreadsheet.getSheetByName(twoSheetName); // see if there's already a sheet with the desired output name, and then just rename it hehe
    if (testing2) {
      spreadsheet.getSheetByName(twoSheetName).setName('OLD COPY ' + twoSheetName);
    }
    spreadsheet.getActiveSheet().setName(twoSheetName);
  }
  
  deltaRows = 0;
  if (allerNum > (allerRows - 6)) { // need to add rows
    deltaRows = ((Math.floor(allerNum/5)) - (((allerRows-1)/5)-2)) * 5; // should always be a multiple of 5
	spreadsheet.getActiveSheet().insertRowsBefore(allerRowLocation, deltaRows);
    spreadsheet.getRange('A'.concat(firstRowaller,':A',firstRowaller+1)).activate(); // for how to concatenate strings in JS https://www.geeksforgeeks.org/javascript-string-prototype-concat-function/
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A'.concat(firstRowaller,':A',allerRowLocation+1+deltaRows)), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); // auto-fill code
    if (quest === 'Yes') {
      spreadsheet.getActiveSheet().getRange('E'.concat(firstRowaller,':G',firstRowaller)).copyTo(spreadsheet.getRange('E'.concat((firstRowaller-1),':G',allerRowLocation+1+deltaRows)), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    }
  } else if (allerNum < (allerRows - 10)) {
    deltaRows = -1 * ((((allerRows-1)/5)-2) - Math.floor(allerNum/5)) * 5;
    spreadsheet.getActiveSheet().deleteRows(allerRowLocation+2+deltaRows, (-1 * deltaRows));
  }
  // can now copy
  nameCopyRange.copyTo(spreadsheet.getRange('B'.concat(firstRowaller-1)), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  ageCopyRange.copyTo(spreadsheet.getRange('C'.concat(firstRowaller-1)), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  contactRange.copyTo(spreadsheet.getRange('E'.concat(firstRowaller-1)), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  if (quest === 'Yes') {
    dayRange.copyTo(spreadsheet.getRange('H'.concat(firstRowaller-1)), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    allergyRange.copyTo(spreadsheet.getRange('J'.concat(firstRowaller-1)), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  } else {
    dayRange.copyTo(spreadsheet.getRange('F'.concat(firstRowaller-1)), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    allergyRange.copyTo(spreadsheet.getRange('G'.concat(firstRowaller-1)), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }
  inSheet.getFilter().removeColumnFilterCriteria(18);
  
  // delete dummy rows
  if (quest === 'Yes') {
    spreadsheet.getSheetByName(oneSheetName).deleteRow(firstRowaller-1);
  } else {
    spreadsheet.getSheetByName(twoSheetName).deleteRow(2);
    spreadsheet.getSheetByName(twoSheetName).hideColumns(3);
  }
  spreadsheet.getSheetByName(oneSheetName).deleteRow(2);
  spreadsheet.getSheetByName(oneSheetName).hideColumns(3);
}

function generateNametags() { // linked to Generate Name Tags only button, also called in makeFullWeek and makeSingleDay if Generate Name Tags input cell is "Yes"
  // must check if the input sheet title is valid
  var onSheet = SpreadsheetApp.getActive(); // always defined at beginning
  var runSheet = onSheet.getActiveSheet(); // Macro Input+Run sheet
  var inputSheetName = runSheet.getRange('B2').getValue(); // the sheet with all kids data
  var weekNum = runSheet.getRange('B3').getValue(); // week # specified
  var test = onSheet.getSheetByName(inputSheetName);
  if (!test) { // https://stackoverflow.com/questions/48974770/check-for-existence-of-a-sheet-in-google-spreadsheets
	throw new Error("Input sheet does not exist!");
  }
  var sourceSheet = onSheet.getSheetByName(inputSheetName);
  onSheet.setActiveSheet(sourceSheet, true);
  onSheet.moveActiveSheet(1); // sheet index is 1-based.
  sourceSheet.getFilter().sort(3, true); // sort kids by first name first (col C)
  sourceSheet.getFilter().sort(4, true); // sort kids by last name (column D)
  var filRange = sourceSheet.getFilter().getRange();
  genNametags(onSheet,sourceSheet,filRange,weekNum);
}

function genNametags(spreadsheet,inSheet,filteredRange,week) {
  var numKids = inSheet.getRange(filteredRange.getLastRow()+11,2).getValue(); // how many nametags to generate
  var numPages = Math.ceil(numKids / 12);
  var kidsLast = numKids - (Math.ceil(numKids / 12) - 1) * 12; // # kids on the last page
  var rowsDel = Math.floor((12 - kidsLast)/2); // # nametag rows (which are 6 rows high)
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Template NAMETAG'), true);
  spreadsheet.duplicateActiveSheet(); // duplicates active sheet and then sets the duplicate as the active sheet
  var outSheetName = 'W'.concat(week,' Nametags');
  var testing = spreadsheet.getSheetByName(outSheetName); // see if there's already a sheet with the desired output name, and then just rename it hehe
  if (testing) {
    spreadsheet.getSheetByName(outSheetName).setName('OLD COPY ' + outSheetName);
  }
  spreadsheet.getActiveSheet().setName(outSheetName);
  spreadsheet.moveActiveSheet(1); // sheet is moved to front so even if the code errors, we can get it easily
  // one nametag page can create 12 nametags, range A1:D37
  for (var p = 0; p < numPages; p++) {
    spreadsheet.getRange('A1:D36').copyTo(spreadsheet.getRange('A'.concat(p*36 + 1)), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }
  if (!(rowsDel === 0)) {
    spreadsheet.getActiveSheet().deleteRows((numPages * 36 + 1 - rowsDel * 6), (rowsDel * 6)); // row that entry #40 is on (row #42) + 1 = row #43.
  }
  var tagRow = -1; // which row of nametags, summed across all sheets
  var tagCol = 3; // column A or C
  var day; // pre-initialized
  var age; // pre-initialized
  var aller = ""; // pre-initialized to be empty string
  var dayAgeAller;
  var testMe;
  var testMeMore;
  for (var t = filteredRange.getRow(); t < filteredRange.getLastRow(); t++) {
    testMe = inSheet.getRange('B'.concat(t+1)).getValue();
    testMeMore = testMe.indexOf('Y');
    if (!(inSheet.getRange('B'.concat(t+1)).getValue().indexOf('Y') >= 0)) {
      // make-shift for loop block
      if (tagCol === 3) { // switch columns for each kid
        tagCol = 1;
        tagRow++; // only change row if the column was changed to position 1
      } else {
        tagCol = 3; // keep same row
      }
      // makeshift for loop block ends
      inSheet.getRange('E'.concat(t+1)).copyTo(spreadsheet.getActiveSheet().getRange(tagRow*6+2,tagCol), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      day = inSheet.getRange('Q'.concat(t+1)).getValue();
      age = inSheet.getRange('H'.concat(t+1)).getValue();
      if (!(inSheet.getRange('R'.concat(t+1)).getValue() === '-')) {
        aller = ' '.concat(inSheet.getRange('R'.concat(t+1)).getValue());
      } else {
        aller = "";
      }
      dayAgeAller = day.concat(' (',age,')',aller);
      spreadsheet.getActiveSheet().getRange(tagRow*6+3,tagCol).setValue(dayAgeAller);
      inSheet.getRange('V'.concat(t+1)).copyTo(spreadsheet.getActiveSheet().getRange(tagRow*6+5,tagCol), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      spreadsheet.getActiveSheet().setRowHeight(tagRow*6+5, 53);
    }
  }
}
