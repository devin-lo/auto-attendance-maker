/** @OnlyCurrentDoc */

function makeAttendance() {
  // declare general-use variables
  var onSheet = SpreadsheetApp.getActive();
  var runSheet = onSheet.getSheetByName("Attendance Maker"); // Macro Input+Run sheet
  var inputSheet = onSheet.getSheetByName("Form Responses 1");
  var outDate = runSheet.getRange('B2').getValue(); // output date, need this to set filter later. May have a string in it ("TRIP - ")
  var dayColumn = runSheet.getRange('D3').getValue(); // column for the single day we are analyzing - now being pulled by a formula as of March Break 2020
  var location = runSheet.getRange('B3').getValue(); // location that we are making the attendance for - new feature as of Dec 20 2019
  var weeksTog = runSheet.getRange('B4').getValue(); // does the camp have a full-week section? New feature as of March Break 2020
  var lunchTog = runSheet.getRange('B5').getValue(); // lunch, new feature as of March Break 2020
  var teams = runSheet.getRange('B6').getValue(); // teams, new feature as of March Break 2020
  var outName = runSheet.getRange('G2').getValue(); // the output sheet's name
  var outSheet = onSheet.getSheetByName(outName);
  var filRange = inputSheet.getFilter().getRange();
  var remBlanks = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues([''])
  .build();
  
  var toggleU = onSheet.getRange('E18:E27').getValues(); // 2D array of certain option toggles
  var toggles = []; // will be populated back into a 1D array, https://stackoverflow.com/questions/14824283/convert-a-2d-javascript-array-to-a-1d-array
  for (var i = 0; i < toggleU.length; i++)
  {
    toggles = toggles.concat(toggleU[i][0]);
  }
  
  var infoCol = []; // all the column values for which data is stored on the sheet. Each index's contents are listed below:
  /*
     0 = 1st kid name
     1 = 1st kid age
     2 = 1st kid allergies
     3 = 2nd kid name
     4 = 2nd kid age
     5 = 2nd kid allergies
     6 = parent 1 phone number
     7 = lunch
     8 = location that the kid signed up for
     9 = field trip
  */
  for (var i = 0; i < toggles.length; i++) {
    if (toggles[i] === 'Y') {
      infoCol = infoCol.concat(runSheet.getRange(18+i,3).getValue());
    }
    else {
      infoCol = infoCol.concat('x');
    }
  }
  
  // if the date has a Friday field trip, then it's not going to have the meal plan. Also, it needs to have the field trip column toggle on or else it will error.
  if (typeof outDate === 'string') {
    var outDatelower = outDate.toLowerCase();
    if (outDatelower.indexOf('trip') !== -1) {
      if (infoCol[9] === 'x') {
        throw new Error("Field trip column toggle is off, please turn it on and try again.");
      }
      lunchTog = 'No';
      if (location === 'Aurora') { // prevent user from creating an Aurora field trip attendance.
        throw new Error("Aurora campus doesn't have field trip!");
      }
    }
    else if (outDatelower.indexOf('studio') !== -1) {
      if (infoCol[9] === 'x') {
        throw new Error("Field trip column toggle is off, please turn it on and try again.");
      }
      lunchTog = 'No';
    }
  }
  
  // next set of code relates to the location feature, added Dec 20 2019.
  if (infoCol[8] !== 'x') {
    var locFilter = SpreadsheetApp.newFilterCriteria()
    .whenTextContains(location)
    .build();
    var aurOffset = 0;
    if (location === 'Aurora') {
      
      aurOffset = 15; // Aurora's attendance #'s are 15 rows below Markham's in the input sheet. Last changed in March Break camp, in preparation for summer camp.
      
      lunchTog = 'No'; // Aurora campus has no meal plan at the moment.
      
    }
  }
  
  
  // create the attendance sheet for that day, if it's missing
  if (!outSheet) { // https://stackoverflow.com/questions/48974770/check-for-existence-of-a-sheet-in-google-spreadsheets
	var tempSheet = onSheet.getSheetByName("ATTENDANCE TEMPLATE");
    if (!tempSheet) {
      throw new Error("Template sheet is missing!"); // eventually come back and hard-code a template sheet.
    }
    onSheet.setActiveSheet(tempSheet, true);
    onSheet.duplicateActiveSheet();
    onSheet.renameActiveSheet(outName);
    onSheet.moveActiveSheet(tempSheet.getIndex());
    outSheet = onSheet.getSheetByName(outName);
  }

  // setting up input sheet filters, and gathering info to prepare blank attendance sheet
  onSheet.setActiveSheet(inputSheet, true);
  if (infoCol[8] !== 'x') {
    inputSheet.getFilter().setColumnFilterCriteria(infoCol[8], locFilter);
  }
  var maxKids = inputSheet.getRange(filRange.getLastRow()+aurOffset+9,dayColumn).getValue(); // this is for the entire day in total
  var maxSiblings = inputSheet.getRange(filRange.getLastRow()+aurOffset+10,dayColumn).getValue(); // # of siblings in the entire day, autocalculated on sheet
  
  // prepare blank attendance sheet
  onSheet.setActiveSheet(outSheet, true);
  outSheet.getRange('A:A').clearContent();
  if (outSheet.getLastRow() > 1) {
    onSheet.duplicateActiveSheet(); // duplicate that sheet if there's content. If there's no content, getLastRow will return as 1 because only the first row has content.
  }
  onSheet.setActiveSheet(outSheet, true);
  outSheet.getRange(2,1,outSheet.getMaxRows()-1,outSheet.getMaxColumns()).clearContent();
  outSheet.getRange('B:C').activate();
  onSheet.getActiveRangeList().setBackground(null); // remove any sibling colouring because it will get changed
  
  outSheet.showColumns(1, outSheet.getMaxColumns()); // unhide all columns
  
  var lastRow = outSheet.getMaxRows();
  
  // add rows as needed.
  if (maxKids > lastRow-6) {
    var rowsAdd = (Math.ceil(((maxKids)+1)/5) - 4) * 5; // should always be a multiple of 5
    outSheet.insertRowsBefore(lastRow, rowsAdd); // should already be formatted right
  }
  
  var singleSiblings;
  var fieldOffset = 0; // offset variable for reading the correct row if the day is a trip or studio option one.
  
  // full-week options if they exist
  if (weeksTog === 'Yes') {
    var outWeek = inputSheet.getRange(1, dayColumn).getValue();
    var weekColumn = findCorrectWeek(outWeek, onSheet.getRangeByName("weekRange"));
    
    var fieldFilter;
    if (typeof outDate === 'string') {
      if (outDatelower.indexOf('studio') !== -1) {
        fieldFilter = SpreadsheetApp.newFilterCriteria()
        .setHiddenValues(['Yes']) // doing it this way will prevent erroring in an Aurora location attendance
        .build();
                
        fieldOffset = 2;
      }
      else if (outDatelower.indexOf('trip') !== -1) {
        fieldFilter = SpreadsheetApp.newFilterCriteria()
        .whenTextContains('Yes')
        .build();
        
        fieldOffset = 4;
      }
    }
    
    var fdKids = inputSheet.getRange(filRange.getLastRow()+aurOffset+fieldOffset+9,weekColumn).getValue(); // this is for the full week kids in total
    var fdSiblings = inputSheet.getRange(filRange.getLastRow()+aurOffset+fieldOffset+10,weekColumn).getValue(); // # of siblings registered for the full week, autocalculated on sheet
    
    singleSiblings = maxSiblings - fdSiblings; // true # of single day siblings
  }
  else { // if there are no full-week options
    singleSiblings = maxSiblings;
  }
  
  // copypaste for single day kids first
  inputSheet.getFilter().setColumnFilterCriteria(dayColumn, remBlanks); // set filter immediately before copy-paste action.
  if ((maxKids - fdKids) > 0) {
    copyPasteInfo(1,onSheet,inputSheet,outSheet,filRange,dayColumn,0,infoCol,lunchTog); // copy-paste all first kids
    if (singleSiblings > 0) {
      inputSheet.getFilter().setColumnFilterCriteria(infoCol[3], remBlanks); // filter the second kid name column
      copyPasteInfo(2,onSheet,inputSheet,outSheet,filRange,dayColumn,singleSiblings,infoCol,lunchTog); // copy-paste all second kids
      inputSheet.getFilter().removeColumnFilterCriteria(infoCol[3]);
    }
  }
  
  inputSheet.getFilter().removeColumnFilterCriteria(dayColumn); // remove single day filter
  
  // next, copypaste for full week kids
  if (weeksTog === 'Yes') {
    inputSheet.getFilter().setColumnFilterCriteria(weekColumn, remBlanks);
    
    if (fieldOffset > 0) { // if there's field trip option on that day, set additional field trip column filter.
      inputSheet.getFilter().setColumnFilterCriteria(infoCol[9], fieldFilter);
    }
    
    copyPasteInfo(1,onSheet,inputSheet,outSheet,filRange,weekColumn,fdKids-fdSiblings,infoCol,lunchTog); // copy-paste all first kids
    if (fdSiblings > 0) {
      inputSheet.getFilter().setColumnFilterCriteria(infoCol[3], remBlanks); // filter the second kid name column
      copyPasteInfo(2,onSheet,inputSheet,outSheet,filRange,weekColumn,fdSiblings,infoCol,lunchTog); // copy-paste all second kids
      inputSheet.getFilter().removeColumnFilterCriteria(infoCol[3]);
    }
    
    // remove any applied filters
    if (fieldOffset > 0) {
      inputSheet.getFilter().removeColumnFilterCriteria(infoCol[9]);
    }
    inputSheet.getFilter().removeColumnFilterCriteria(weekColumn);
  }
  
  var newArray = outSheet.getRange('G:G').getValues(); // need to change day type
  var campTypeVal = replTranspose(newArray); // transpose while truncating string values
  campTypeVal = transpose(campTypeVal); // re-transpose back to correct form
  outSheet.getRange('G:G').setValues(campTypeVal); // set new day-type values in
  
  newArray = outSheet.getRange('H:H').getValues(); // need to change lunch answer
  var lunchTrunc = replTranspose(newArray); // transpose while truncating string values
  lunchTrunc = transpose(lunchTrunc); // re-transpose back to correct form
  outSheet.getRange('H:H').setValues(lunchTrunc); // set new lunch values in
  
  outSheet.sort(3, true);
  outSheet.getRange('A2').setValue(1); // fix the numbering in left-most column
  outSheet.getRange('A3').setValue(2);
  outSheet.getRange('A2:A3').activate();
  lastRow = outSheet.getMaxRows();
  onSheet.getActiveRange().autoFill(onSheet.getRange('A2:A'.concat(lastRow)), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); // auto-fill code
  
  // final clean-up
  
  // hide age and team columns if team toggle is off, and also hide lunch column if toggle is off
  if (teams === "No") {
    outSheet.hideColumns(4,2);
  }
  else {
    outSheet.insertRowsAfter(lastRow, 4);
    outSheet.getRange(lastRow+2,5).setValue("=COUNTIF(E2:E".concat(lastRow,",\"R\")"));
    outSheet.getRange(lastRow+2,6).setValue('Red');
    outSheet.getRange(lastRow+3,5).setValue("=COUNTIF(E2:E".concat(lastRow,",\"Y\")"));
    outSheet.getRange(lastRow+3,6).setValue('Yellow');
    outSheet.getRange(lastRow+4,5).setValue("=COUNTIF(E2:E".concat(lastRow,",\"B\")"));
    outSheet.getRange(lastRow+4,6).setValue('Blue');
  }
  if (lunchTog === "No") {
    outSheet.hideColumns(8);
  }
  
  // remove location filter
  if (infoCol[8] !== 'x') {
    inputSheet.getFilter().removeColumnFilterCriteria(infoCol[8]);
  }
  
  // https://stackoverflow.com/questions/24894648/get-today-date-in-google-appscript <- can use Utilities.formatDate since I just want to output a string anyway
  // https://stackoverflow.com/questions/18596933/google-apps-script-formatdate-using-users-time-zone-instead-of-gmt <- does return a string for use in Utilities.formatDate
  runSheet.getRange('E5').setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "EEE MMM d, yyyy HH:mm:ss"));
  outSheet.getRange('A1').activate();
}

function findCorrectDay(date, theRange) {
  var dayValues = theRange.getValues();
  var dateLength = dayValues[0].length;
  var dayCheck;
  for (var i = 0; i < dateLength; i++) {
    dayCheck = dayValues[0][i];
    if (dayCheck.getTime() === date.getTime()) { // https://stackoverflow.com/questions/492994/compare-two-dates-with-javascript
      var foundCol = theRange.getColumn()+i;
      return foundCol;
    }
  }
}

function findCorrectWeek(week, leRange) { // similar to findCorrectDay
  var weekValues = leRange.getValues();
  var weekLength = weekValues[0].length;
  var weekCheck;
  for (var i = 0; i < weekLength; i++) {
    weekCheck = weekValues[0][i];
    if (weekCheck === week) {
      var founditCol = leRange.getColumn()+i;
      return founditCol;
    }
  }
}

function copyPasteInfo(which, spreadsheet, inSheet, outputSheet, filteredRange, campTypeCol, dummyRows, columnInd) {
  // for any pass of the function where dummyRows is not 0
  outputSheet.insertRowsBefore(2,dummyRows+1); // inserts appropriate # of rows at top. 1 if first batch of kids (dummy row), 1+#siblings if any additional kids (dummy row and new kids being added)
  if (dummyRows > 0) {
    spreadsheet.getActiveSheet().deleteRows(outputSheet.getMaxRows()-(dummyRows),dummyRows);
  }
  
  // for either first kid or second kid
  var nm;
  var ag;
  var aller;
  if (which === 1) {
    nm = 0;
    ag = 1;
    aller = 2;
  }
  else {
    nm = 3;
    ag = 4;
    aller = 5;
  }
  var names = inSheet.getRange(filteredRange.getRow()+1,columnInd[nm],filteredRange.getLastRow()-1,1);
  var kidAges = inSheet.getRange(filteredRange.getRow()+1,columnInd[ag],filteredRange.getLastRow()-1,1);
  var allergies = inSheet.getRange(filteredRange.getRow()+1,columnInd[aller],filteredRange.getLastRow()-1,1);
  
  // common to both first and second kid
  var phone = inSheet.getRange(filteredRange.getRow()+1,columnInd[6],filteredRange.getLastRow()-1,1);
  var campType = inSheet.getRange(filteredRange.getRow()+1,campTypeCol,filteredRange.getLastRow()-1,1); // campTypeCol already obtained previously
    
  // copy values to sheet
  names.copyTo(outputSheet.getRange('C2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  kidAges.copyTo(outputSheet.getRange('D2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  phone.copyTo(outputSheet.getRange('F2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  campType.copyTo(outputSheet.getRange('G2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  allergies.copyTo(outputSheet.getRange('N2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  // lunch is done last
  if (columnInd[7] !== 'x') {
    var lunches = inSheet.getRange(filteredRange.getRow()+1,columnInd[7],filteredRange.getLastRow()-1,1);
    lunches.copyTo(outputSheet.getRange('H2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }
  
  // final clean-up of this step
  spreadsheet.getActiveSheet().deleteRow(2); // remove the dummy top row
}

function replTranspose(a) { // i combined two search results - one to transpose and one to replace array values
  return a[0].map(function(col, i){ // transpose (use map solution) https://stackoverflow.com/questions/17428587/transposing-a-2d-array-in-javascript
    return a.map(function(row){ // replace text, works best at row level (which is what we have here) https://stackoverflow.com/questions/26480857/how-do-i-replace-text-in-a-spreadsheet-with-google-apps-script
      var mod = row[i].toString();
      var temp = mod.split(" - ");
      mod = temp[0];
      mod = mod.replace(/Full Day/i,'-');
      mod = mod.replace(/ Only/i,'');
      mod = mod.replace(/Yes/i,'Y');
      mod = mod.replace(/No/i,'');
      return mod;
    });
  });
}

function transpose(a) { // additional help for map: https://www.w3schools.com/jsref/jsref_map.asp
  return a[0].map(function(col, i){
    return a.map(function(row){
        return row[i];
    });
  });
}
