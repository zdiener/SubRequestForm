/** 
 *   This is a Google Script file designed to work with a specific Google Spreadsheet.
 *   This script enables a form to populate a list of teachers and subs for different classes, and to send an email to those teachers.
 *
 * @OnlyCurrentDoc
 */
 
 
function copyFormatting() {
  var ss = SpreadsheetApp.getActive();
  var sub_sheet = ss.getSheetByName('Sub Request');
  var format_sheet = ss.getSheetByName('format');
  
  format_sheet.getRange(1, 1, 23, 11).copyTo(sub_sheet.getRange( 1, 1, 23, 11));
}

function onEdit (e) {
  checkClasses();
}

function getSubsSS() {
  var ss = SpreadsheetApp.getActive();
//  Using getActive() may introduce bugs fixed by using open By Id(), but opeById() may require extra permissions.
  return ss;
}

function getSub_Request() {
  var ss = getSubsSS();
  var sheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName('Sub Request'));
  return sheet;
}

function getClasses() {
  var classes = [];
  var ss = getSubsSS();
  var sheets = ss.getSheets();
  for (var i = 0; sheets.length > i; i++) {
    if (sheets[i].getName() == 'Sub Request' || sheets[i].getName() == 'format') {continue}
    classes.push((sheets[i]).getName());
  }
  return classes; 
}

function isEmpty(arg) {
  return arg == '';
}

function hidePhoneNumbers() {
  var sheet = getSub_Request();
  sheet.hideColumn(sheet.getRange('C:C'))
}

function showPhoneNumbers() {
  var sheet = getSub_Request();
  sheet.unhideColumn(sheet.getRange('C:C'))
}

function removeDuplicates(teachers, title) { // title determines if range is 'teachers' or 'subs'
  // Get teachers and subs. If a teacher was already listed from a previous class, don't list the duplicate.
  if (title != 'teachers' && title != 'subs') { Logger.log('removeDuplicates did not specify Teacher or Sub'); return 0 };
  var ss = getSubsSS();
  var teachers_range = ss.getRangeByName(title);
  var data = teachers_range.getValues();
  data = data.concat(teachers);
  var newData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    if (!row[2]) { continue };
    for(j in newData){
      if(row[2] == newData[j][2]){
        duplicate = true;
      }
    }
    if(!duplicate){ //remove empty rows here too
      newData.push(row);
    }
  }
  ss.setNamedRange(title, ss.getSheetByName('Sub Request').getRange(teachers_range.getRow(), teachers_range.getColumn(), newData.length, newData[0].length));
  ss.getRangeByName(title).setValues(newData).setHorizontalAlignment('left').setBorder(false, false, false, false, true, false, '#D9D9D9', null);
}

function setClasses(target_cell){
  // Set the data validation for target cell to be a dropdown of the sheet names
  var ss = getSubsSS();
  var sheet = getSub_Request();
  var cell = ss.getSheetByName('Sub Request').getRange(target_cell);
  var classes = getClasses();
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(classes, true).build();

  if ( ss.getRangeByName('classes').getNumRows() > 3 ) {
    ss.getSheetByName('Sub Request').insertRowAfter(ss.getRangeByName('classes').getNumRows() + 5);
  }

  ss.getSheetByName('Sub Request').getRange('B6').copyFormatToRange(ss.getSheetByName('Sub Request'), cell.getColumn(), cell.getColumn(), cell.getRow(), cell.getRow());
  cell.setDataValidation(rule);
 
}

function updateMessage() {
  var ss = getSubsSS();
  var classes = ss.getRangeByName('classes').getValues();
  var days_times = ss.getRangeByName('days').getDisplayValues();
  var class_message = 'I need a sub for the following classes:\n';
  var new_msg = '';
  var user = Session.getActiveUser().getEmail()
  for (i in classes) {
    if (classes[i][0].toString() == '') { continue };
    
    new_msg = new_msg + '\n** ' + classes[i][0].toString() + ' **\n';
    
    if (!days_times[i][0].toString() == '') {
      new_msg = new_msg + days_times[i][0].toString();
    }
    if(!days_times[i][1].toString() == '') {
      new_msg = new_msg + ' at ' + days_times[i][1].toString();
    }
    
    new_msg = new_msg + '\n';
  }
  class_message = class_message + new_msg;
  class_message = class_message + '\nThanks!\n' + user;
  ss.getRangeByName('message').setValue(class_message);
}

function checkClasses() {
  Logger.log('checkClasses');
  // Check has a class has been picked
  // Place a new class dropdown under the last filled cell
  var ss = getSubsSS();
  
  var class_range = ss.getRangeByName('classes');
  var check_range = ss.getRangeByName('class_check');
  
  var chosen_classes = class_range.getValues();
  var check_classes = check_range.getValues();
  
  var first = class_range.isBlank();
  var second = Boolean(chosen_classes.toString() == check_classes.toString());

  if ( class_range.isBlank() ) { 
      return 0;
  } else if ( chosen_classes.toString() == check_classes.toString() ) {
      var days = ss.getRangeByName('days').getValues();
      var checkDays = ss.getRangeByName('days_check').getValues();
      
      if ( days.join() != checkDays.join() ){
        updateMessage();
      };
  } else {
      updateMessage();
      for (i = 0; i < class_range.getNumRows() ; i++) {
        source_class = class_range.getCell(i+1, 1).getDisplayValue();
        class_check = check_range.getCell(i+1, 1).getDisplayValue();
        
        if (source_class != class_check) {
          check_range.getCell(i+1, 1).setValue(source_class);
          getTeachers(class_range.getCell(i+1, 1).getA1Notation());
        }
      }
      var rows = class_range.getNumRows()+1;
      ss.setNamedRange('classes', class_range.offset(0, 0, rows));
      ss.setNamedRange('class_check', check_range.offset(0, 0, rows));
      ss.setNamedRange('days', ss.getRangeByName('days').offset(0, 0, rows));
      ss.setNamedRange('days_check', ss.getRangeByName('days_check').offset(0, 0, rows));
      
      var format_arr = [];
      while ( format_arr.push(['Ddd". "mmmm" "d" "','@']) < rows );
      ss.getRangeByName('days').setNumberFormats(format_arr);
      var sheet = getSub_Request();
      
      class_range = ss.getRangeByName('classes');
      var cell = class_range.getCell(class_range.getNumRows(), 1);
      setClasses(cell.getA1Notation()); 
      return 0;
  }
}

function getTeachers(cell) {
  var ss = getSubsSS();
  var form_sheet = getSub_Request();
  var teachers = new Array();
  var subs = new Array();
  
  //Get the chosen class. If no cell specified get the first in the list.
  if (!cell) {cell = 'B6';};
  var class = form_sheet.getRange(cell).getDisplayValue();
  
  //Store the chosen class' sheet as class_sheet
  var class_sheet = ss.getSheetByName(class);
  
  //Grab the class sheets range of teachers. 
  var rows = class_sheet.getLastRow();
  rows = rows-2; //Don't count the header rows
  if (rows < 1) { 
    SpreadsheetApp.getUi().alert('No subs found for ' + class + '.\n If you know of subs, please add them to the class sub list. You can find the sub lists in the tabs at the bottom of the page. ');
    return teachers };
 
 var title = 'teachers';
 /*
  var source_range = class_sheet.getRange(3,1,rows,3);
  teachers = source_range.getValues();
  source_range = class_sheet.getRange(3,5,rows,3);
  teachers = teachers.concat(source_range.getValues());
  removeDuplicates(teachers, title);
  */
  
  teachers = class_sheet.getRange(3,1,rows,3).getValues();
  removeDuplicates(teachers, title);
  teachers = removeNulls(teachers);
  title = 'subs';
  subs = class_sheet.getRange(3,5,rows,3).getValues();
  removeDuplicates(subs, title);
  
  return teachers;
}

function resetSheet() {
  var ss = getSubsSS();
  var sheet = getSub_Request();
  if (ss.getRangeByName('classes')) { 
    var classes = ss.getRangeByName('classes').getA1Notation()
    sheet.getRange(classes).clear().clearDataValidations();
  }
  sheet.clear();
  copyFormatting();
  /*
  sheet.getRange('B2:E2').setValues([
    ['What Classes do you need a sub for?','','What days?','What Times?']
  ]);
  sheet.getRange('F2:H4').setValues([['Subject:','','I need a sub!'],['','',''],['Message:','','']]);
  sheet.getRange('B8:D9').setValues([['Teachers','',''],['Name','Phone','Email']]);
  */
  ss.setNamedRange('classes', sheet.getRange('B6'));
  ss.setNamedRange('class_check', sheet.getRange('AA1'));
  ss.setNamedRange('teachers', sheet.getRange('B13:D13'));
  ss.setNamedRange('subs', sheet.getRange('B16:D16'));
  ss.setNamedRange('subject', sheet.getRange('H5'));
  ss.setNamedRange('message', sheet.getRange('H7:J23'));
  ss.setNamedRange('days', sheet.getRange('D6:E6'));
  ss.setNamedRange('days_check', sheet.getRange('AB1:AC1'));
  
  sheet.getRange(1,8,100).setNumberFormat('@STRING@');

  ss.getRangeByName('message').merge().setVerticalAlignment('top').setValue(
    "I need a sub for the following classes:\n\n\n\nThanks!");
  setClasses('B6');
}

function removeNulls(a) {
  var b = [];
  for (var i = 0; i < a.length; i++) {
    if (a[i][0] != "" && a[i][1] != "" && a[i][2] != "")
    b.push(a[i]);
  }
}
    

function filt(a) { 
 var b = []; 
 for(var i = 0;i < a.length;i++) { 
  if (a[i][0] !== undefined && a[i][0] != null) { 
   b.push(a[i][0]); 
  }
 } 
 var a = [];
 for(var i = 0;i < b.length;i++) { 
  if (b[i] != undefined && b[i] != "") { 
   a.push(b[i]); 
  }
 } 
 return a; 
}

function emailTeachers() {
  var ss = SpreadsheetApp.getActive();
  var sub_request = getSub_Request();
  var teachers_range = ss.getRangeByName('teachers');
  var teachers = sub_request.getRange(teachers_range.getRow(), teachers_range.getLastColumn(), teachers_range.getNumRows()).getValues();
  var ui = SpreadsheetApp.getUi();
  
  teachers = filt(teachers);
  teachers = teachers.join(', ');
  
  var result = ui.alert(
     'Please confirm',
     'Email '+ teachers +' about subbing your class?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    var subject = ss.getRangeByName('subject').getValue();
    var message = ss.getRangeByName('message').getValue();
    MailApp.sendEmail(teachers, subject, message);
    // MailApp.sendEmail('Zac.diener@gmail.com', 'This is the subject', teachers);
    ui.alert('Email sent.');
    resetSheet();
  } else {
    // User clicked "No" or X in the title bar.
    return 0;
  }
  
  return teachers;
}
