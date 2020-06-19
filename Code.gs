var mycal = 'abc@abc.com'; // Your Google Calendar ID
var d = new Date();var year = d.getFullYear();var month = d.getMonth();var day = d.getDate();
var sDate = new Date(year, month - 12, day); //Export Strat Date (Today - 12months)
var eDate = new Date(year, month + 12, day); //Export End Date (Today + 12months)

function onOpen() {
  
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
  .createMenu('Export Calendar')
  .addItem('All Schedule (Once)', 'showAlert')
  .addItem('Sort by StartDate', 'sortingByStartDate')
  .addItem('Sort by Last Updated', 'sortingByLastUpdate') 
  .addItem('Sort by User', 'sortingByUser')      
  .addToUi();
}

function showAlert() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var result = ui.alert(
    'Export All Schedule',
    'The existing data is initialized. continue?',
    ui.ButtonSet.YES_NO);
  
  // Process the user's response.
  if (result == ui.Button.YES) {
    first_once_all_export();
  } else {
    ui.alert('Canceled');
  }
}


function first_once_all_export(){ //After the sheet is initialized, set the header and get all calendar contents (first time only)
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clearContents(); //Sheet Initialization
  
  var cal = CalendarApp.getCalendarById(mycal);
  var events = cal.getEvents(sDate, eDate);
  
  
  
  // Header settings 
  var header = [["Event ID", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event"]]
  var range = sheet.getRange(1,1,1,14);
  range.setValues(header);
  
  
  for (var i=0;i<events.length;i++) {
    var row=sheet.getLastRow()+1;
    var myformula_placeholder = '';
    // Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
    // NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
    var details=[[events[i].getId(), events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, ('' + events[i].getVisibility()), events[i].getDateCreated(), events[i].getLastUpdated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent()]];
    var range=sheet.getRange(row,1,1,14);
    range.setValues(details);
    
    // Writing formulas from scripts requires that you write the formulas separate from non-formulas
    // Write the formula out for this specific row in column 7 to match the position of the field myformula_placeholder from above: foumula over columns F-E for time calc
    var cell=sheet.getRange(row,7);
    cell.setFormula('=(HOUR(F' +row+ ')+(MINUTE(F' +row+ ')/60))-(HOUR(E' +row+ ')+(MINUTE(E' +row+ ')/60))');
    cell.setNumberFormat('.00');
    
  }
  sortingByLastUpdate ();
}

function export_gcal_to_gsheet(){ //Get updated schedules through real-time triggers
  
  if (sortingByLastUpdate () == true) {
    
    var sheet = SpreadsheetApp.getActiveSheet();
    var cal = CalendarApp.getCalendarById(mycal);
    var events = cal.getEvents(sDate, eDate);
    
    // Get Standard time from sheet (last updated time)
    var stdDate;
    if (sheet.getLastRow() <= 1) {
      
      // Header settings  
      var header = [["Event ID", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event"]]
      var range = sheet.getRange(1,1,1,14);
      range.setValues(header);
      
      
      stdDate = new Date();
      stdDate = transformDate(stdDate);
    } else {
      stdDate = sheet.getRange(sheet.getLastRow(),10).getValues()[0];
      stdDate = transformDate(stdDate[0]);
    }
        
 
    Logger.log("Standard Date : " +stdDate);
    
    // var filteredEvents = events.filter(function(e){return e.getLastUpdated() == stdDate[0]});
    
    for (var i=0;i<events.length;i++) {
      
      if(transformDate(events[i].getLastUpdated()) > stdDate){ // Update only events larger than the standard time
        Logger.log("Greater than the standard");   
        Logger.log(i + " : " + events[i].getTitle() + " | " + transformDate(events[i].getLastUpdated()));
        
        var row=sheet.getLastRow()+1;
        var myformula_placeholder = '';
        
        var details=[[events[i].getId(), events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, ('' + events[i].getVisibility()), events[i].getDateCreated(), events[i].getLastUpdated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent()]];
        var range=sheet.getRange(row,1,1,14);
        range.setValues(details);
        
        var cell=sheet.getRange(row,7);
        cell.setFormula('=(HOUR(F' +row+ ')+(MINUTE(F' +row+ ')/60))-(HOUR(E' +row+ ')+(MINUTE(E' +row+ ')/60))');
        cell.setNumberFormat('.00');
        
      }
    }
  }
  sortingByLastUpdate ();
  
}

function transformDate (e){ //Time Formatting
  return Utilities.formatDate(e, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
}


function sortingByLastUpdate () {  //Sort by last update
  var sheet = SpreadsheetApp.getActiveSheet();
  var last = sheet.getLastRow();
  if (last > 1) {
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  range.sort({column: 10});
  }
  return true;
}

function sortingByStartDate () {  //Sort by start date
  var sheet = SpreadsheetApp.getActiveSheet();
  var last = sheet.getLastRow();
  if (last > 1) {  
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  range.sort({column: 5});
  }
}

function sortingByUser () {  //Sort by user
  var sheet = SpreadsheetApp.getActiveSheet();
  var last = sheet.getLastRow();
  if (last > 1) {  
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  range.sort({column: 12});
  }
}
