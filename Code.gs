/** PLEASE UPDATE THESE 2 VARS WITH YOUR SHARED CALENDAR ID and SHEET ID THAT YOU CREATE **/
/** Please see Readme.md for more details. **/
var personalCalendarID = "181h7gtr00dlj5t35d028dah4s@group.calendar.google.com"
var sheetID = "1U-qaKx8tfnOM3DyCOuGEP6YeMbqKBPtX4NH8LC6Q0hk"


function clearPersonalCalendar() {
  var today = new Date()
  var oneWeek = new Date()
  oneWeek.setDate(today.getDate() + 7)
  var lastWeek = new Date()
  lastWeek.setDate(today.getDate() - 7)
  var personalCalendar = CalendarApp.getCalendarById(personalCalendarID)
  var personalEvents = personalCalendar.getEvents(lastWeek, oneWeek)
  personalEvents.forEach(function(event) {event.deleteEvent()})
  var sheet = SpreadsheetApp.openById(sheetID)
  // acquire data from spreadsheet
  var range = sheet.getDataRange()
  var rows = range.getValues()
  if (rows.length > 0) {
    sheet.deleteRows(1, rows.length)
    sheet.appendRow(["gSuite ID", "Personal ID", "Title", "Start", "End", "Status", "All Day"])
  }
}

function logGSuiteEventsToSync() {
  var today = new Date()
  var oneWeek = new Date()
  oneWeek.setDate(today.getDate() + 7)
  var lastWeek = new Date()
  lastWeek.setDate(today.getDate() - 7)

  var gSuiteCalendar = CalendarApp.getDefaultCalendar()
  var gSuiteEvents = gSuiteCalendar.getEvents(today, oneWeek)

  gSuiteEvents.forEach(function(event){
    Logger.log("[\n ID:"+event.getId()+"\n Title:"+event.getTitle()+"\n Start:"+event.getStartTime()+"\n End:"+event.getEndTime()+"\n]")
  })  
}


function syncCalendars() {
  // track when this script was executed
  logTriggerStart()

  // we'll be looking at syncing 3 weeks worth of event (-1 to +2 weeks)
  var today = new Date()
  var endDate = new Date()
  endDate.setDate(today.getDate() + 14)
  var beginDate = new Date()
  beginDate.setDate(today.getDate() - 7)

  // acquire a reference to your personal google account calendar
  var personalCalendar = CalendarApp.getCalendarById(personalCalendarID)
  var personalEvents = personalCalendar.getEvents(beginDate, endDate)

  // acquire a reference to your default calendar (which will be relative to the account this script executes under)
  // note: this script should be executed within your g suite account for this lookup to work as expected
  var gSuiteCalendar = CalendarApp.getDefaultCalendar()
  var gSuiteEvents = gSuiteCalendar.getEvents(beginDate, endDate)

  var gSuiteEventsLookup = {}
  gSuiteEvents.forEach(function(event){
    gSuiteEventsLookup[computeId(event)] = event
  })
  
  var personalEventsLookup = {}
  personalEvents.forEach(function(event) {
    personalEventsLookup[event.getId()] = event
  })

  // id of spreadsheet (needed to track calendar events)
  var sheetID = "1U-qaKx8tfnOM3DyCOuGEP6YeMbqKBPtX4NH8LC6Q0hk"
  var sheet = SpreadsheetApp.openById(sheetID)
  // acquire data from spreadsheet
  var range = sheet.getDataRange()
  var rows = range.getValues()

  var linkRecordsByGSuiteID = {}
  var linkRecordsByPersonalID = {}
  rows.forEach(function(row, index){
    if (index == 0) return; // Skip the Header.
    
    var linkRecord = readRecord(row)
    linkRecordsByGSuiteID[linkRecord.gSuiteID] = linkRecord;
    linkRecordsByPersonalID[linkRecord.personalID] = linkRecord;
  })
  
  gSuiteEvents.forEach(function(event){
    var linkRecord = linkRecordsByGSuiteID[computeId(event)];
    var title = event.getTitle()
    var allDay = event.isAllDayEvent()
    var start = new Date(event.getStartTime())
    var end = new Date(event.getEndTime())
    var status = event.getMyStatus()
    var description = event.getDescription()
    var location = event.getLocation()
    if (status == 'INVITED') {
      title = "INVITED: "+ title
    } else if (status == 'YES') {
      title = "ACCEPTED: "+ title
    } else if (status == 'NO') {
      title = "DECLINED: "+ title
    } else if (status == 'MAYBE') {
      title = "TENTATIVE: "+ title
    }
    if (linkRecord) {
      // updateEvent(event, personalEventsLookup[linkRecord.personalID]);
      if (linkRecord.title != title || linkRecord.start != start || linkRecord.end != end || linkRecord.allDay != allDay || linkRecord.location != location || linkRecord.description != description) {
        var personalEvent = personalCalendar.getEventById(linkRecord.personalID)
        if (linkRecord.title != title) personalEvent.setTitle(title)
        if (linkRecord.start != start || linkRecord.end != end || linkRecord.allDay != allDay) {
          if (allDay) {
            personalEvent.setAllDayDates(start, end)
          } else {
            personalEvent.setTime(start, end)
          }
        }
        if (linkRecord.description != description) personalEvent.setDescription(description)
        if (linkRecord.location != location) personalEvent.setLocation(location)
        updateRecord({"gSuiteID": linkRecord.gSuiteID, "personalID": linkRecord.personalID, "title": title, "start": start, "end": end, "status": status, "allDay": allDay, "location": location, "description": description}, sheet)
      }
    } else if (status != 'NO') {
      var newPersonalEvent = personalCalendar.createEvent(title, start, end, {description: description, location: location})
      updateRecord({"gSuiteID": computeId(event), "personalID": newPersonalEvent.getId(), "title": title, "start": start, "end": end, "status": status, "allDay": allDay, "location": location, "description": description}, sheet)
    }
  })
  
  var lastRow = rows.length - 1; // Last Row from initial Set of Rows (as rows still points to original data range.)
  for (var i=lastRow; i>0; i--) {
    var row = rows[i];
    var linkRecord = readRecord(row)
    var event = gSuiteEventsLookup[linkRecord.gSuiteID]
    if (!event || event.getMyStatus() == 'NO') {
      // DeleteEvent
      try {
        personalCalendar.getEventById(linkRecord.personalID).deleteEvent()
      } catch (e) {
        // Do nothing. Event might have been deleted by previous run. If not, the next loop will catch the missed ones.
      }
      // DeleteRow
      sheet.deleteRow(i+1) // Rows indexed from 1 :(
    }
  }
  
  personalEvents.forEach(function(event){
    var linkRecord = linkRecordsByPersonalID[event.getId()];
    if (!linkRecord) {
      try {
        event.deleteEvent();
      } catch (e) {
        // Do nothing. Might've been deleted in the previous step.
      }
    }
  })  
}

function readRecord(row) {
  return { gSuiteID: row[0], personalID: row[1], title: row[2], start: row[3], end: row[4], status: row[5], allDay: row[6], location: row[7], description: row[8] }
}

function computeId(event) {
  if (event.isRecurringEvent()) {
    return "" + event.getId() + "-" + event.getStartTime();
  }
  return event.getId();
}

function updateRecord(linkRecord, sheet) {
  var recordUpdated = false;
  var rows = sheet.getDataRange().getValues();
  for(index = 1; index < rows.length; index++) { // Starting from 1 to skip header row
    var row = rows[index];
    record = readRecord(row);
    if (record.gSuiteID == linkRecord.gSuiteID) {
      var rangeForCurrentRecord = sheet.getRange("B" + (index+1) + ":I" + (index+1))
      rangeForCurrentRecord.setValues([[linkRecord.personalID, linkRecord.title, linkRecord.start, linkRecord.end, linkRecord.status, linkRecord.allDay, linkRecord.location, linkRecord.description]])
      recordUpdated = true;
      break;
    }
  }
  if (!recordUpdated) {
    sheet.appendRow([linkRecord.gSuiteID, linkRecord.personalID, linkRecord.title, linkRecord.start, linkRecord.end, linkRecord.status, linkRecord.allDay, linkRecord.location, linkRecord.description])
  }
}

function logTriggerStart() {
  var d = new Date()
  var hour = d.getHours().toString()
  var minute = d.getMinutes().toString()

  Logger.log("Event has been triggered: %s:%s", hour, minute)
}

function logEvents(e) {
  Logger.log(e)
  Logger.log("\nID: %s\nTitle: %s\nStart: %s\n End: %s", e.getId(), e.getTitle(), e.getStartTime(), e.getEndTime())
}

function createSpreadsheet(companyName) {
  var sheet = SpreadsheetApp.create(companyName + " Sync")
  sheet.appendRow(["gSuite ID", "Personal ID", "Title", "Start", "End", "Status", "All Day", "Location", "Description"])
  Logger.log(sheet.getId())
}
