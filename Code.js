function mainFunction() {
  
  // ******************************** 1. PREPARE THE FILE *********************************
  // 1.a Create a duplicate of last week’s report. Rename it with today’s date.

  // generates the timestamp and stores in variable formattedDate as year-month-date hour-minute-second
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd' 'HH:mm:ss");
  
  // gets the current month
  var newMonth = formattedDate.slice(5,7);

  // gets the current day
  var newDay = formattedDate.slice(8,10);
  
  // gets the current year
  var newYear = formattedDate.slice(0,4);
  
  // gets the name of the original file
  var name = SpreadsheetApp.getActiveSpreadsheet().getName();
  
  // sets the appropriate name for the new report
  var newName = name.slice(0, 27) + " " + "-" + " " + newMonth + "/" + newDay + "/" + newYear;
  
  // gets the destination folder of the Weekly Reports
  var destination = DriveApp.getFolderById("");
  
  // gets the current Google Sheet file
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());

  // makes copy of "file" with "name" at the "destination"
  var newReport = file.makeCopy(newName, destination);
  
  // ******************************** 1. PREPARE THE FILE *********************************
  // 1.b For the “Performance Trend” sheet - copy the values of column B, insert one column to the right of column B and hard paste. Rename the week for column B.
  
  // gets the Performance Trend sheet and stores it into performanceTrendSheet
  var performanceTrendSheet = SpreadsheetApp.open(newReport).getSheetByName("Performance Trend");
  
  // inserts a new column in Performance Trend sheet to the left of column B
  performanceTrendSheet.insertColumnBefore(2);
  
  // sets the B1 cell to the current date
  performanceTrendSheet.getRange(1, 2).setValue(`Week ${newMonth}/${newDay}`);
  
  // ******************************** 1. PREPARE THE FILE *********************************
  // 1.c Copy out the 5 “Other Jira Tickets Closed” columns on the Weekly Performance sheet to the other side of the blank column(Make sure you use Ctrl+Shift+V). Update the week range in the title.

  // gets the Weekly Performance sheet and stores it into weeklyPerformanceSheet
  var weeklyPerformanceSheet = SpreadsheetApp.open(newReport).getSheetByName("Weekly Performance");
  
  // inserts 5 new columns to the left of column T
  weeklyPerformanceSheet.insertColumnsBefore(20, 5);
  
  // selects the newly inserted columns
  var newRange = weeklyPerformanceSheet.getRange(4, 20, 60, 5);

  // selects last week's Other Jira Tickets Closed
  var rangeToCopy = weeklyPerformanceSheet.getRange(4, 25, 60, 5);
  var valuesToCopy = rangeToCopy.getValues();
  
  // copies last week's Other Jira Tickets Closed into the newly inserted columns
  newRange.setValues(valuesToCopy);
  
  // merges the cells for where the title goes
  weeklyPerformanceSheet.getRange('T2:X3').merge();
  
  //Change it so that it is 7 days in the past.
  var pastDate = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  var newFormattedDate = Utilities.formatDate(new Date(pastDate), "GMT", "yyyy-MM-dd' 'HH:mm:ss");
  
  // get the past month
  var pastMonth = newFormattedDate.slice(5,7);
  
  //get the past day
  var pastDay = newFormattedDate.slice(8,10);
  
  // Inserts week number into title
  weeklyPerformanceSheet.getRange('T2:X3').setValue(`Other Jira Tickets Closed [Week ${pastMonth}/${pastDay} - ${newMonth}/${newDay}]`);

  // Inserts Heading "Jira Queue"
  weeklyPerformanceSheet.getRange('V1:V1').setValue("Jira Queue");
  
  // ******************************** 1. PREPARE THE FILE *********************************
  // 1.d Delete from the second row down everything on the “Completed BigRocks - Ended Sprint”, “Other BigRocks Completed Last Week”, “Other Tickets Completed”, “Planned BigRocks” & “Planned BigRocks SS”
  
  // gets the Completed BigRocks - Ended Sprint sheet and stores it into completedBigRocks
  var completedBigRocks = SpreadsheetApp.open(newReport).getSheetByName("Completed BigRocks - Ended Sprint");
  
  // deletes all rows in completedBigRocks starting with the 2nd row
  completedBigRocks.deleteRows(2, completedBigRocks.getLastRow());
  
  // gets the Other BigRocks Completed Last Week sheet and stores it into otherCompletedBigRocks
  var otherCompletedBigRocks = SpreadsheetApp.open(newReport).getSheetByName("Other BigRocks Completed Last Week");

  // deletes all rows in otherCompletedBigRocks starting with the 2nd row
  otherCompletedBigRocks.deleteRows(2, otherCompletedBigRocks.getLastRow());
  
  // gets the Other Tickets Completed sheet and stores it into otherTicketsCompleted
  var otherTicketsCompleted = SpreadsheetApp.open(newReport).getSheetByName("Other Tickets Completed");

  // deletes all rows in otherTicketsCompleted starting with the 2nd row
  otherTicketsCompleted.deleteRows(2, otherTicketsCompleted.getLastRow());
  
  // gets the Planned BigRocks sheet and stores it into plannedBigRocks
  var plannedBigRocks = SpreadsheetApp.open(newReport).getSheetByName("Planned BigRocks");
  
  //deletes all rows in plannedBigRocks starting with the 2nd row
  plannedBigRocks.deleteRows(2, plannedBigRocks.getLastRow());
  
  // gets the Planned BigRocks SS sheet and stores it into plannedBigRocksSS
  var plannedBigRocksSS = SpreadsheetApp.open(newReport).getSheetByName("Planned BigRocks SS");
  
  // stores previous sprint into previousSprint
  var previousSprint = plannedBigRocksSS.getRange('Q2:Q2').getValue().slice(8, 10);
  
  // deletes all rows in plannedBigRocksSS starting with the 2nd row
  plannedBigRocksSS.deleteRows(2, plannedBigRocksSS.getLastRow());
  
  // ******************************** 1. PREPARE THE FILE *********************************
  // 1.e Move all the entries in the “BigRocks Ending This Week” to “Incomplete BigRocks”. Move all the entries in the “New BigRocks Starting This Week” to “BigRocks ending this week”. 
  //     Sort the Incomplete BigRocks sheet alphabetically using the “Owner” column.
  
  // gets the BigRocks Ending This Week sheet and stores it into bigRocksEndingThisWeek
  var bigRocksEndingThisWeek = SpreadsheetApp.open(newReport).getSheetByName("BigRocks Ending This Week");
  
  // gets the Incomplete BigRocks sheet and stores it into incompleteBigRocks
  var incompleteBigRocks = SpreadsheetApp.open(newReport).getSheetByName("Incomplete BigRocks");
  
  // gets the New BigRocks Starting This Week sheet and stores it into newBigRocksStartingThisWeek
  var newBigRocksStartingThisWeek = SpreadsheetApp.open(newReport).getSheetByName("New BigRocks Starting This Week");
  
  // makes multiple selects and variations of the first 200 row entries in BigRocks Ending This Week
  var rangeBigRocksEndingThisWeek = bigRocksEndingThisWeek.getRange(2, 1, 200, 13);
  var valuesRangeBigRocksEndingThisWeek = rangeBigRocksEndingThisWeek.getValues();
  
  var rangeFirstSplitBigRocksEndingThisWeek = bigRocksEndingThisWeek.getRange(2, 1, 200, 3);
  var valuesRangeFirstSplitBigRocksEndingThisWeek = rangeFirstSplitBigRocksEndingThisWeek.getValues();
  
  var rangeSecondSplitBigRocksEndingThisWeek = bigRocksEndingThisWeek.getRange(2, 4, 200, 10);
  var valuesRangeSecondSplitBigRocksEndingThisWeek = rangeSecondSplitBigRocksEndingThisWeek.getValues();
  
  // selects the first 200 row entries in Incomplete BigRocks
  var rangeIncompleteBigRocks = incompleteBigRocks.getRange(2, 1, 200, 13);
  var valuesRangeIncompleteBigRocks = rangeIncompleteBigRocks.getValues();
  
  // selects the first 3 columns and 200 rows in New BigRocks Starting This Week
  var rangeFirstSplitNewBigRocksStartingThisWeek = newBigRocksStartingThisWeek.getRange(2, 1, 200, 3);
  var valuesRangeFirstSplitNewBigRocksStartingThisWeek = rangeFirstSplitNewBigRocksStartingThisWeek.getValues();
  
  // selects the remaining columns and 200 rows in New BigRocks Starting This Week
  var rangeSecondSplitNewBigRocksStartingThisWeek = newBigRocksStartingThisWeek.getRange(2, 4, 200, 10);
  var valuesRangeSecondSplitNewBigRocksStartingThisWeek = rangeSecondSplitNewBigRocksStartingThisWeek.getValues();
  
  // Move all the entries in the “BigRocks Ending This Week” to “Incomplete BigRocks”
  rangeIncompleteBigRocks.setValues(valuesRangeBigRocksEndingThisWeek);
  
  // Move all the entries in the “New BigRocks Starting This Week” to “BigRocks ending this week”.
  rangeFirstSplitBigRocksEndingThisWeek.setValues(valuesRangeFirstSplitNewBigRocksStartingThisWeek);
  rangeSecondSplitBigRocksEndingThisWeek.setValues(valuesRangeSecondSplitNewBigRocksStartingThisWeek);
  
  // Sort the Incomplete BigRocks sheet alphabetically using the “Owner” column.
  incompleteBigRocks.sort(1);

  // ******************************** 1. PREPARE THE FILE *********************************
  // 1.f Update the Owner, Estimated Completion Date, Estimate (in hours) & Next Steps columns of the “Incomplete BigRocks” & “BigRocks Ending this week” entries by checking the tracker.
    
  // gets the SWOT & BigRocks Spreadsheet
  // TODO ID value below needs to change to appropriate data source for go live
  var fileSwotBigRocks = DriveApp.getFileById('');
  
  // gets the SWOT & BigRocks sheet
  var swotBigRocks = SpreadsheetApp.open(fileSwotBigRocks).getSheetByName("SWOT & BigRocks");
  
  // gets the incomplete bigrocks issues range to update
  var issuesIncompleteBigRocks = incompleteBigRocks.getRange(2, 5, incompleteBigRocks.getLastRow(), 1);
  var valuesIssuesIncompleteBigRocks = issuesIncompleteBigRocks.getValues().toString().replace(/,+/g, ',').trim();
  
  // gets the bigrocks ending this week issues to update
  var issuesBigRocksEndingThisWeek = bigRocksEndingThisWeek.getRange(2, 5, bigRocksEndingThisWeek.getLastRow(), 1);
  var valuesIssuesBigRocksEndingThisWeek = issuesBigRocksEndingThisWeek.getValues().toString().replace(/,+/g, ',').trim();
  
  // Function that removes "," if found at the very beggining or end of the passed string
  function trimCommas(aString) {
    if (aString.startsWith(',')) {
      aString = aString.slice(1, aString.length - 1);
    }

    if (aString.endsWith(',')) {
      aString = aString.slice(0, -1);
    }
    return aString;
  }

  // trimming commas from valuesIssuesIncompleteBigRocks and valuesIssuesBigRocksEndingThisWeek
  valuesIssuesIncompleteBigRocks = trimCommas(valuesIssuesIncompleteBigRocks);
  valuesIssuesBigRocksEndingThisWeek = trimCommas(valuesIssuesBigRocksEndingThisWeek);
  
  // sets the issues to update for the jiraPull function
  var rangeInstructions = SpreadsheetApp.open(newReport).getSheetByName("Instructions").getRange("B5:B5");
  rangeInstructions.setValue(valuesIssuesIncompleteBigRocks);
  
  var C_MAX_RESULTS = 250;
  
  var prefix = "";
  PropertiesService.getUserProperties().setProperty("prefix", prefix.toUpperCase());
  
  var host = "";
  PropertiesService.getUserProperties().setProperty("host", host);
  
  var userAndPassword = "";
  var x = Utilities.base64Encode(userAndPassword);
  PropertiesService.getUserProperties().setProperty("digest", "Basic " + x);
  
  var issueTypes = ""; 
  // "Story,Epic";
  PropertiesService.getUserProperties().setProperty("issueTypes", issueTypes);

  
  // Function to return all the field definitions for the project in a key/value pair
  function getFields() {
    return JSON.parse(getDataForAPI("field"));
    
  }  
  
  
  // function to return all the story data - either from a list on the instruction sheet, 
  // otherwise all the non-resolved issues for the project are returned
  // See here for api documentation: https://developer.atlassian.com/cloud/jira/platform/rest/#api-api-2-search-get
  function getStories() {
    var allData = {issues:[]};
    var data = {startAt:0,maxResults:0,total:1};
    var startAt = 0;
    var jql = "search?jql=project%20%3D%20" + PropertiesService.getUserProperties().getProperty("prefix") + "%20and%20status%20!%3D%20resolved%20and%20type%20in%20("+ encodeURIComponent(getStoryTypes()) + ")%20order%20by%20rank%20&maxResults=" + C_MAX_RESULTS;
    var issues = SpreadsheetApp.open(newReport).getSheetByName("Instructions").getRange("B5:B5").getValue();
    if (issues != "") {
      var jql = "search?jql=key%20in%20%28"+ issues + "%29%20order%20by%20rank%20&maxResults=" + C_MAX_RESULTS;

    }  
    while (data.startAt + data.maxResults < data.total) {
      Logger.log("Making request for %s entries", C_MAX_RESULTS);
      data =  JSON.parse(getDataForAPI(jql + "&startAt=" + startAt));  
      allData.issues = allData.issues.concat(data.issues);
      startAt = data.startAt + data.maxResults;
    }  
    
    return allData;
  }   
  
  function getStoryTypes() {
    var types = PropertiesService.getUserProperties().getProperty("issueTypes");
    types = types.replace(/[\""]/g, '\\"')
    var allTypes = types.split(',');
    var newTypes = "";
    for (var i=0;i<allTypes.length;i++) {
      if (newTypes !="") {
        newTypes += ","
      }  
      newTypes += '"' + allTypes[i].trim() + '"';
    }  
    Logger.log(newTypes);
    return newTypes;
  }   
  
  // function that actually makes the http request
  function getDataForAPI(path) {
    var url = PropertiesService.getUserProperties().getProperty("host") + "/rest/api/2/" + path;
    var digestfull = PropertiesService.getUserProperties().getProperty("digest");
    
    var headers = { "Accept":"application/json", 
                "Content-Type":"application/json", 
                "method": "GET",
                 "headers": {"Authorization": digestfull},
                   "muteHttpExceptions": true
               };
    
    var resp = UrlFetchApp.fetch(url,headers );
    if (resp.getResponseCode() != 200) {
      Browser.msgBox("Error retrieving data for url " + url + ":" + resp.getContentText());
      return "";
    }  
    else {
      return resp.getContentText();
    }  
    
  } 
  
  // Main function  
  function jiraPull() {
    
    
    // Retrieve data using API
    var allFields = getAllFields();
    var data = getStories();  
    if (allFields === "" || data === "") {
      Browser.msgBox("Error pulling data from Jira - aborting now.");
      return;
    }  
    
    // Retrieve column headings from backlog sheet.
    var ss = SpreadsheetApp.open(newReport).getSheetByName("Backlog");
    var headings = ss.getRange(1, 1, 1, 26).getValues()[0];
    
    // Process the stories and extract the data that matches the column headings into an array
    var y = new Array();
    for (i=0;i<data.issues.length;i++) {
      var d=data.issues[i];
      y.push(getStory(d,headings,allFields));
    }  
    
    // Output the contents of the array into the spreadsheet by clearing existing rows and adding new ones
    ss = SpreadsheetApp.open(newReport).getSheetByName("Backlog");
    var last = ss.getLastRow();
    if (last >= 2) {
      ss.getRange(2, 1, ss.getLastRow()-1,ss.getLastColumn()).clearContent();  
    }  
    if (y.length > 0 && data.issues.length > 0 && y[0].length > 0) {
      ss.getRange(2, 1, data.issues.length,y[0].length).setValues(y);
    }
    
  }
  
  // Get array of field ids and names
  function getAllFields() {
    
    var theFields = getFields();
    var allFields = new Object();
    allFields.ids = new Array();
    allFields.names = new Array();
    
    for (var i = 0; i < theFields.length; i++) {
        allFields.ids.push(theFields[i].id);
        allFields.names.push(theFields[i].name.toLowerCase());
    }  
    
    return allFields;
    
  }  
  
  
  // function that takes the story data and column headings, and tries to find the data that relates to those headings
  function getStory(data,headings,fields) {
   
    var story = [];
    for (var i = 0;i < headings.length;i++) {
      if (headings[i] !== "") {
        var fieldData = getDataForHeading(data,headings[i].toLowerCase(),fields);
        if (fieldData != null) {
          fieldData = parseObject(fieldData);
        }  
        story.push(fieldData);
      }  
    }        
    
    return story;
    
  }  
  
  // Given a matched property from the returned data, this tries to then handle spsocial cases of arrays and objects (Strings are left untouched)
  function parseObject(data) {
    
    var stringData = "";
    if (Array.isArray(data)) {
    
      for (var i = 0; i < data.length; i++) {
        if (stringData != "") {
            stringData+=",";
        }  
        if ( typeof data[i] === "object") {
          if (data[i].hasOwnProperty("id") && data[i].hasOwnProperty("value") && data[i].hasOwnProperty("self")) {
            stringData+= data[i]["value"];
          } 
          else if (data[i].hasOwnProperty("displayName")) {
            stringData+= data[i]["displayName"];
          } 
          else if (data[i].hasOwnProperty("name")) {
            stringData+= data[i]["name"];
          } 
          else {
            stringData+= JSON.stringify(data)
          }  
        }
        else {
          
          stringData+=data[i];
        }  
      }
    } 
    else if ( typeof data === "object") {
      if (data.hasOwnProperty("id") && data.hasOwnProperty("value") && data.hasOwnProperty("self")) {
            stringData+= data["value"];
      } 
      else if (data.hasOwnProperty("displayName")) {
            stringData+= data["displayName"];
      }  
      else if (data.hasOwnProperty("name")) {
            stringData+= data["name"];  
      } 
      else {
            stringData+= JSON.stringify(data)
          }  
      
    }
    else {
      stringData += data;
    }  
    return stringData;
  }  
  
  // Given a heading, interrogates the data to find a field with that name
  function getDataForHeading(data,heading,fields) {
    
        if (data.hasOwnProperty(heading)) {
          return data[heading];
        }  
        else if (data.fields.hasOwnProperty(heading)) {
          return data.fields[heading];
        }  
    
        var fieldName = getFieldName(heading,fields);
    
        if (fieldName !== "") {
          if (data.hasOwnProperty(fieldName)) {
            return data[fieldName];
          }  
          else if (data.fields.hasOwnProperty(fieldName)) {
            return data.fields[fieldName];
          }  
        }
    
        var splitName = heading.split(" ");
    
        if (splitName.length == 2) {
          if (data.fields.hasOwnProperty(splitName[0]) ) {
            if (data.fields[splitName[0]] && data.fields[splitName[0]].hasOwnProperty(splitName[1])) {
              return data.fields[splitName[0]][splitName[1]];
            }
            return "";
          }  
        }  
    
        return "Could not find value for " + heading;
        
  }  
  
  function getFieldName(heading,fields) {
    var index = fields.names.indexOf(heading);
    if ( index > -1) {
       return fields.ids[index]; 
    }
    return "";
  }  
  
// Fetches the updated "Incomplete Big Rocks" from Jira and pastes the data into the "Backlog" sheet
// TODO debug the data mismatch, figure out why the assignee doesn't corespond to the corect Jira ticket
  jiraPull();  

// Updates the Owner column for Incomplete Big Rocks
  var backlog = SpreadsheetApp.open(newReport).getSheetByName("Backlog");
  var incompleteBigRocksOwnerUpdated = backlog.getRange(2, 2, 200, 1).getValues();
  var incompletedBigRocksOwner = incompleteBigRocks.getRange(2, 1, 200, 1);
  incompletedBigRocksOwner.setValues(incompleteBigRocksOwnerUpdated);

// Updates the Estimated Completion Date column for Incomplete Big Rocks
  var incompleteBigRocksEstimatedCompletionDateUpdated = backlog.getRange(2, 3, 200, 1).getValues();
  var incompleteBigRocksEstimatedCompletionDate = incompleteBigRocks.getRange(2, 13, 200, 1);
  incompleteBigRocksEstimatedCompletionDate.setValues(incompleteBigRocksEstimatedCompletionDateUpdated);

// Updates the Estimate (in hours) column for Incomplete Big Rocks
  var incompleteBigRocksEstimateUpdated = backlog.getRange(2, 4, 200, 1).getValues();
  var incompleteBigRocksEstimate = incompleteBigRocks.getRange(2, 11, 200, 1);
  incompleteBigRocksEstimate.setValues(incompleteBigRocksEstimateUpdated);

// Clears "Backlog" sheet
  backlog.getRange(2, 1, backlog.getLastRow(), backlog.getLastColumn()).clearContent();

// Updates the "Instructions" issues list with the Big Rocks Ending this Week issues
  rangeInstructions.clearContent();
  rangeInstructions.setValue(valuesIssuesBigRocksEndingThisWeek);

// Fetches the updated "Big Rocks Ending This Week" from Jira and pastes the data into the "Backlog" sheet
  jiraPull();

// Updates the Owner column for BigRocks Ending This Week
  var bigRocksEndingThisWeekOwnerUpdated = backlog.getRange(2, 2, 200, 1).getValues();
  var bigRocksEndingThisWeekOwner = bigRocksEndingThisWeek.getRange(2, 1, 200, 1);
  bigRocksEndingThisWeekOwner.setValues(bigRocksEndingThisWeekOwnerUpdated);

  // Updates the Estimated Completion Date column for Incomplete Big Rocks
  var bigRocksEndingThisWeekEstimatedCompletionDateUpdated = backlog.getRange(2, 3, 200, 1).getValues();
  var bigRocksEndingThisWeekEstimatedCompletionDate = bigRocksEndingThisWeek.getRange(2, 13, 200, 1);
  bigRocksEndingThisWeekEstimatedCompletionDate.setValues(bigRocksEndingThisWeekEstimatedCompletionDateUpdated);

  // Updates the Estimate (in hours) column for Incomplete Big Rocks
  var bigRocksEndingThisWeekEstimateUpdated = backlog.getRange(2, 4, 200, 1).getValues();
  var bigRocksEndingThisWeekEstimate = incompleteBigRocks.getRange(2, 11, 200, 1);
  bigRocksEndingThisWeekEstimate.setValues(bigRocksEndingThisWeekEstimateUpdated);

  // Clears "Backlog" sheet
  backlog.getRange(2, 1, backlog.getLastRow(), backlog.getLastColumn()).clearContent();

  // Clears the "Instructions" issues list
  rangeInstructions.clearContent();
} 