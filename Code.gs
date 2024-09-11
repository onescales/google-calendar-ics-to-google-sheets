function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ICS Import')
    .addItem('Upload and Import ICS', 'showUploadDialog')
    .addToUi();
}

function showUploadDialog() {
  var html = HtmlService.createHtmlOutputFromFile('UploadForm')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload ICS File');
}

function handleICSUpload(icsContent) {
  var lines = icsContent.split("\n");
  var events = [];
  var event = {};

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    
    if (line.startsWith("BEGIN:VEVENT")) {
      event = {};
    } else if (line.startsWith("END:VEVENT")) {
      events.push(event);
    } else if (line.startsWith("SUMMARY:")) {
      event.summary = line.replace("SUMMARY:", "");
    } else if (line.startsWith("DTSTART")) {
      event.start = formatDateTime(line.replace(/.*:/, ""));
    } else if (line.startsWith("DTEND")) {
      event.end = formatDateTime(line.replace(/.*:/, ""));
    } else if (line.startsWith("DESCRIPTION:")) {
      event.description = line.replace("DESCRIPTION:", "");
    } else if (line.startsWith("LOCATION:")) {
      event.location = line.replace("LOCATION:", "");
    }
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.appendRow(["Summary", "Start Date", "End Date", "Description", "Location"]);
  
  events.forEach(function(event) {
    sheet.appendRow([event.summary, event.start, event.end, event.description, event.location]);
  });

  return "Import Complete";
}

// Function to format the datetime string from the .ics format
function formatDateTime(icsDateTime) {
  // Example icsDateTime: "20230126T003000Z" or "20230126T003000-0500"
  var year = icsDateTime.substring(0, 4);
  var month = icsDateTime.substring(4, 6);
  var day = icsDateTime.substring(6, 8);
  var hour = icsDateTime.substring(9, 11);
  var minute = icsDateTime.substring(11, 13);
  var second = icsDateTime.substring(13, 15);

  // Check for time zone information
  var timezone = icsDateTime.length > 15 ? icsDateTime.substring(15) : 'Z';

  // Format date-time as "YYYY-MM-DD HH:MM" and handle timezones
  var formattedDateTime = year + "-" + month + "-" + day + " " + hour + ":" + minute;

  if (timezone === 'Z') {
    formattedDateTime += " UTC"; // Append UTC if 'Z' is present
  } else if (timezone.startsWith("-") || timezone.startsWith("+")) {
    // Handle time zone offsets like "-0500"
    var tzHours = timezone.substring(0, 3); // "-05" or "+01"
    var tzMinutes = timezone.substring(3, 5); // "00"
    formattedDateTime += " UTC" + tzHours + ":" + tzMinutes;
  }

  return formattedDateTime;
}
