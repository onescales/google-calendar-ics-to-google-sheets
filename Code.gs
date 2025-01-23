function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ICS Import')
    .addItem('Upload and Import ICS', 'showUploadDialog')
    .addToUi();
}

function showUploadDialog() {
  var html = HtmlService.createHtmlOutputFromFile('UploadForm')
      .setWidth(400)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload ICS File');
}

function handleICSUpload(icsContent, importAllDates, fromDate, toDate) {
  var lines = icsContent.split("\n");
  var events = [];
  var event = {};
  var rruleData = null;

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    
    if (line.startsWith("BEGIN:VEVENT")) {
      event = {};
      rruleData = null;
    } else if (line.startsWith("END:VEVENT")) {
      if (rruleData) {
        // Generate recurring events
        var expandedEvents = expandRecurringEvent(event, rruleData);
        events = events.concat(expandedEvents);
      } else {
        // Only add non-recurring events
        events.push(event);
      }
    } else if (line.startsWith("SUMMARY:")) {
      event.summary = line.replace("SUMMARY:", "");
    } else if (line.startsWith("DTSTART")) {
      event.start = line.replace(/.*:/, "");
    } else if (line.startsWith("DTEND")) {
      event.end = line.replace(/.*:/, "");
    } else if (line.startsWith("DESCRIPTION:")) {
      event.description = line.replace("DESCRIPTION:", "");
    } else if (line.startsWith("LOCATION:")) {
      event.location = line.replace("LOCATION:", "");
    } else if (line.startsWith("RRULE:")) {
      rruleData = parseRRule(line.replace("RRULE:", ""));
    }
  }

  // Filter events based on date range if not importing all dates
  if (!importAllDates && fromDate && toDate) {
    events = events.filter(function(event) {
      var eventDate = parseICSDateTime(event.start);
      return eventDate >= new Date(fromDate) && eventDate <= new Date(toDate);
    });
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.appendRow(["Summary", "Start Date", "End Date", "Description", "Location", "Recurrence Info"]);
  
  events.forEach(function(event) {
    sheet.appendRow([
      event.summary,
      formatDateTime(event.start),
      formatDateTime(event.end),
      event.description,
      event.location,
      event.recurrenceInfo || ''
    ]);
  });

  return "Import Complete";
}

function parseRRule(rrule) {
  var parts = rrule.split(";");
  var rruleObj = {};
  
  parts.forEach(function(part) {
    var [key, value] = part.split("=");
    rruleObj[key] = value;
  });
  
  return rruleObj;
}

function expandRecurringEvent(baseEvent, rruleData) {
  var expandedEvents = [];
  var startDate = parseICSDateTime(baseEvent.start);
  var endDate = parseICSDateTime(baseEvent.end);
  var duration = endDate - startDate;  // Duration in milliseconds
  
  // Default to 1 year of recurring events if no COUNT or UNTIL is specified
  var maxOccurrences = rruleData.COUNT ? parseInt(rruleData.COUNT) : 52; // ~1 year of weekly events
  var until = rruleData.UNTIL ? parseICSDateTime(rruleData.UNTIL) : 
              new Date(startDate.getTime() + (1 * 365 * 24 * 60 * 60 * 1000)); // 1 year
  
  var freq = rruleData.FREQ;
  var interval = rruleData.INTERVAL ? parseInt(rruleData.INTERVAL) : 1;
  
  var currentDate = new Date(startDate);
  var count = 0;
  
  while (currentDate <= until && count < maxOccurrences) {
    var newEvent = {
      summary: baseEvent.summary,
      start: formatICSDateTime(currentDate),
      end: formatICSDateTime(new Date(currentDate.getTime() + duration)),
      description: baseEvent.description,
      location: baseEvent.location,
      recurrenceInfo: `Recurring ${freq.toLowerCase()}`
    };
    
    expandedEvents.push(newEvent);
    
    // Increment the date based on frequency
    switch (freq) {
      case 'DAILY':
        currentDate.setDate(currentDate.getDate() + interval);
        break;
      case 'WEEKLY':
        currentDate.setDate(currentDate.getDate() + (7 * interval));
        break;
      case 'MONTHLY':
        currentDate.setMonth(currentDate.getMonth() + interval);
        break;
      case 'YEARLY':
        currentDate.setFullYear(currentDate.getFullYear() + interval);
        break;
    }
    
    count++;
  }
  
  return expandedEvents;
}

function parseICSDateTime(icsDateTime) {
  var year = parseInt(icsDateTime.substring(0, 4));
  var month = parseInt(icsDateTime.substring(4, 6)) - 1;  // JS months are 0-based
  var day = parseInt(icsDateTime.substring(6, 8));
  var hour = parseInt(icsDateTime.substring(9, 11));
  var minute = parseInt(icsDateTime.substring(11, 13));
  var second = parseInt(icsDateTime.substring(13, 15));
  
  return new Date(Date.UTC(year, month, day, hour, minute, second));
}

function formatICSDateTime(date) {
  return date.getUTCFullYear().toString().padStart(4, '0') +
         (date.getUTCMonth() + 1).toString().padStart(2, '0') +
         date.getUTCDate().toString().padStart(2, '0') + 'T' +
         date.getUTCHours().toString().padStart(2, '0') +
         date.getUTCMinutes().toString().padStart(2, '0') +
         date.getUTCSeconds().toString().padStart(2, '0') + 'Z';
}

function formatDateTime(icsDateTime) {
  var year = icsDateTime.substring(0, 4);
  var month = icsDateTime.substring(4, 6);
  var day = icsDateTime.substring(6, 8);
  var hour = icsDateTime.substring(9, 11);
  var minute = icsDateTime.substring(11, 13);
  var second = icsDateTime.substring(13, 15);

  var timezone = icsDateTime.length > 15 ? icsDateTime.substring(15) : 'Z';
  var formattedDateTime = year + "-" + month + "-" + day + " " + hour + ":" + minute;

  if (timezone === 'Z') {
    formattedDateTime += " UTC";
  } else if (timezone.startsWith("-") || timezone.startsWith("+")) {
    var tzHours = timezone.substring(0, 3);
    var tzMinutes = timezone.substring(3, 5);
    formattedDateTime += " UTC" + tzHours + ":" + tzMinutes;
  }

  return formattedDateTime;
}
