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
  try {
    Logger.log("Starting ICS import...");

    // Remove a range of invisible Unicode characters that might be causing problems.
    // This removes: U+200B (zero-width space), U+200C, U+200D, U+200E, U+200F,
    // U+2028 (line separator), U+2029 (paragraph separator), and U+FEFF (BOM).
    icsContent = icsContent.replace(/[\u200B\u200C\u200D\u200E\u200F\u2028\u2029\uFEFF]/g, '');

    // Now split into raw lines on both Unix (\n) and Windows (\r\n) line breaks.
    var rawLines = icsContent.split(/\r?\n/);
    // Unfold lines: if a line begins with a space or tab, it's a continuation of the previous line.
    var lines = [];
    rawLines.forEach(function(line) {
      if (/^[ \t]/.test(line)) {
        if (lines.length > 0) {
          lines[lines.length - 1] += line.trim();
        }
      } else {
        lines.push(line);
      }
    });
    Logger.log("Total unfolded lines: " + lines.length);

    var events = [];
    var event = {};
    var rruleData = null;

    // Process each unfolded line.
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      
      if (line.startsWith("BEGIN:VEVENT")) {
        event = {};
        rruleData = null;
      } else if (line.startsWith("END:VEVENT")) {
        if (!event.start) {
          Logger.log("Skipped event with no DTSTART.");
        } else if (rruleData) {
          var expandedEvents = expandRecurringEvent(event, rruleData);
          Logger.log("Expanded recurring event '" + (event.summary || "No Summary") + "' to " + expandedEvents.length + " occurrence(s).");
          events = events.concat(expandedEvents);
        } else {
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
      // Additional properties (like ORGANIZER or ATTENDEE) are ignored for now.
    }
    
    Logger.log("Total events parsed before date filtering: " + events.length);
    
    // Filter events by date range if required.
    if (!importAllDates && fromDate && toDate) {
      var from = new Date(fromDate);
      var to = new Date(toDate);
      events = events.filter(function(evt) {
        var eventDate = parseICSDateTime(evt.start);
        return eventDate >= from && eventDate <= to;
      });
      Logger.log("Events remaining after date filtering: " + events.length);
    }
    
    if (events.length === 0) {
      Logger.log("No events found. Check the ICS file format and your filter settings.");
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.clear();
    
    // Build the output rows (header + event rows).
    var outputRows = [];
    outputRows.push(["Summary", "Start Date", "End Date", "Description", "Location", "Recurrence Info"]);
    
    events.forEach(function(evt) {
      outputRows.push([
        evt.summary || "",
        formatDateTime(evt.start),
        formatDateTime(evt.end),
        evt.description || "",
        evt.location || "",
        evt.recurrenceInfo || ""
      ]);
    });
    
    if (outputRows.length > 1) {
      sheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
      Logger.log("Wrote " + (outputRows.length - 1) + " event row(s) to the sheet.");
    } else {
      Logger.log("No data to write to the sheet.");
    }
    
    return "Import Complete";
  } catch (e) {
    Logger.log("Error in handleICSUpload: " + e.toString());
    throw e;
  }
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
  
  // Use COUNT if provided, otherwise default to 52 occurrences.
  var maxOccurrences = rruleData.COUNT ? parseInt(rruleData.COUNT, 10) : 52;
  // Use UNTIL if provided, otherwise default to 1 year from the start.
  var until = rruleData.UNTIL ? parseICSDateTime(rruleData.UNTIL) : new Date(startDate.getTime() + (365 * 24 * 60 * 60 * 1000));
  var freq = rruleData.FREQ;
  var interval = rruleData.INTERVAL ? parseInt(rruleData.INTERVAL, 10) : 1;
  
  // Define today's end-of-day in UTC.
  var now = new Date();
  var todayUTC = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 23, 59, 59, 999));
  Logger.log("Today (UTC end): " + todayUTC.toISOString());
  
  var currentDate = new Date(startDate);
  var count = 0;
  
  while (currentDate <= until && count < maxOccurrences) {
    if (currentDate > todayUTC) {
      Logger.log("Occurrence on " + currentDate.toISOString() + " is after today. Stopping expansion.");
      break;
    }
    
    var newEvent = {
      summary: baseEvent.summary,
      start: formatICSDateTime(currentDate),
      end: formatICSDateTime(new Date(currentDate.getTime() + duration)),
      description: baseEvent.description,
      location: baseEvent.location,
      recurrenceInfo: "Recurring " + (freq ? freq.toLowerCase() : "")
    };
    expandedEvents.push(newEvent);
    Logger.log("Added occurrence: " + newEvent.start);
    
    // Increment currentDate based on frequency.
    switch (freq) {
      case 'DAILY':
        currentDate.setUTCDate(currentDate.getUTCDate() + interval);
        break;
      case 'WEEKLY':
        currentDate.setUTCDate(currentDate.getUTCDate() + (7 * interval));
        break;
      case 'MONTHLY':
        currentDate.setUTCMonth(currentDate.getUTCMonth() + interval);
        break;
      case 'YEARLY':
        currentDate.setUTCFullYear(currentDate.getUTCFullYear() + interval);
        break;
      default:
        Logger.log("Unrecognized frequency: " + freq + ". Stopping expansion.");
        count = maxOccurrences; // Force exit on unrecognized frequency.
        break;
    }
    count++;
  }
  return expandedEvents;
}

function parseICSDateTime(icsDateTime) {
  if (!icsDateTime || icsDateTime.length < 15) {
    Logger.log("Invalid ICS datetime: " + icsDateTime);
    return new Date(0);
  }
  var year = parseInt(icsDateTime.substring(0, 4), 10);
  var month = parseInt(icsDateTime.substring(4, 6), 10) - 1;  // JavaScript months are zero-based.
  var day = parseInt(icsDateTime.substring(6, 8), 10);
  var hour = parseInt(icsDateTime.substring(9, 11), 10);
  var minute = parseInt(icsDateTime.substring(11, 13), 10);
  var second = parseInt(icsDateTime.substring(13, 15), 10);
  
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
  if (!icsDateTime || icsDateTime.length < 15) {
    return icsDateTime;
  }
  var year = icsDateTime.substring(0, 4);
  var month = icsDateTime.substring(4, 6);
  var day = icsDateTime.substring(6, 8);
  var hour = icsDateTime.substring(9, 11);
  var minute = icsDateTime.substring(11, 13);
  
  var timezone = icsDateTime.length > 15 ? icsDateTime.substring(15) : 'Z';
  var formatted = year + "-" + month + "-" + day + " " + hour + ":" + minute;
  
  if (timezone === 'Z') {
    formatted += " UTC";
  } else if (timezone.startsWith("-") || timezone.startsWith("+")) {
    var tzHours = timezone.substring(0, 3);
    var tzMinutes = timezone.substring(3, 5);
    formatted += " UTC" + tzHours + ":" + tzMinutes;
  }
  return formatted;
}
