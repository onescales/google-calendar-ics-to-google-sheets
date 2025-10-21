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

    var rawLines = icsContent.split('\n');
    
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

    var events = [];
    var event = {};
    var rruleData = null;
    var inEvent = false;
    var inAlarm = false;
    var attendees = []; // Track attendees for current event

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      
      if (line.startsWith("BEGIN:VEVENT")) {
        event = {};
        rruleData = null;
        attendees = []; // Reset attendees for new event
        inEvent = true;
        inAlarm = false; 
      } else if (line.startsWith("END:VEVENT")) {
        inEvent = false;
        inAlarm = false;
        
        // Add attendees to event before processing
        if (attendees.length > 0) {
          event.attendees = attendees.join(", ");
        }
        
        if (!event.start) {
          Logger.log("Skipped event with no DTSTART.");
        } else if (rruleData) {
          var expandedEvents = expandRecurringEvent(event, rruleData);
          Logger.log("-> Expanded recurring event '" + (event.summary || "No Summary") + "' to " + expandedEvents.length + " occurrence(s).");
          events = events.concat(expandedEvents);
        } else {
          events.push(event);
        }
      } else if (inEvent) {
        if (line.startsWith("BEGIN:VALARM")) {
          inAlarm = true;
        } else if (line.startsWith("END:VALARM")) {
          inAlarm = false;
        } else if (!inAlarm) {
          if (line.startsWith("SUMMARY:")) {
            event.summary = line.substring(8);
          } else if (line.startsWith("DTSTART")) {
            var colonIndex = line.indexOf(':');
            if (colonIndex > -1) {
              event.start = line.substring(colonIndex + 1);
            } else {
              event.start = line.replace(/DTSTART[^:]*:?/, "");
            }
          } else if (line.startsWith("DTEND")) {
            var colonIndex = line.indexOf(':');
            if (colonIndex > -1) {
              event.end = line.substring(colonIndex + 1);
            } else {
              event.end = line.replace(/DTEND[^:]*:?/, "");
            }
          } else if (line.startsWith("DESCRIPTION:")) {
            event.description = line.substring(12);
          } else if (line.startsWith("LOCATION:")) {
            event.location = line.substring(9);
          } else if (line.startsWith("RRULE:")) {
            rruleData = parseRRule(line.substring(6));
          } else if (line.startsWith("ATTENDEE")) {
            // Extract attendee information
            var attendee = parseAttendee(line);
            if (attendee) {
              attendees.push(attendee);
            }
          }
        }
      }
    }
    
    Logger.log("Total events parsed before date filtering: " + events.length);
    
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
    
    var outputRows = [];
    outputRows.push(["Summary", "Start Date", "End Date", "Description", "Location", "Attendees", "Recurrence Info"]);
    
    events.forEach(function(evt) {
      outputRows.push([
        evt.summary || "",
        formatDateTime(evt.start),
        formatDateTime(evt.end),
        evt.description || "",
        evt.location || "",
        evt.attendees || "",
        evt.recurrenceInfo || ""
      ]);
    });
    
    if (outputRows.length > 1) {
      sheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
      Logger.log("Wrote " + (outputRows.length - 1) + " event row(s) to the sheet.");
    } else {
      Logger.log("No data to write to the sheet.");
    }
    
    return "Import Complete: " + (outputRows.length - 1) + " event(s) imported";
  } catch (e) {
    Logger.log("Error in handleICSUpload: " + e.toString());
    throw e;
  }
}

function parseAttendee(line) {
  // ATTENDEE lines can have parameters like CN (common name), ROLE, PARTSTAT, etc.
  // Example: ATTENDEE;CN=John Doe;ROLE=REQ-PARTICIPANT:mailto:john@example.com
  
  var email = "";
  var name = "";
  
  // Extract email (after mailto:)
  var mailtoIndex = line.indexOf("mailto:");
  if (mailtoIndex > -1) {
    email = line.substring(mailtoIndex + 7).trim();
  }
  
  // Extract common name (CN parameter)
  var cnMatch = line.match(/CN=([^;:]+)/);
  if (cnMatch && cnMatch[1]) {
    name = cnMatch[1].trim();
    // Remove quotes if present
    name = name.replace(/^["']|["']$/g, '');
  }
  
  // Return formatted attendee string
  if (name && email) {
    return name + " (" + email + ")";
  } else if (email) {
    return email;
  } else if (name) {
    return name;
  }
  
  return null;
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
  var duration = endDate - startDate;
  
  var freq = rruleData.FREQ;
  var interval = rruleData.INTERVAL ? parseInt(rruleData.INTERVAL, 10) : 1;
  var byDay = rruleData.BYDAY;
  
  var now = new Date();
  var todayUTC = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 23, 59, 59, 999));
  
  var until = rruleData.UNTIL ? parseICSDateTime(rruleData.UNTIL) : todayUTC;
  
  if (until > todayUTC) {
    until = todayUTC;
  }
  
  var currentDate = new Date(startDate);
  var count = 0;
  var maxOccurrences = rruleData.COUNT ? parseInt(rruleData.COUNT, 10) : 10000;
  
  if (freq === 'MONTHLY' && byDay) {
    var byDayInfo = parseByDay(byDay);
    var year = startDate.getUTCFullYear();
    var month = startDate.getUTCMonth();
    
    while (count < maxOccurrences) {
      var occurrenceDate = findNthWeekdayInMonth(year, month, byDayInfo.weekday, byDayInfo.position);
      
      if (!occurrenceDate) {
        count++;
        month += interval;
        if (month > 11) {
          year += Math.floor(month / 12);
          month = month % 12;
        }
        continue;
      }
      
      occurrenceDate.setUTCHours(startDate.getUTCHours());
      occurrenceDate.setUTCMinutes(startDate.getUTCMinutes());
      occurrenceDate.setUTCSeconds(startDate.getUTCSeconds());
      
      if (occurrenceDate > until) {
        break;
      }
      
      if (occurrenceDate >= startDate) {
        var newEvent = {
          summary: baseEvent.summary || "",
          start: formatICSDateTime(occurrenceDate),
          end: formatICSDateTime(new Date(occurrenceDate.getTime() + duration)),
          description: baseEvent.description || "",
          location: baseEvent.location || "",
          attendees: baseEvent.attendees || "",
          recurrenceInfo: "Recurring monthly (" + byDay + ")"
        };
        expandedEvents.push(newEvent);
      }
      
      count++;
      month += interval;
      if (month > 11) {
        year += Math.floor(month / 12);
        month = month % 12;
      }
    }
  } else {
    while (currentDate <= until && count < maxOccurrences) {
      var newEvent = {
        summary: baseEvent.summary || "",
        start: formatICSDateTime(currentDate),
        end: formatICSDateTime(new Date(currentDate.getTime() + duration)),
        description: baseEvent.description || "",
        location: baseEvent.location || "",
        attendees: baseEvent.attendees || "",
        recurrenceInfo: "Recurring " + (freq ? freq.toLowerCase() : "")
      };
      expandedEvents.push(newEvent);
      
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
          count = maxOccurrences;
          break;
      }
      count++;
    }
  }
  
  return expandedEvents;
}

function parseByDay(byDay) {
  var weekdayMap = {
    'SU': 0, 'MO': 1, 'TU': 2, 'WE': 3, 'TH': 4, 'FR': 5, 'SA': 6
  };
  
  var match = byDay.match(/^(-?\d+)?([A-Z]{2})$/);
  if (!match) {
    Logger.log("Invalid BYDAY format: " + byDay);
    return { position: 1, weekday: 5 };
  }
  
  var position = match[1] ? parseInt(match[1], 10) : 1;
  var weekdayCode = match[2];
  var weekday = weekdayMap[weekdayCode];
  
  return { position: position, weekday: weekday };
}

function findNthWeekdayInMonth(year, month, weekday, position) {
  if (position > 0) {
    var firstDay = new Date(Date.UTC(year, month, 1));
    var firstWeekday = firstDay.getUTCDay();
    var daysToAdd = (weekday - firstWeekday + 7) % 7;
    daysToAdd += (position - 1) * 7;
    var targetDate = new Date(Date.UTC(year, month, 1 + daysToAdd));
    
    if (targetDate.getUTCMonth() !== month) {
      return null;
    }
    
    return targetDate;
  } else if (position === -1) {
    var lastDay = new Date(Date.UTC(year, month + 1, 0));
    var lastWeekday = lastDay.getUTCDay();
    var daysToSubtract = (lastWeekday - weekday + 7) % 7;
    return new Date(Date.UTC(year, month + 1, 0 - daysToSubtract));
  }
  
  return null;
}

function parseICSDateTime(icsDateTime) {
  if (!icsDateTime) {
    Logger.log("Invalid ICS datetime: empty or null");
    return new Date(0);
  }
  
  var dateValue = icsDateTime.trim().replace(/Z$/, '');
  
  if (dateValue.length < 15) {
    if (dateValue.length === 8) {
        var year = parseInt(dateValue.substring(0, 4), 10);
        var month = parseInt(dateValue.substring(4, 6), 10) - 1;
        var day = parseInt(dateValue.substring(6, 8), 10);
        return new Date(Date.UTC(year, month, day, 0, 0, 0));
    }
    Logger.log("Invalid ICS datetime format: " + icsDateTime);
    return new Date(0);
  }
  
  var year = parseInt(dateValue.substring(0, 4), 10);
  var month = parseInt(dateValue.substring(4, 6), 10) - 1;
  var day = parseInt(dateValue.substring(6, 8), 10);
  var hour = parseInt(dateValue.substring(9, 11), 10);
  var minute = parseInt(dateValue.substring(11, 13), 10);
  var second = parseInt(dateValue.substring(13, 15), 10);
  
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
  if (!icsDateTime) return "";
  
  if (icsDateTime.length === 8 && !icsDateTime.includes('T')) {
      var year = icsDateTime.substring(0, 4);
      var month = icsDateTime.substring(4, 6);
      var day = icsDateTime.substring(6, 8);
      return year + "-" + month + "-" + day + " (All day)";
  }

  if (icsDateTime.length < 15) {
    return icsDateTime;
  }
  
  var year = icsDateTime.substring(0, 4);
  var month = icsDateTime.substring(4, 6);
  var day = icsDateTime.substring(6, 8);
  var hour = icsDateTime.substring(9, 11);
  var minute = icsDateTime.substring(11, 13);
  
  var formatted = year + "-" + month + "-" + day + " " + hour + ":" + minute;
  
  if (icsDateTime.endsWith('Z')) {
    formatted += " UTC";
  }
  return formatted;
}
