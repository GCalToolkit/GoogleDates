/**
 * Utility functions for Google Dates
 * This file contains all the helper functions that don't need to appear in the Run menu
 */

var GCalTools = {};

// Main function to process contacts and create events
GCalTools.createSpecialEventsForAllContacts = function(calendarId) {
  const peopleService = People.People;
  const calendarService = CalendarApp;
  let stats = { processed: 0, created: 0, skipped: 0, errors: 0 };
  
  let pageToken = null;
  const pageSize = 100;
  let calendarNotFound = false;  // Add a flag to track calendar not found state

  try {
    // Verify calendar exists before starting to process contacts
    try {
      calendarService.getCalendarById(calendarId);
    } catch (error) {
      Logger.log(`Error: Calendar not found or invalid ID: ${calendarId}`);
      return stats;
    }
    
    do {
      var response;
      response = peopleService.Connections.list('people/me', {
        pageSize: pageSize,
        personFields: 'names,birthdays,events,memberships',
        pageToken: pageToken
      });

      const connections = response.connections || [];

      connections.forEach(connection => {
        if (calendarNotFound) return;  // Skip processing if calendar not found
        const names = connection.names || [];
        const memberships = connection.memberships || [];
        let hasLabel = false;
        memberships.forEach(membership => {
          if (membership.contactGroupMembership != null && membership.contactGroupMembership.contactGroupId.includes(contactLabelID)) {
            hasLabel = true;
          }
        });

        const contactName = names.length > 0 ? names[0].displayName : 'Unnamed Contact';

          if (!onlyContactLabel || hasLabel) {
              // Process Birthdays
              const birthdays = connection.birthdays || [];
              birthdays.forEach(birthday => {
                  if (calendarNotFound) return;  // Skip processing if calendar not found
                  if (GCalTools.shouldProcessDate(birthday.date)) {
                      let title = GCalTools.formatEventTitle(birthdayTitleFormat || `{name}'s Birthday`, contactName, "Birthday");
                      let description = addCustomDescriptions ? GCalTools.formatEventTitle(birthdayDescription, contactName, "Birthday") : "";
                      let result = GCalTools.createOrUpdateEvent(calendarService, calendarId, contactName, birthday.date, title, description);
                      if (result === 'notfound') {
                          calendarNotFound = true;  // Set flag to stop processing
                          return;
                      }

                      GCalTools.updateStats(stats, result);
                  }
              });
          }
        
        // Process Special Events (e.g., anniversaries, custom events with labels)
        if (!calendarNotFound && !onlyBirthdays) {  // Only process if calendar is found
          const events = connection.events || [];
          events.forEach(event => {
            if (calendarNotFound) return;  // Skip processing if calendar not found
            if (GCalTools.shouldProcessDate(event.date)) {
              const eventLabel = event.formattedType || noLabelTitle;
              let title = GCalTools.formatEventTitle(specialEventTitleFormat || `{name}'s {eventType}`, contactName, eventLabel);
              let description = addCustomDescriptions ? GCalTools.formatEventTitle(specialEventDescription, contactName, eventLabel) : "";
              let result = GCalTools.createOrUpdateEvent(calendarService, calendarId, contactName, event.date, title, description);
              if (result === 'notfound') {
                  calendarNotFound = true;  // Set flag to stop processing
                  return;
              }
              GCalTools.updateStats(stats, result);
            }
          });
        }
      });

      pageToken = response.nextPageToken;
    } while (pageToken && !calendarNotFound);  // Stop paging if calendar not found
    
    // Log summary statistics
    Logger.log(`Summary: Processed ${stats.processed} contacts, created ${stats.created} events, skipped ${stats.skipped} existing events, encountered ${stats.errors} errors`);
    
  } catch (error) {
    Logger.log("Check the CONFIGURATION section is correct: " + error.message);
  }
  
  return stats;
};

// Helper for date filtering
GCalTools.shouldProcessDate = function(date) {
  if (!date) return false;
  
  // If arrays are empty, include all months/days
  const includeAllMonths = filterMonths.length === 0;
  const includeAllDays = filterDays.length === 0;
  
  // Check if the date's month is in the filterMonths array
  const monthMatches = includeAllMonths || filterMonths.includes(date.month);
  
  // Check if the date's day is in the filterDays array
  const dayMatches = includeAllDays || filterDays.includes(date.day);
  
  // Both month and day must match the filter criteria
  return monthMatches && dayMatches;
};

// Format titles with variables
GCalTools.formatEventTitle = function(format, name, eventType) {
  return format.replace('{name}', name).replace('{eventType}', eventType);
};

// Update statistics
GCalTools.updateStats = function(stats, result) {
  stats.processed++;
  if (result === 'created') stats.created++;
  else if (result === 'skipped') stats.skipped++;
  else if (result === 'error') stats.errors++;
};

// Create or update a calendar event
GCalTools.createOrUpdateEvent = function(calendarService, calendarId, contactName, eventDate, eventTitle, eventDescription = "") {
  var typeOfEvent = useOriginalBirthdayCalendar ? "birthday" : "default";
  if (eventDate) {

    // Handle cases where the event year might be undefined
    const year = eventDate.year || new Date().getFullYear(); // Use current year if not specified
    const startDate = new Date(year, eventDate.month - 1, eventDate.day);
    const endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + 1);

    // Check for existing events specifically for this contact on the event date
    let existingEvents = [];
    try {
        existingEvents = calendarService.getCalendarById(calendarId).getEvents(startDate, endDate);    
    } catch (error) {
        Logger.log(`Error: Calendar not found or invalid ID: ${calendarId}`);
        return 'notfound';
    }

    const eventExists = existingEvents.some(event => event.getTitle() === eventTitle);
    var event;
    // Create the event if it doesn't already exist
    if (!eventExists) {
      if (dryRun) {
        Logger.log(`DRY RUN: Would create ${eventTitle} for ${contactName} on ${startDate.toDateString()}`);
        return 'created';
      }
      
      if (!useOriginalBirthdayCalendar) {
        try {
          // Use CalendarApp to create a regular event in a regular calendar
          event = calendarService.getCalendarById(calendarId).createAllDayEventSeries(
            eventTitle,
            startDate,
            CalendarApp.newRecurrence().addYearlyRule(),
            { description: eventDescription }
          );
          
          // Apply all configured reminders
          if (!useDefaultReminders && reminders.length > 0) {
            reminders.forEach(reminder => {
              if (reminder.method === "email") {
                event.addEmailReminder(reminder.minutes);
              } else if (reminder.method === "popup") {
                event.addPopupReminder(reminder.minutes);
              }
            });
          }
        } catch (error) {
          Logger.log(`Error creating event for ${contactName}: ${error.message}`);
          return 'error';
        }
      } else {
        // Use Calendar Service to create a 'birthday' event in the primary calendar
        var sdd = startDate.getDate();
        var smm = startDate.getMonth() + 1;
        var syyyy = startDate.getFullYear();
        var edd = endDate.getDate();
        var emm = endDate.getMonth() + 1;
        var eyyyy = endDate.getFullYear();

        rrule = "RRULE:FREQ=YEARLY";
        // Exception for Feb 29th!
        if (smm === 2 && sdd === 29) rrule = "RRULE:FREQ=YEARLY;BYMONTH=2;BYMONTHDAY=-1";

        if (!useDefaultReminders && reminders.length > 0) {
          // Create event with custom reminders
          var reminderOverrides = reminders.map(r => ({
            method: r.method,
            minutes: r.minutes
          }));
            
          eventJSON = {
            start: { date: syyyy + "-" + smm + "-" + sdd },
            end: { date: eyyyy + "-" + emm + "-" + edd },
            eventType: 'birthday',
            recurrence: [rrule],
            summary: eventTitle,
            transparency: "transparent",
            visibility: "private",
            reminders: {
              useDefault: false,
              overrides: reminderOverrides
            }
          }
        } else {
          eventJSON = {
            start: { date: syyyy + "-" + smm + "-" + sdd },
            end: { date: eyyyy + "-" + emm + "-" + edd },
            eventType: 'birthday',
            recurrence: [rrule],
            summary: eventTitle,
            transparency: "transparent",
            visibility: "private",
            reminders: {
              useDefault: true
            }
          }
        }
        
        // NOTE: Don't add description to birthday event types as it will cause an API error
        // The 'birthday' event type doesn't support descriptions in Google Calendar API
        // Only add description for non-birthday events
        if (eventDescription && typeOfEvent !== "birthday") {
          eventJSON.description = eventDescription;
        }
        
        try {
          event = Calendar.Events.insert(eventJSON, calendarId);
        } catch (error) {
          Logger.log(`Error creating event for ${contactName}: ${error.message}`);
          return 'error';
        }
      }
      Logger.log(`${eventTitle} created for ${contactName} on ${startDate.toDateString()}`);
      return 'created';
    } else {
      Logger.log(`${eventTitle} already exists for ${contactName} on ${startDate.toDateString()}`);
      return 'skipped';
    }
  }
  return 'error';
};

// Function to format minutes into a more readable format
GCalTools.formatMinutes = function(minutes) {
  if (minutes < 60) return `${minutes} minutes`;
  if (minutes === 60) return "1 hour";
  if (minutes < 1440) {
    const hours = minutes / 60;
    return `${hours} hours`;
  }
  const days = minutes / 1440;
  if (days === 1) return "1 day";
  return `${days} days`;
};

// UI Functions

// Function to create a menu in Google Sheets UI for easy access
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Google Dates')
    .addItem('Update Birthdays & Events', 'updateBirthdays')
    .addSeparator()
    .addItem('Delete Birthdays', 'deleteBirthdays')
    .addSeparator()
    .addItem('Preview Changes (Dry Run)', 'dryRunUpdate')
    .addSeparator()
    .addItem('Show Configuration', 'showConfiguration')
    .addToUi();
}

// Also add menu for Docs and Forms since this might be used in different contexts
function onOpenDocs() {
  var ui = DocumentApp.getUi();
  createCustomMenu(ui);
}

function onOpenForms() {
  var ui = FormApp.getUi();
  createCustomMenu(ui);
}

// Helper function to create consistent menu across apps
function createCustomMenu(ui) {
  ui.createMenu('Google Dates')
    .addItem('Update Birthdays & Events', 'updateBirthdays')
    .addSeparator()
    .addItem('Delete Birthdays', 'deleteBirthdays')
    .addSeparator()
    .addItem('Preview Changes (Dry Run)', 'dryRunUpdate')
    .addSeparator()
    .addItem('Show Configuration', 'showConfiguration')
    .addToUi();
}

// Improved configuration UI to display all important settings
GCalTools.showConfigUi = function() {
  var remindersList = [];
  if (useDefaultReminders) {
    remindersList.push("Using calendar default reminders");
  } else {
    if (popupReminder1 > 0) remindersList.push(`Popup reminder ${GCalTools.formatMinutes(popupReminder1)} before`);
    if (popupReminder2 > 0) remindersList.push(`Popup reminder ${GCalTools.formatMinutes(popupReminder2)} before`);
    if (emailReminder1 > 0) remindersList.push(`Email reminder ${GCalTools.formatMinutes(emailReminder1)} before`);
    if (emailReminder2 > 0) remindersList.push(`Email reminder ${GCalTools.formatMinutes(emailReminder2)} before`);
    if (remindersList.length === 0) remindersList.push("No reminders configured");
  }
  
  var monthsDisplay = filterMonths.length === 0 ? "All" : filterMonths.sort((a, b) => a - b).join(", ");
  var daysDisplay = filterDays.length === 0 ? "All" : filterDays.sort((a, b) => a - b).join(", ");
  
  var currentSettings = `
    <style>
      .settings-table { border-collapse: collapse; width: 100%; }
      .settings-table td, .settings-table th { border: 1px solid #ddd; padding: 8px; }
      .settings-table tr:nth-child(even) { background-color: #f2f2f2; }
      .settings-row-header { font-weight: bold; }
    </style>
    <h2>Google Dates Configuration</h2>
    <table class="settings-table">
      <tr><td class="settings-row-header">Calendar ID:</td><td>${calendarId}</td></tr>
      <tr><td class="settings-row-header">Using Original Birthday Calendar:</td><td>${useOriginalBirthdayCalendar}</td></tr>
      <tr><td class="settings-row-header">Only Birthdays:</td><td>${onlyBirthdays}</td></tr>
      <tr><td class="settings-row-header">Reminders:</td><td>${remindersList.join("<br>")}</td></tr>
      <tr><td class="settings-row-header">No Label Title:</td><td>${noLabelTitle}</td></tr>
      <tr><td class="settings-row-header">Only Contact Label:</td><td>${onlyContactLabel}</td></tr>
      <tr><td class="settings-row-header">Contact Label ID:</td><td>${contactLabelID}</td></tr>
      <tr><td class="settings-row-header">Dry Run Mode:</td><td>${dryRun}</td></tr>
      <tr><td class="settings-row-header">Filter by Months:</td><td>${monthsDisplay}</td></tr>
      <tr><td class="settings-row-header">Filter by Days:</td><td>${daysDisplay}</td></tr>
      <tr><td class="settings-row-header">Add Custom Descriptions:</td><td>${addCustomDescriptions}</td></tr>
    </table>
    <p>To modify settings, edit the configuration section in the script editor.</p>
    <p><small>This tool is provided by GCalTools (www.gcaltools.com)</small></p>
  `;
  
  var html = HtmlService.createHtmlOutput(currentSettings)
    .setWidth(450)
    .setHeight(400)
    .setTitle('Google Dates Configuration');
  
  // Try different UI contexts since we don't know which Google app is active
  try {
    SpreadsheetApp.getUi().showModalDialog(html, 'Google Dates Configuration');
  } catch (e) {
    try {
      DocumentApp.getUi().showModalDialog(html, 'Google Dates Configuration');
    } catch (e2) {
      try {
        FormApp.getUi().showModalDialog(html, 'Google Dates Configuration');
      } catch (e3) {
        try {
          // If running in standalone script, use this method
          var htmlOutput = HtmlService.createHtmlOutput(currentSettings)
            .setWidth(450)
            .setHeight(400);
          ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL)
          ScriptApp.getUi().showModalDialog(htmlOutput, 'Google Dates Configuration');
        } catch (e4) {
          // If all else fails, log the configuration
          Logger.log("Could not display UI. Current configuration:");
          Logger.log("Calendar ID: " + calendarId);
          Logger.log("Using Original Birthday Calendar: " + useOriginalBirthdayCalendar);
          Logger.log("Only Birthdays: " + onlyBirthdays);
          // More logs...
        }
      }
    }
  }
};

GCalTools.showConfiguration = function() {
  // For standalone execution from script editor
  Logger.log("=== Current Google Dates Configuration ===");
  Logger.log("Calendar ID: " + calendarId);
  Logger.log("Using Original Birthday Calendar: " + useOriginalBirthdayCalendar);
  Logger.log("Only Birthdays: " + onlyBirthdays);
  Logger.log("No Label Title: " + noLabelTitle);
  Logger.log("Only Contact Label: " + onlyContactLabel);
  Logger.log("Contact Label ID: " + contactLabelID);
  Logger.log("Custom Birthday Title Format: " + (birthdayTitleFormat || "(default)"));
  Logger.log("Custom Special Event Title Format: " + (specialEventTitleFormat || "(default)"));
  Logger.log("Add Custom Descriptions: " + addCustomDescriptions);
  Logger.log("Filter by Months: " + (filterMonths.length === 0 ? "All" : filterMonths.sort((a, b) => a - b).join(", ")));
  Logger.log("Filter by Days: " + (filterDays.length === 0 ? "All" : filterDays.sort((a, b) => a - b).join(", ")));
  Logger.log("Delete Search Pattern: " + (deleteSearchPattern || "(default)"));
  Logger.log("Delete Only Future Events: " + deleteOnlyFutureEvents);
  Logger.log("Dry Run Mode: " + dryRun);
  Logger.log("Reminders:");
  if (useDefaultReminders) {
    Logger.log("  Using calendar default reminders");
  } else {
    if (popupReminder1 > 0) Logger.log(`  Popup reminder ${GCalTools.formatMinutes(popupReminder1)} before`);
    if (popupReminder2 > 0) Logger.log(`  Popup reminder ${GCalTools.formatMinutes(popupReminder2)} before`);
    if (emailReminder1 > 0) Logger.log(`  Email reminder ${GCalTools.formatMinutes(emailReminder1)} before`);
    if (emailReminder2 > 0) Logger.log(`  Email reminder ${GCalTools.formatMinutes(emailReminder2)} before`);
    if (popupReminder1 === 0 && popupReminder2 === 0 && emailReminder1 === 0 && emailReminder2 === 0) {
      Logger.log("  No reminders configured");
    }
  }
  Logger.log("=== End of Configuration ===");
  
  // Also try to show UI if possible
  try {
    GCalTools.showConfigUi();
  } catch (e) {
    Logger.log("Configuration has been logged. Check the Logs panel (Ctrl+Enter or Cmd+Enter) to view it.");
  }
  
  return "Configuration has been logged. Check the Logs panel (Ctrl+Enter or Cmd+Enter) to view it.";
};

/**
 * Delete events from a calendar matching criteria
 * @param {string} calendarId - ID of the calendar to delete events from
 * @param {string} pattern - Text pattern to match in event titles (for non-primary calendars)
 * @param {boolean} onlyFutureEvents - Whether to only delete future events
 * @param {boolean} isDryRun - Whether to simulate deletion without actually deleting
 * @return {number} Number of events deleted
 */
GCalTools.deleteEvents = function(calendarId, pattern, onlyFutureEvents, isDryRun) {
  var eventsDeleted = 0;

  try {
    // First validate that the calendar exists before proceeding
    try {
      var calendar = CalendarApp.getCalendarById(calendarId);
      if (!calendar) {
        Logger.log(`Error: Calendar not found or invalid ID: ${calendarId}`);
        return -1; // Return a special code to indicate calendar not found
      }
    } catch (calError) {
      Logger.log(`Error: Calendar not found or invalid ID: ${calendarId}`);
      return -1; // Return a special code to indicate calendar not found
    }
    
    var pageToken;
    do {
      var optionalArgs = { 
        pageToken: pageToken,
      };
      
      // Only get future events if enabled
      if (onlyFutureEvents) {
        var today = new Date();
        optionalArgs.timeMin = today.toISOString();
      }
      
      var response = Calendar.Events.list(calendarId, optionalArgs);
      var events = response.items;
      
      if (!events || events.length === 0) {
        Logger.log("No events found.");
        return eventsDeleted;
      }
      
      for (var i = 0; i < events.length; i++) {
        var event = events[i];
        
        // More flexible deletion criteria
        var shouldDelete = event.eventType === "birthday";
        
        if (!shouldDelete && calendarId !== "primary") {
          shouldDelete = event.summary.includes(pattern);
        }
        
        if (shouldDelete) {
          if (isDryRun) {
            Logger.log("DRY RUN: Would delete event: " + event.summary);
          } else {
            Calendar.Events.remove(calendarId, event.id);
            Logger.log("Deleted event: " + event.summary);
          }
          eventsDeleted++;
        }
      }
      
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    Logger.log(`Total events deleted: ${eventsDeleted}`);
    return eventsDeleted;
  } catch (e) {
    Logger.log("Error: " + e.message);
    return eventsDeleted;
  }
};
