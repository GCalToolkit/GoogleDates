// Provided for free by GCalTools www.gcaltools.com.  Use at your own risk!
// VIDEO: https://youtu.be/8GrGT8SWs-8?si=4ZdRLjQTOx7vApET

// INSTRUCTIONS to Create Birthdays and Special Events from Google Contacts in any Google Calendar


// 1) Make a COPY of this project

// 2) Configure the settings for this script under "CONFIGURATION" below AND CLICK "SAVE PROJECT" ABOVE BEFORE YOU RUN!

// 3) Make sure "updateBirthdays" is selected in the menu above, then click "Run"

// 4) Click 'Advanced' during warnings to proceed, give permissions (to your own account, not mine!)

// 5) DONE! You can run this script again to add new birthdays without creating duplicates

// 6) Optional: Create a "Trigger" to run updateBirthdays daily to add birthdays from new contacts (see VIDEO linked above)


// ******** CONFIGURATION START : CLICK "SAVE PROJECT" ABOVE BEFORE YOU RUN! ********


// REQUIRED:

// CHANGE THIS to 'var useOriginalBirthdayCalendar = true' TO USE GOOGLE's DEFAULT BIRTHDAY CALENDAR
// THIS WILL CREATE DUPLICATES BIRTHDAYS/SPECIAL EVENTS IF YOU HAVE ENABLED SYNCING FROM CONTACTS IN BIRTHDAY SETTINGS
var useOriginalBirthdayCalendar = false;

// CHANGE THIS ID TO THE CALENDAR ID FOR THE CALENDAR YOU WANT TO USE
// THIS WILL BE IGNORED IF useOriginalBirthdayCalendar ABOVE IS true
var calendarId = "xxxxxxxxxxxxxxxxxxxxxxx@group.calendar.google.com";

// ALL SETTINGS BELOW ARE OPTIONAL:

// CHANGE THIS TO THE NAME YOU WANT WHEN THERE IS NO LABEL FOR THE SPECIAL EVENT
var noLabelTitle = "Special Event";

// CHANGE THIS TO 'var onlyBirthdays = true' IF YOU ONLY WANT TO COPY BIRTHDAYS, NOT SPECIAL EVENTS
var onlyBirthdays = false;

// CHANGE THIS TO 'var onlyContactLabel = true' IF YOU ONLY WANT TO COPY BIRTHDAYS FOR CONTACTS WITH A SPECIFIC LABEL
// DON'T FORGET TO SET THE contactLabelID BELOW IF THIS IS true
// USEFUL IF YOU HAVE TOO MANY CONTACTS AND THE SCRIPT WON'T FINISH
var onlyContactLabel = false;
// TO GET THE contactLabelID OPEN https://contacts.google.com/ CLICK YOUR LABEL AND NOTE THE PAGE ADDRESS
// THE LAST PART OF THE ADDRESS IS THE contactLabelID: https://contacts.google.com/label/[contactLabelID]
var contactLabelID = "xxxxxxxxxxxxxx";

// REMINDER IN MINUTES
// Use the following variables to configure your reminders
// Set to 0 to disable a particular reminder type
// Set useDefaultReminders to true to use the calendar's default reminder settings

// USE CALENDAR'S DEFAULT REMINDERS?
// WARNING: GOOGLE REMINDERS BUG!
// When using Original Birthday Calendar, the Main Calendar 'Event notifications' will be applied, NOT 'All Day Event notifications' and NOT 'Birthday' calendar reminder settings!
// When using a Secondary Calendar, that calendars default reminders will be applied - which you can change later.
// PREFER USING DEFAULT REMINDERS SO YOU CAN CHANGE THEM IN THE CALENDAR SETTINGS LATER.
var useDefaultReminders = true;

// EMAIL REMINDERS (set to 0 to disable)
var emailReminder1 = 0;          // minutes before (e.g., 60 for 1 hour, 1440 for 1 day)
var emailReminder2 = 0;          // minutes before (second email reminder, 0 to disable)

// POPUP REMINDERS (set to 0 to disable)
var popupReminder1 = 60 * 12;    // 12 hours before (720 minutes)
var popupReminder2 = 0;          // minutes before (second popup reminder, 0 to disable)

// For hours/days, here are some examples:
// 30 minutes = 30
// 1 hour = 60
// 12 hours = 60 * 12 = 720
// 1 day = 60 * 24 = 1440
// 2 days = 60 * 24 * 2 = 2880
// 1 week = 60 * 24 * 7 = 10080

// *** DON'T MODIFY THIS SECTION ***
// This creates the reminders array from your settings above
var reminders = [];
if (!useDefaultReminders) {
  if (popupReminder1 > 0) reminders.push({ method: "popup", minutes: popupReminder1 });
  if (popupReminder2 > 0) reminders.push({ method: "popup", minutes: popupReminder2 });
  if (emailReminder1 > 0) reminders.push({ method: "email", minutes: emailReminder1 });
  if (emailReminder2 > 0) reminders.push({ method: "email", minutes: emailReminder2 });
}
// *** END OF REMINDERS CONFIGURATION ***

// ADDITIONAL CONFIGURATION OPTIONS:

// CHANGE THIS TO true TO PREVIEW CHANGES WITHOUT ACTUALLY CREATING/DELETING EVENTS
// Or run the "dryRunUpdate" function to see what would be changed
var dryRun = false;

// CUSTOMIZE EVENT TITLE FORMATS (leave blank to use default format)
// Available variables: {name} - contact's name, {eventType} - the type of event
var birthdayTitleFormat = ""; // Default: "{name}'s Birthday"
var specialEventTitleFormat = ""; // Default: "{name}'s {eventType}"

// ADD CUSTOM DESCRIPTIONS TO EVENTS (Only possible if you are using a secondary calendar)
var addCustomDescriptions = false;
var birthdayDescription = "Birthday celebration for {name}";
var specialEventDescription = "{eventType} for {name}";

// FILTER EVENTS BY DATE RANGE
// USEFUL IF YOU HAVE TOO MANY CONTACTS AND THE SCRIPT WON'T FINISH
// Set specific months or days to include, or leave empty [] for all
// Examples:
// var filterMonths = [1, 2];    // Only January and February
// var filterMonths = [12];      // Only December
// var filterMonths = [];        // All months (default)
var filterMonths = [];          // 1-12 for specific months, empty array for all months

// Examples:
// var filterDays = [1, 15, 31]; // Only 1st, 15th and 31st days of the month
// var filterDays = [25];        // Only 25th day of the month
// var filterDays = [];          // All days (default)
var filterDays = [];            // 1-31 for specific days, empty array for all days

// MORE FLEXIBLE CLEANUP OPTIONS
var deleteSearchPattern = ""; // Custom text to search for when deleting events (empty for default)
var deleteOnlyFutureEvents = false; // Set to true to keep past events when deleting


// ******** CONFIGURATION END : CLICK "SAVE PROJECT" ABOVE BEFORE YOU RUN! ********



// These are the only two functions that appear in the "Run" menu
// DO NOT EDIT BELOW THIS LINE UNLESS YOU UNDERSTAND WHAT YOU ARE DOING

/**
 * Updates birthdays and special events from Google Contacts to your calendar
 */
function updateBirthdays() {
  if (dryRun) {
    Logger.log("DRY RUN MODE: No events will be created, only logged");
  }
  GCalTools.createSpecialEventsForAllContacts(useOriginalBirthdayCalendar ? "primary" : calendarId);
}

/**
 * Deletes birthday and special events from the calendar based on your configuration
 */
function deleteEvents() {
  const calendarToDeleteId = useOriginalBirthdayCalendar ? 'primary' : calendarId;
  let patterns = [];
  
  if (useOriginalBirthdayCalendar) {
    // For primary calendar, use the standard pattern that matches the birthday flag
    patterns.push(deleteSearchPattern || "Birthday");
  } else {
    // For secondary calendars, we'll gather labels from contacts in the utility function
    if (deleteSearchPattern) {
      patterns.push(deleteSearchPattern);
    } else {
      patterns.push("Birthday"); // Always include Birthday
    }
  }
  
  if (dryRun) {
    Logger.log("DRY RUN MODE: No events will be deleted, only logged");
  }
  
  return GCalTools.deleteEvents(calendarToDeleteId, patterns, deleteOnlyFutureEvents, dryRun);
}

/**
 * Displays what would be changed in the Execution Log without editing the calendar
 * (useful for testing)
 */

function dryRunUpdate() {
  dryRun = true;
  updateBirthdays();
  dryRun = false;
}

/**
 * Displays what would be deleted in the Execution Log without editing the calendar
 * (useful for testing)
 */

function dryRunDelete() {
  dryRun = true;
  deleteEvents();
  dryRun = false;
}

/**
 * Displays the current configuration in the Execution Log
 */

function showConfiguration() {
  return GCalTools.showConfiguration();
}