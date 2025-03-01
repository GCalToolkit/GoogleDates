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
var onlyContactLabel = false;
// TO GET THE contactLabelID OPEN https://contacts.google.com/ CLICK YOUR LABEL AND NOTE THE PAGE ADDRESS
// THE LAST PART OF THE ADDRESS IS THE contactLabelID: https://contacts.google.com/label/[contactLabelID]
var contactLabelID = "xxxxxxxxxxxxxx";

// REMINDER IN MINUTES
// addReminder must be set to none, email or popup. When using 'none' the calendar's default reminder will be applied, if set.
var addReminder = "none";
var reminderMinutes = 60 * 12; // 12 HOURS EARLIER = 12:00PM THE PREVIOUS DAY
// For hours/days write arithmetic e.g for 10pm four days earlier use: 3 * 24 * 60 + 2 * 60
// Note: birthdays start at 00:00 so the above is 3 days + 2 hours earlier (4 days earlier)


// ******** CONFIGURATION END : CLICK "SAVE PROJECT" ABOVE BEFORE YOU RUN! ********




// DO NOT EDIT BELOW THIS LINE UNLESS YOU UNDERSTAND WHAT YOU ARE DOING

function updateBirthdays() {
  createSpecialEventsForAllContacts(useOriginalBirthdayCalendar ? "primary" : calendarId);
}

function createSpecialEventsForAllContacts(calendarId) {
  const peopleService = People.People;
  const calendarService = CalendarApp;

  let pageToken = null;
  const pageSize = 100;

  try {
    do {
      var response;
      response = peopleService.Connections.list('people/me', {
        pageSize: pageSize,
        personFields: 'names,birthdays,events,memberships',
        pageToken: pageToken
      });

      const connections = response.connections || [];

      connections.forEach(connection => {
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
            createOrUpdateEvent(calendarService, calendarId, contactName, birthday.date, `${contactName}'s Birthday`);
          });
        }
        // Process Special Events (e.g., anniversaries, custom events with labels)
        if (!onlyBirthdays) {
          const events = connection.events || [];
          events.forEach(event => {
            const eventLabel = event.formattedType || noLabelTitle; // Use the label from formattedType or a default
            createOrUpdateEvent(calendarService, calendarId, contactName, event.date, `${contactName}'s ${eventLabel}`);
          });
        }
      });

      pageToken = response.nextPageToken;
    } while (pageToken);
  } catch (error) {
    Logger.log("Check the CONFIGURATION section is correct: " + error.message);
  }
}

function createOrUpdateEvent(calendarService, calendarId, contactName, eventDate, eventTitle) {
  var typeOfEvent = useOriginalBirthdayCalendar ? "birthday" : "default";
  if (eventDate) {

    // Handle cases where the event year might be undefined
    const year = eventDate.year || new Date().getFullYear(); // Use current year if not specified
    const startDate = new Date(year, eventDate.month - 1, eventDate.day);
    const endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + 1);

    // Check for existing events specifically for this contact on the event date
    const existingEvents = calendarService.getCalendarById(calendarId).getEvents(startDate, endDate);
    const eventExists = existingEvents.some(event => event.getTitle() === eventTitle);
    var event;
    // Create the event if it doesn't already exist
    if (!eventExists) {
      if (!useOriginalBirthdayCalendar) {
        // Use CalendarApp to create a regular event in a regular calendar
        event = calendarService.getCalendarById(calendarId).createAllDayEventSeries(
          eventTitle,
          startDate,
          CalendarApp.newRecurrence().addYearlyRule()
        );
        switch (addReminder) {
          case "email":
            event.addEmailReminder(reminderMinutes);
            break;
          case "popup":
            event.addPopupReminder(reminderMinutes);
            break;
          default:
            break;
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

        if (addReminder != "none") {
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
              overrides: [
                {
                  method: addReminder,
                  minutes: reminderMinutes
                }
              ]
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
            visibility: "private"
          }
        }
        event = Calendar.Events.insert(eventJSON, calendarId);
      }
      Logger.log(`${eventTitle} created for ${contactName} on ${startDate.toDateString()}`);
    } else {
      Logger.log(`${eventTitle} already exists for ${contactName} on ${startDate.toDateString()}`);
    }
  }
}


  // DELETE BIRTHDAYS
  // To be used if contacts have been deleted/edited, or you've changed the "'Birthday" text and want to start afresh.
  // Running this will delete ALL birthdays (from Contacts AND this script) on the special "Birthday" calendar
  // You can recreate Birthdays from Contacts by unchecking then rechecking "Sync from Contacts"
  // On the Settings page for the Birthdays calendar at this  address:
  // https://calendar.google.com/calendar/r/settings/birthdays

  // If you did not use the Official Birthday calendar for your events (i.e you set useOriginalBirthdayCalendar to false)
  // you can change the calendarToDeleteId to the Calendar ID of the calendar you used for your birthdays.
  // and the script will delete any events containing "Birthday" in the title.

  // For example:
  // calendarToDeleteId = "jhkhcjskchsjk26783chjsdkchsj178chsjdcks@group.calendar.google.com";

  // If you used a different word (e.g "Geburtstag") then change the word "Birthday" where you see:
  // event.summary.includes("Birthday") in the script below to: event.summary.includes("Geburtstag")

function deleteBirthdays() {
    // CHANGE primary ONLY ON THIS LINE!
    calendarToDeleteId = "primary";

    try {
    var pageToken;
    do {
      var response = Calendar.Events.list(calendarToDeleteId, { pageToken: pageToken });
      var events = response.items;
      
      if (!events || events.length === 0) {
        Logger.log("No events found.");
        return;
      }
      
      for (var i = 0; i < events.length; i++) {
        var event = events[i];
        
        // CHANGE ...event.summary.includes("Birthday") to ...event.summary.includes("Geburtstag") if necessary:
        if (event.eventType === "birthday" || (calendarToDeleteId != "primary" && event.summary.includes("Birthday"))) {
          // Checks if event type is a true birthday OR contains 'Birthday' if you aren't using the Official Birthday Calendar
          Calendar.Events.remove(calendarToDeleteId, event.id);
          Logger.log("Deleted event: " + event.summary);
        }
      }
      
      pageToken = response.nextPageToken;
    } while (pageToken);
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}