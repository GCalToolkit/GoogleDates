// Provided for free by GCalTools www.gcaltools.com.  Use at your own risk.

// INSTRUCTIONS to Create Birthday Events from Contacts' Birthdays in any Google Calendar

// 1) Make a COPY of this project
//
// 2) Open your copy and set the calendarId below (line 24) to your own calendar ID 
//    from the settings page of the Google calendar that you want to use as your Birthday Calendar.
//
// 3) Configure Default Notifications for the Calendar if needed
// 
// 3) Make sure "Calendar" and "People" are listed under "Services" on the left. If not, click the + and add them
//
// 4) Select "updateBirthdays" above, then click "Run"
//
// 5) Click 'Advanced' during warnings to proceed, give permissions (to yourself, not me!)
//
// 6) DONE! You can run this script again to add new birthdays without creating duplicates

// Optional: Set up a "Trigger" in the menu on the left to run updateBirthdays daily to add birthdays from new contacts.

// ******** CHANGE THIS ID TO YOUR CALENDAR ID or use "primary" for your main calendar ********

var calendarId = "xxxxxxxxxxxxxxxxxxxxxxxxxx@group.calendar.google.com";




// DO NOT EDIT BELOW THIS LINE UNLESS YOU UNDERSTAND WHAT YOU ARE DOING

function updateBirthdays() {
  createBirthdayEventsForAllContacts(calendarId);
}

function createBirthdayEventsForAllContacts(calendarId) {
  const peopleService = People.People;
  const calendarService = CalendarApp;

  let pageToken = null;
  const pageSize = 100;

  do {
    var response;
try {
  response = People.People.Connections.list('people/me', {
    pageSize: pageSize,
    personFields: 'names,birthdays',
    pageToken: pageToken
  });
} catch (error) {
  Logger.log('Error fetching connections: ' + error.message);
}

    const connections = response.connections || [];

    connections.forEach(connection => {
      const names = connection.names || [];
      const birthdays = connection.birthdays || [];

      if (names.length > 0 && birthdays.length > 0) {
        const contactName = names[0].displayName;
        const birthday = birthdays[0].date;

        if (birthday) {
          // Handle cases where the birthday year might be undefined
          const year = birthday.year || new Date().getFullYear(); // Use current year if not specified
          const startDate = new Date(year, birthday.month - 1, birthday.day);
          const endDate = new Date(startDate);
          endDate.setDate(endDate.getDate() + 1);

          // Check for existing events specifically for this contact on the birthday date
          const existingEvents = calendarService.getCalendarById(calendarId).getEvents(startDate, endDate);

          const eventExists = existingEvents.some(event => 
            event.getTitle() === `${contactName}'s Birthday`
          );
          var event;
          // Create the event only if it doesn't already exist
          if (!eventExists) {
            event = calendarService.getCalendarById(calendarId).createAllDayEventSeries(
              `${contactName}'s Birthday`,
              startDate,
              CalendarApp.newRecurrence().addYearlyRule(),
            );

            Logger.log(`Birthday event created for ${contactName} on ${startDate.toDateString()}`);
          } else {
            Logger.log(`Birthday event already exists for ${contactName} on ${startDate.toDateString()}`);
          }
        }
      }
    });

    pageToken = response.nextPageToken;
  } while (pageToken);
}

