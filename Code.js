// Provided for free by GCalTools www.gcaltools.com: GCALTOOLS IS NOT RESPONSIBLE FOR ANY PROBLEMS!

// VIDEO: https://youtu.be/8GrGT8SWs-8?si=4ZdRLjQTOx7vApET
// CONTRIBUTE: https://github.com/GCalToolkit/GoogleDates

// INSTRUCTIONS to Create Birthdays and Special Events from Google Contacts in any Google Calendar

// 1) Make a COPY of this project: the script will run under your account and YOUR DATA IS COMPLETELY PRIVATE.

// 2) Open your copy and set the calendarId below (line 30) to your own calendar ID 
//    from the settings page of the Google calendar that you want to use as your Birthday/Events Calendar.

// 3) Configure Default Notifications for the Calendar if needed
 
// 3) Make sure "Calendar" and "People" are listed under "Services" on the left. If not, click the + and add them

// 4) Select "updateBirthdays" above, then click "Run"

// 5) Click 'Advanced' during warnings to proceed, give permissions (to yourself, not me!)

// 6) DONE! You can run this script again to add new birthdays without creating duplicates

// Optional: Create a "Trigger" in the menu on the left to run updateBirthdays daily to add birthdays from new contacts
// Optional: Set "noLabelTitle" to the title you want for special events with no label (line 33)
// Optional: Set "onlyBirthdays" to true (no quotation marks) if you ONLY want Birthdays (line 36)

// CONFIGURATION

// ******** CHANGE THIS ID TO YOUR CALENDAR ID or use "primary" for your main calendar ********
var calendarId = "xxxxxxxxxxxxxxxxxxxxxxxxxx@group.calendar.google.com";

// ******** CHANGE THIS TO THE NAME YOU WANT WHEN THERE IS NO LABEL FOR THE DATE ********
var noLabelTitle = "Special Event";

// ******** CHANGE THIS TO 'var onlyBirthdays = true' IF YOU ONLY WANT TO COPY BIRTHDAYS ********
var onlyBirthdays = false;


// DO NOT EDIT BELOW THIS LINE UNLESS YOU UNDERSTAND WHAT YOU ARE DOING

function updateBirthdays() {
  createSpecialEventsForAllContacts(calendarId);
}

function createSpecialEventsForAllContacts(calendarId) {
  const peopleService = People.People;
  const calendarService = CalendarApp;

  let pageToken = null;
  const pageSize = 100;

  do {
    var response;
    try {
      response = People.People.Connections.list('people/me', {
        pageSize: pageSize,
        personFields: 'names,birthdays,events',
        pageToken: pageToken
      });
    } catch (error) {
      Logger.log('Error fetching connections: ' + error.message);
    }

    const connections = response.connections || [];

    connections.forEach(connection => {
      const names = connection.names || [];
      const contactName = names.length > 0 ? names[0].displayName : 'Unnamed Contact';

      // Process Birthdays
      const birthdays = connection.birthdays || [];
      birthdays.forEach(birthday => {
        createOrUpdateEvent(calendarService, calendarId, contactName, birthday.date, `${contactName}'s Birthday`);
      });

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
}

function createOrUpdateEvent(calendarService, calendarId, contactName, eventDate, eventTitle) {
  if (eventDate) {
    // Handle cases where the event year might be undefined
    const year = eventDate.year || new Date().getFullYear(); // Use current year if not specified
    const startDate = new Date(year, eventDate.month - 1, eventDate.day);
    const endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + 1);

    // Check for existing events specifically for this contact on the event date
    const existingEvents = calendarService.getCalendarById(calendarId).getEvents(startDate, endDate);

    const eventExists = existingEvents.some(event => event.getTitle() === eventTitle);

    // Create the event only if it doesn't already exist
    if (!eventExists) {
      const event = calendarService.getCalendarById(calendarId).createAllDayEventSeries(
        eventTitle,
        startDate,
        CalendarApp.newRecurrence().addYearlyRule()
      );

      Logger.log(`${eventTitle} created for ${contactName} on ${startDate.toDateString()}`);
    } else {
      Logger.log(`${eventTitle} already exists for ${contactName} on ${startDate.toDateString()}`);
    }
  }
}