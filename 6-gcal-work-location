const CALENDAR_ID = 'primary';
const WFH_NAME = 'John';

/**
 * Convert Date object to string of type 'YYYY-MM-DD'
 * @param {Date} date The date object to be converted
 * @returns {String}  Formatted date as a string
 */
function dateToString(date) {
  const dateMonth = (date.getMonth() + 1).toString().padStart(2, "0");
  const dateDay = (date.getDate()).toString().padStart(2, "0");
  return `${date.getFullYear()}-${dateMonth}-${dateDay}`;
}

/**
  * Parses working location properties of an event into a string.
  * See https://developers.google.com/calendar/api/v3/reference/events#resource
  */
function parseWorkingLocation(event) {
  if (event.eventType != "workingLocation") {
    throw new Error("'" + event.summary + "' is not a working location event.");
  }

  var location = 'No Location';
  const workingLocation = event.workingLocationProperties;
  if (workingLocation) {
    if (workingLocation.type === 'homeOffice') {
      location = 'Home';
    }
    if (workingLocation.type === 'officeLocation') {
      location = workingLocation.officeLocation.label;
    }
    if (workingLocation.type === 'customLocation') {
      location = workingLocation.customLocation.label;
    }
  }
  return `${event.start.date}: ${location}`;
}

/**
 * Make event object for working location-type
 * @param {String} workDate String-formatted date to use for all-day event 
 * @param {Boolean} isHomeOffice Bool determining if current date is from home or at office 
 * @param {String|undefined} updateEventId The event ID to update, only to pass to `createEvent`  
 */
function createWorkingLocation(workDate, isHomeOffice, updateEventId) {
  let endDate = new Date(workDate);
  endDate.setDate(endDate.getDate() + 1); // Increment by 1 day
  const officeProperties = {
      type: 'officeLocation',
      officeLocation: { label: "Kontoret" },
  };
  const homeProperties = {
    type: 'homeOffice',
    homeOffice: true,
  };

  let event = {
    start: { date: workDate },
    end: { date: dateToString(endDate) },
    eventType: "workingLocation",
    visibility: "public",
    transparency: "transparent",
    workingLocationProperties: isHomeOffice ? homeProperties : officeProperties,
  };
  console.log(event);
  createEvent(event, updateEventId);
}


/**
  * Creates a Calendar event.
  * See https://developers.google.com/calendar/api/v3/reference/events/insert
  */
function createEvent(event, updateEventId) {
  const calendarId = CALENDAR_ID;

  try {
    var response;
    if (updateEventId){
      response = Calendar.Events.update(event, calendarId, updateEventId);
    }
    else {
      response = Calendar.Events.insert(event, calendarId);
    }
    var event = (response.eventType === 'workingLocation') ? parseWorkingLocation(response) : response;
    console.log(event);
  } catch (exception) {
    console.log(exception.message);
  }
}

/**
 * Fetches all workingLocation-type events within specified date and returns its event ID
 * @param {String} dateString Date to search for as string
 * @returns {String|undefined} Event ID as a string or null if there are no events for specified date
 */
function getWorkLocationEventIdFromDate(dateString){
  let startTimestamp = new Date(dateString);
  startTimestamp.setDate(startTimestamp.getDate() - 1);
  let timeMin = dateToString(startTimestamp) + "T23:59:59Z";
  let timeMax = dateString + "T00:00:00Z";

  let response = Calendar.Events.list(
    CALENDAR_ID, optionalArgs = {
      eventTypes: ["workingLocation"],
      showDeleted: false,
      singleEvents: true,
      timeMin: timeMin,
      timeMax: timeMax,
      }
  );
  let events = response.items;
  if (events.length > 1){
    throw new Error("More than one event of type workingLocation found for date.");
  }
  else if (events.length == 0){
    return null;
  }
  return events[0].id;
}

/**
 * Main function that scans spreadsheet and calls other functions
 */
function updateWorkingLocation() {
  const dateColumn = 3;
  const nameColumn = 4;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet(); 
  const today = new Date();
  const todayStr = dateToString(today);
  for (let row = 130; row < 1000; row++){

    // First check if date column is blank. That means we are done
    const displ_date_i = sheet.getRange(row, dateColumn).getDisplayValue();
    if (displ_date_i === ""){
      break;
    }

    // Assume dates on format dd.mm.yyyy and zero-padded for d<10
    const date_i = displ_date_i.split(".").reverse().join("-");
    const dateobj_i = new Date(date_i);
    
    if(dateobj_i.getDay() === 0 || dateobj_i.getDay() === 6){
      // Skips adding events during the weekend
      continue
    }

    const name = sheet.getRange(row, nameColumn).getValue();

    // Get working-location type events for date_i
    const existingEventId = getWorkLocationEventIdFromDate(date_i);

    // Create (or update) event:
    createWorkingLocation(date_i, name === WFH_NAME, existingEventId);

  }

}
