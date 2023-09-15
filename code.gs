function createCalendarEvent(rowNumber) {
  // enter the row number for processing
  var rowNumber = 2

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CinemaBookings");
  var range = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn());
  var values = range.getValues();
  var listOfCinemas = [
    { "name": "Cineworld Leicester Square", "address": "Leicester Square, 5-6 Leicester Square, WC2H 7NA" },
    { "name": "Cineworld West India Quay", "address": "x" },  
  ];

  // film information from the google sheet
  var filmName = values[0][6]  
  var filmScreeningDate = values[0][0];
  var filmScreeningStartTime = values[0][1];
  var filmScreeningEndTime = values[0][3];
  var filmBookingConfirmation = values[0][5]  
  var cinemaName = values[0][4]  

  // normalise the date and time objects
  var filmScreeningStartTimeCorrected = new Date(correctDate(filmScreeningStartTime, filmScreeningDate));
  var filmScreeningEndTimeCorrected = new Date(correctDate(filmScreeningEndTime, filmScreeningDate));  

  // grab the cinema information
  var cinemaLookupObject = findDictionaryWithValue(cinemaName, listOfCinemas)

  var cinemaLookUpLocation = JSON.stringify(cinemaLookupObject.address)
  var cinemaLookUpName = JSON.stringify(cinemaLookupObject.name)
  var calendarEventDescription =  filmName + "\n" + JSON.parse(cinemaLookUpName) + "\n" + JSON.parse(cinemaLookUpLocation) + "\nSeat: " + values[0][7]
  
  var calendarEventTitle = filmName + " " + filmBookingConfirmation

  // create a new calendar event
  var event = CalendarApp.createEvent(
    calendarEventTitle, 
    filmScreeningStartTimeCorrected, 
    filmScreeningEndTimeCorrected, 
    {location: JSON.parse(cinemaLookUpLocation), description: calendarEventDescription});
}

function correctDate(wrongScreeningDatetimeObject, correctScreeningDate) {
  var date = correctScreeningDate.getDate();
  var month = correctScreeningDate.getMonth();
  var year = correctScreeningDate.getFullYear();

  wrongScreeningDatetimeObject.setYear(year)   
  wrongScreeningDatetimeObject.setMonth(month)   
  wrongScreeningDatetimeObject.setDate(date)   

  return wrongScreeningDatetimeObject;
}

function findDictionaryWithValue(valueToFind, listOfDictionaries) {
  for (var i = 0; i < listOfDictionaries.length; i++) {
    var dictionary = listOfDictionaries[i];
    for (var key in dictionary) {
      if (dictionary[key] === valueToFind) {
        return dictionary;
      }
    }
  }
  return null;
}