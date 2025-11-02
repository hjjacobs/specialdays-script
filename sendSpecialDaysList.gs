/*
 * Global settings
 */
const currentDate                = new Date();
const currentYear                = currentDate.getFullYear();
const currentMonth               = currentDate.getMonth();
const currentMonthString         = LanguageApp.translate(Utilities.formatDate(currentDate, "CET", "MMMM"), 'en', 'nl');
const MAIL_RECIPIENT             = "<<MAIL ADDRESS>>";   // Mail address of the recipient for the birthday list
const MAIL_SUBJECT_BIRTHDAYS     = "Verjaardagen voor " + currentMonthString + " " + currentYear;
const MAIL_SUBJECT_ANNIVERSARIES = "Jubilea voor " + currentMonthString + " " + currentYear;
const MAIL_HEADER_BIRTHDAYS      = "<h3>Verjaardagen in " + currentYear + "</h3><p>";
const MAIL_HEADER_ANNIVERSARIES  = "<h3>Jubilea in " + currentYear + "</h3><p>"
const MAIL_FOOTER                = "</p><p>\nTot de volgende maand!</p>";
const MAIL_MONTH_HEADER_BD       = "<h4>In " + currentMonthString + " jarig</br></h4>";
const MAIL_REST_YEAR_HEADER_BD   = "<h4>Verjaardagen rest van het jaar</br></h4>";
const MAIL_MONTH_HEADER_AN       = "<h4>Jubilea in " + currentMonthString + "</br></h4>";
const MAIL_REST_YEAR_HEADER_AN   = "<h4>Jubilea rest van het jaar</br></h4>";

/*
 * Send an email with a list of special days (birthday, anniversary) of people in my Google contacts for the remainder of the current year.
 * - Also add calculated age of the people
 */
function sendSpecialDaysList() {
  currentDate.setHours(0,0,0,0); // Set time part to zero for comparison purposes

  let contacts = getAllContacts();
  let persons = convertContactsToPersons(contacts);
  
  let htmlBody = generateBirthdayList(persons);
  mailList(MAIL_SUBJECT_BIRTHDAYS, htmlBody);

  htmlBody = generateAnniversaryList(persons);
  mailList(MAIL_SUBJECT_ANNIVERSARIES, htmlBody);
}

/*
 * Get all contacts using People API
 */
function getAllContacts() {
  let contacts = [];
  let pageToken = null;
  
  do {
    const response = People.People.Connections.list('people/me', {
      pageSize: 1000,
      personFields: 'names,birthdays,events',
      pageToken: pageToken
    });
    
    if (response.connections) {
      contacts = contacts.concat(response.connections);
    }
    
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  // Filter out contacts without names (companies, etc.)
  return contacts.filter(contact => {
    return contact.names && contact.names.length > 0 && contact.names[0].displayName;
  });
}

/*
 * Send mail using the MailApp
 */
function mailList(subject, body) {
  MailApp.sendEmail({
    to: MAIL_RECIPIENT,
    subject: subject,
    htmlBody: body
  });
}

/*
 * Person class
 * Constructor pattern
 */
function Person(name) {
  this.name = name;
  this.age  = null;
  this.birthday = null;
  this.birthdayThisYear = null;
  this.anniversaryYears = null;
  this.anniversary = null;
  this.anniversaryThisYear = null;

  this.setAge = function(age) {
    this.age = age;
  }

  this.setBirthday = function(year, month, day) {
    this.birthday = year+"-"+month+"-"+day;
    this.birthdayThisYear = currentYear+"-"+month+"-"+day;
  };

  this.setAnniversaryYears = function (years) {
    this.anniversaryYears = years;
  }

  this.setAnniversary = function(year, month, day) {
    this.anniversary = year+"-"+month+"-"+day;
    this.anniversaryThisYear = currentYear+"-"+month+"-"+day;
  };
}

/*
 * Convert contacts to a more flattened person structure.
 * The person object contains only data necessary to create the list.
 */
function convertContactsToPersons(contacts) {
  let persons = [];

  for (let contact of contacts) {
    let name = contact.names && contact.names.length > 0 ? contact.names[0].displayName : null;
    if (!name) continue;

    let person = new Person(name);
    let hasBirthdayOrAnniversary = false;

    // Process birthdays
    if (contact.birthdays && contact.birthdays.length > 0) {
      for (let birthday of contact.birthdays) {
        if (birthday.date && birthday.date.year && birthday.date.month && birthday.date.day) {
          let year = birthday.date.year;
          let month = birthday.date.month;
          let day = birthday.date.day;
          
          let calculatedAge = currentYear - year;
          
          person.setBirthday(year, month, day);
          person.setAge(calculatedAge);
          hasBirthdayOrAnniversary = true;
        }
      }
    }

    // Process anniversaries (stored in events with type 'anniversary')
    if (contact.events && contact.events.length > 0) {
      for (let event of contact.events) {
        if (event.type === 'anniversary' && event.date && event.date.year && event.date.month && event.date.day) {
          let year = event.date.year;
          let month = event.date.month;
          let day = event.date.day;
          
          let calculatedYears = currentYear - year;
          
          person.setAnniversary(year, month, day);
          person.setAnniversaryYears(calculatedYears);
          hasBirthdayOrAnniversary = true;
        }
      }
    }

    // Only add a person to persons with a birthday and/or an anniversary
    if (hasBirthdayOrAnniversary) {
      persons.push(person);
    }
  }

  return persons;
}

/*
 * Sort persons array on birthday this year and on name when the same birthday
 */
function sortPersonsOnBirthday(persons) {
  persons.sort((a, b) => {
    let da = new Date(a.birthdayThisYear),
        db = new Date(b.birthdayThisYear);

    let ret = da - db;

    if (ret == 0) {
        let fa = a.name.toLowerCase(), fb = b.name.toLowerCase();
        if (fa < fb) {
            ret = -1;
        }
        if (fa > fb) {
            ret = 1;
        }    
    }
    
    return ret;
  });
}

/*
 * Sort persons array on anniversary this year and on name
 */
function sortPersonsOnAnniversary(persons) {
  persons.sort((a, b) => {
    let da = new Date(a.anniversaryThisYear),
        db = new Date(b.anniversaryThisYear);

    let ret = da - db;

    if (ret == 0) {
        let fa = a.name.toLowerCase(), fb = b.name.toLowerCase();
        if (fa < fb) {
            ret = -1;
        }
        if (fa > fb) {
            ret = 1;
        }    
    }
    
    return ret;
  });
}

/*
 * Filter applicable birthdays
 * Generate a HTML body of remaining persons
 */
function generateBirthdayList(persons) {
  let currentMonthHeaderPrinted = false; // currentMonthHeaderPrinted - print a header when there are birthdays for the current month
  let restOfYearHeaderPrinted   = false; // restOfYearHeaderPrinted - print a header for the birthdays for the rest of year

  let htmlBody = MAIL_HEADER_BIRTHDAYS;

  // Filter persons
  // Only use persons who have a birthday and whose birthdays are in the future (including today)
  let filteredPersons = persons.filter(function (e){
    if (e.birthday == null) {
      return false;
    }

    let thisYearDate = new Date(e.birthdayThisYear);

    return thisYearDate >= currentDate;
  });

  sortPersonsOnBirthday(filteredPersons);

  // Add a line with birthday data for each birthday in the current month and future months
  filteredPersons.forEach((e) => {
    let birthdayDate = new Date(e.birthday);
    let thisYearDate = new Date(e.birthdayThisYear);

    // Only use contacts whose birthdays are in the future (including today)

    if ((currentMonth == thisYearDate.getMonth()) && !currentMonthHeaderPrinted) {
      htmlBody = htmlBody + MAIL_MONTH_HEADER_BD;
      
      currentMonthHeaderPrinted = true;
    } else if ((thisYearDate.getMonth() > currentMonth) && !restOfYearHeaderPrinted) {
      htmlBody = htmlBody + MAIL_REST_YEAR_HEADER_BD;
      
      restOfYearHeaderPrinted = true;
    }

    // Get the day of the week in dutch
    let dayOfWeek = LanguageApp.translate(Utilities.formatDate(thisYearDate, "CET", "EEEE"), 'en', 'nl');

    htmlBody = htmlBody + `Op ${dayOfWeek} ${birthdayDate.getDate()}-${(birthdayDate.getMonth() + 1)} wordt ${e.name} <b>${e.age}</b> jaar (${birthdayDate.getFullYear()}).<br/>`;
  });

  // Close mail body
  htmlBody = htmlBody + MAIL_FOOTER;

  return htmlBody;
}

/*
 * Generate a HTML body of all applicable anniversaries,
 */
function generateAnniversaryList(persons) {
  let currentMonthHeaderPrinted = false; // currentMonthHeaderPrinted - print a header when there are birthdays for the current month
  let restOfYearHeaderPrinted   = false; // restOfYearHeaderPrinted - print a header for the birthdays for the rest of year

  let htmlBody = MAIL_HEADER_ANNIVERSARIES;

  // Filter persons
  // Only use contacts who have an anniversary and whose anniversaries are in the future (including today)
  let filteredPersons = persons.filter(function (e){
    if (e.anniversary == null) {
      return false;
    }

    let thisYearDate = new Date(e.anniversaryThisYear);

    return thisYearDate >= currentDate;
  });

  sortPersonsOnAnniversary(filteredPersons);

  // Add a line with anniversary data for each anniversary in the current month and future months
  filteredPersons.forEach((e) => {
    let anniversaryDate = new Date(e.anniversary);
    let thisYearDate = new Date(e.anniversaryThisYear);

    // Only use contacts whose birthdays are in the future (including today)

    if ((currentMonth == thisYearDate.getMonth()) && !currentMonthHeaderPrinted) {
      htmlBody = htmlBody + MAIL_MONTH_HEADER_AN;
      
      currentMonthHeaderPrinted = true;
    } else if ((thisYearDate.getMonth() > currentMonth) && !restOfYearHeaderPrinted) {
      htmlBody = htmlBody + MAIL_REST_YEAR_HEADER_AN;
      
      restOfYearHeaderPrinted = true;
    }

    // Get the day of the week in dutch
    let dayOfWeek = LanguageApp.translate(Utilities.formatDate(thisYearDate, "CET", "EEEE"), 'en', 'nl');

    htmlBody = htmlBody + `Op ${dayOfWeek} ${anniversaryDate.getDate()}-${(anniversaryDate.getMonth() + 1)} heeft ${e.name} een <b>${e.anniversaryYears}</b>-jarig jubileum (${anniversaryDate.getFullYear()}).<br/>`;
  });

  // Close mail body
  htmlBody = htmlBody + MAIL_FOOTER;

  return htmlBody;
}