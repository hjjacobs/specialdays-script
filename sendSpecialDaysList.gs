/*
 * Global settings
 */
const currentDate                = new Date();
const currentYear                = currentDate.getFullYear();
const currentMonth               = currentDate.getMonth();
const MAIL_RECIPIENT             = "<<MAIL ADDRESS>>";   // Mail address of the recipient for the birthday list
const MAIL_SUBJECT_BIRTHDAYS     = "Resterende verjaardagen voor " + currentYear;
const MAIL_SUBJECT_ANNIVERSARIES = "Resterende jubilea voor " + currentYear;
const MAIL_HEADER_BIRTHDAYS      = "<h3>Verjaardagen tot eind " + currentYear + ":</h3><p>";
const MAIL_HEADER_ANNIVERSARIES  = "<h3>Jubilea tot eind " + currentYear + ":</h3><p>"
const MAIL_FOOTER                = "</p><p>\nTot de volgende keer!</p>";
const MAIL_MONTH_HEADER_BD       = "<h4>Deze maand jarig</br></h4>";
const MAIL_REST_YEAR_HEADER_BD   = "<h4>Verjaardagen rest van het jaar</br></h4>";
const MAIL_MONTH_HEADER_AN       = "<h4>Jubilea deze maand</br></h4>";
const MAIL_REST_YEAR_HEADER_AN   = "<h4>Jubilea rest van het jaar</br></h4>";

/*
 * Send an email with a list of special days (birthday, anniversary) of people in my Google contacts for the remainder of the current year.
 * - Also add calculated age of the people
 */
function sendSpecialDaysList() {

  // Filter contacts: a contact must have a full name (eliminate companies)
  // Get list of all contacts - order is determined by the getContacts method
  var contacts = ContactsApp.getContacts().filter(function (e) {
    return e.getFullName() != "";
  });
    
  let persons = convertContactsToPersons(contacts);
  let htmlBody = generateBirthdayList(persons);
  mailList(MAIL_SUBJECT_BIRTHDAYS, htmlBody);

  htmlBody = generateAnniversaryList(persons);
  mailList(MAIL_SUBJECT_ANNIVERSARIES, htmlBody);
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
 * Convert contacts to a more flattend person structure.
 * The person object contain only data necessary to create the list.
 */
function convertContactsToPersons(contacts) {
  let persons = [];

  for (let contact of contacts) {
    let birthdays     = contact.getDates(ContactsApp.Field.BIRTHDAY);
    let anniversaries = contact.getDates(ContactsApp.Field.ANNIVERSARY);
    var person        = new Person(contact.getFullName());

    for (let birthday of birthdays) {
      let day = birthday.getDay().toFixed();           // day of birth - removing the decimals
      let year = birthday.getYear().toFixed();         // year of birth - removing the decimals
      let month = (birthday.getMonth().ordinal()) + 1; // month of birth - convert enum to number (zero based)

      let calculatedAge = Math.round(currentYear - year).toFixed();

      person.setBirthday(year, month, day);
      person.setAge(calculatedAge);
    }

    for (let anniversary of anniversaries) {
      let day = anniversary.getDay().toFixed();           // day of birth - removing the decimals
      let year = anniversary.getYear().toFixed();         // year of birth - removing the decimals
      let month = (anniversary.getMonth().ordinal()) + 1; // month of birth - convert enum to number (zero based)

      let calculatedYears = Math.round(currentYear - year).toFixed();

      person.setAnniversary(year, month, day);
      person.setAnniversaryYears(calculatedYears);
    }

    // Only add a person to persons with a birthday and/or an anniversary
    if ((person.birthday != null) || (person.anniversary != null)) {
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

    htmlBody = htmlBody + `Op ${birthdayDate.getDate()}-${(birthdayDate.getMonth() + 1)} wordt ${e.name} <b>${e.age}</b> jaar (${birthdayDate.getFullYear()}).<br/>`;
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

    htmlBody = htmlBody + `Op ${anniversaryDate.getDate()}-${(anniversaryDate.getMonth() + 1)} heeft ${e.name} een <b>${e.anniversaryYears}</b>-jarig jubileum (${anniversaryDate.getFullYear()}).<br/>`;
  });

  // Close mail body
  htmlBody = htmlBody + MAIL_FOOTER;

  return htmlBody;
}