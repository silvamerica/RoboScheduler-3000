/**
 * Installs the trigger
 * @return {void}
 */
function installTrigger() {
  var db = ScriptDb.getMyDb(),
      settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings'),
      settings = getSettings_(settingsSheet);
  if (!db.query({key: 'trigger'}).hasNext()) {
    var trigger = ScriptApp.newTrigger("runDaily")
          .timeBased()
          .everyDays(1)
          .atHour(parseInt(settings.startTime.value.getHours()))
          // @TODO(nsilva): Add localization to the trigger install
          // .inTimezone("America/Los_Angeles")
          .create();

    // Save trigger ID to the database
    db.save({
      key: 'trigger',
      value: trigger.getUniqueId()
    });
    // Each trigger has a unique ID
    Logger.log("Unique ID of Trigger: " + trigger.getUniqueId());
    Browser.msgBox('RoboScheduler 3000', 'The Periodic Scheduler has been installed. It will create the calendar event and send invites for the *next* meeting around the time of your current meeting.', Browser.Buttons.OK);
  } else {
    Browser.msgBox('RoboScheduler 3000', 'The Periodic Scheduler is already installed. To reinstall, remove it first.', Browser.Buttons.OK);
  }
}

/**
 * Run the scheduler daily
 * @return {void}
 */
function runDaily() {
  var settingsSheet = this.sheets.getSheetByName('Settings'),
      settings = getSettings_(settingsSheet);
      todaysMeeting = new Date(settings.nextMeeting.value);
  var today = new Date();
  todaysMeeting.setDate(nextMeeting.getDate() - settings.intervalInDays.value);
  if (todaysMeeting.setHours(0, 0, 0, 0) === today.setHours(0, 0, 0, 0)) {
    run();
  }
}

/**
 * Removes the installed trigger
 * @return {void}
 */
function removeTrigger() {
  var db = ScriptDb.getMyDb(),
      querySet = db.query({key: 'trigger'}),
      allTriggers = ScriptApp.getProjectTriggers(),
      record;
  // Loop over all triggers
  while (querySet.hasNext()) {
    record = querySet.next();
    for (var i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getUniqueId() === record.value) {
        // Found the trigger and now delete it
        ScriptApp.deleteTrigger(allTriggers[i]);
        break;
      }
    }
    db.remove(record);
  }
  Browser.msgBox('RoboScheduler 3000', 'The Periodic Scheduler has been removed.', Browser.Buttons.OK);
}

/**
 * Sends an invite for a meeting based on the settings sheet.
 * Updates the setting sheet.
 * 
 * @return {void}
 */
function run() {
  // Get the active spreadsheet workbook
  this.sheets = SpreadsheetApp.getActiveSpreadsheet();
  
  // Load the team sheet and the settings from the settings sheet
  var teamSheet = this.sheets.getSheetByName('Team'),
      settingsSheet = this.sheets.getSheetByName('Settings'),
      settings = getSettings_(settingsSheet);  
  
  // Collect up the team members to be invited, and modulate to wrap around if it goes off the end of the list
  var lastTeamMemberIndex = teamSheet.getLastRow(),
      teamMembers = teamSheet.getRange(2, 1, lastTeamMemberIndex - 1, 2),
      adjustedNextTeamMemberIndex = settings.nextTeamMember.value - 1,
      numTeamMembers = teamMembers.getNumRows(),
      nextTeamMembers = [],
      previousTeamMembers = [],
      allTeamEmails;

      Logger.log(teamMembers);
  for (var i = adjustedNextTeamMemberIndex; i < (adjustedNextTeamMemberIndex + settings.teamMembersPerEvent.value); i++) {
    nextTeamMembers.push({
      name: teamMembers.getCell((i % numTeamMembers), 1),
      email: teamMembers.getCell((i % numTeamMembers), 2)
    });
  }
  for (var j = adjustedNextTeamMemberIndex - settings.teamMembersPerEvent.value, index; j < adjustedNextTeamMemberIndex; j++) {
    index = (j < 0) ? j + numTeamMembers : j;
    previousTeamMembers.push({
      name: teamMembers.getCell((index), 1),
      email: teamMembers.getCell((index), 2)
    });
  }

  allTeamEmails = teamSheet.getRange(2, 2, lastTeamMemberIndex - 1).getValues();
  // Send the invite
  createInvite_(allTeamEmails, nextTeamMembers, settings);

  // Increment the settings
  incrementSettings_(settingsSheet, teamSheet, settings);

  // Update the team document with some indicator of the new state
  updateTeamSheet_(teamMembers, previousTeamMembers, nextTeamMembers);

  Browser.msgBox('RoboScheduler 3000', 'Event Scheduled!', Browser.Buttons.OK);
}

/**
 * Retrieves settings range and returns an object
 * 
 * @param {sheet}   sheet Settings Sheet
 * @return {object} representings the settings
 */
function getSettings_(sheet) {
  var settings = sheet.getRange("A2:B29").getValues(),
      returnObject = {};
  for (var i = 0, len = settings.length; i < len; i++) {
    if (settings[i][0] !== "") {
      returnObject[settings[i][0]] = {
        value:settings[i][1],
        location: 'B'+(i+2)
      };
    }
  }
  return returnObject;  
}

/**
 * Increments nextTeamMember and nextMeeting
 * 
 * @param {sheet}  settingsSheet The settings sheet
 * @param {sheet}  teamSheet The team sheet
 * @param {object} settings The settings object (that may have changed)
 * @return {void}
 */
function incrementSettings_(settingsSheet, teamSheet, settings) { 
  var lastTeamMemberIndex = teamSheet.getLastRow(),
      nextTeamMemberIndex = settings.nextTeamMember.value + settings.iterateBy.value;
  if (nextTeamMemberIndex > lastTeamMemberIndex) {
    nextTeamMemberIndex = (nextTeamMemberIndex % lastTeamMemberIndex) + 1;
  }
  
  settingsSheet.getRange(settings.nextTeamMember.location).setValue(nextTeamMemberIndex);
  
  var nextMeeting = new Date(settings.nextMeeting.value);
  nextMeeting.setDate(nextMeeting.getDate() + settings.intervalInDays.value);
  settingsSheet.getRange(settings.nextMeeting.location).setValue(nextMeeting);
}

/**
 * Creates an invite and invites all attendees to it
 * 
 * @param  {array}  allTeamEmails    An array of emails taken from the spreadsheet
 * @param  {array}  eventTeamMembers An array of objects taken from the spreadsheet
 * @param  {object} settings         The settings object
 * @return {void}
 */
function createInvite_(allTeamEmails, eventTeamMembers, settings) {
  var title            = settings.eventTitle.value,
      description      = settings.eventDescription.value + '\n\nCreated via RoboScheduler 3000',
      teamMemberNames  = [],
      teamMemberEmails = [];
  for (var i = 0, len = eventTeamMembers.length; i < len; i++) {
    teamMemberNames.push(eventTeamMembers[i].name.getValues()[0]);
    teamMemberEmails.push(eventTeamMembers[i].email.getValues()[0]);
  }
  settings.teamMember = teamMemberNames; // This feels wrong. :-(
  description = template_(description, settings);
  var whomToInvite;
  if (settings.inviteAllMembers.value === 1) {
    whomToInvite = allTeamEmails.join(',');
  } else {
    whomToInvite = teamMemberEmails.join(',');
  }
  createEventInvitePeople_(whomToInvite, title, description, settings);
}

/**
 * Creates an event, invites guests, sends invitation emails.
 * 
 * @param  {string} emailString A comma separated list of emails
 * @param  {string} title       The event title
 * @param  {string} description A description for the event
 * @param  {object} settings    The settings object
 * @return {void}
 */
function createEventInvitePeople_(emailString, title, description, settings) {  
  var cal = CalendarApp.getCalendarById(settings.calendarId.value),
      start = new Date(settings.nextMeeting.value);
      end = new Date(settings.nextMeeting.value);
  
  start.setHours(settings.startTime.value.getHours());
  end.setHours(settings.endTime.value.getHours());

  Logger.log('Date: ' + start);
  Logger.log('E-Mail: ' + emailString);
  Logger.log('Title: ' + title);
  Logger.log('Description: ' + description);

  if (settings.debug.value == 1) {
    Browser.msgBox('RoboScheduler 3000 Debug',
      'Date: ' + start + '\n\n' +
      'E-Mail: ' + emailString + '\n\n' +
      'Title: ' + title + '\n\n' +
      'Description: ' + description,
      Browser.Buttons.OK
    );
  } else {
    var event = cal.createEvent(title, start, end, {
        description : description,
        guests : emailString,
        sendInvites : 'true'
    });
    var reminders = settings.remindersInMinutes.value.split(',');
    for (var i = 0, len = reminders.length; i < len; i++) {
      Logger.log(reminders[i]);
      event.addPopupReminder(reminders[i].trim());
    }
    Logger.log('Event ID: ' + event.getId());
  }
}

function updateTeamSheet_(teamMembers, previousTeamMembers, nextTeamMembers) {
  teamMembers.clearFormat();
  var i, len;
  for (i = 0, len = previousTeamMembers.length; i < len; i++) {
    previousTeamMembers[i].name.setBackground('#F08080');
    previousTeamMembers[i].email.setBackground('#F08080');
  }
  for (i = 0, len = nextTeamMembers.length; i < len; i++) {
    nextTeamMembers[i].name.setBackground('#90EE90');
    nextTeamMembers[i].email.setBackground('#90EE90');
  }
}

/**
 * Substitute Placeholder Values in a Template
 * 
 * @param  {string} template The template
 * @param  {object} data     Key-value pairs of placeholder names and values
 * @return {string}          
 */
function template_(template, data) {
  var p, i;
  for (p in data) {
      if (data.hasOwnProperty(p)) {
          if (data[p] instanceof Array) {
            i = 0;
            template = template.replace(new RegExp('\\{\\{' + p + '\\}\\}', 'g'), function() {
              return data[p][i++];
            });
          } else {
            template = template.replace(new RegExp('\\{\\{' + p + '\\}\\}', 'g'), data[p].value);
          }
      }
  }
  return template;
}