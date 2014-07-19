/**
 * Add a custom menu to the active spreadsheet.
 * @return void
 */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  // When the user selects "addMenuExample" menu, and clicks "Menu Entry 1", the function function1 is executed.
  menuEntries.push({name: "Install Trigger", functionName: "installTrigger"});
  menuEntries.push({name: "Menu Entry 2", functionName: "removeTrigger"});
  ss.addMenu("RoboScheduler 3000", menuEntries);
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
  var teamMembers = teamSheet.getRange(settings.nextTeamMember.value, 2, settings.teamMembersPerEvent.value).getValues(),
      lastTeamMemberIndex = teamSheet.getLastRow(),
      teamMemberIndices = [];

  for (var i = settings.nextTeamMember.value, max = i + settings.teamMembersPerEvent.value; i < max; i++) {
    if (i > lastTeamMemberIndex) {
      teamMembers[i - settings.nextTeamMember.value] = teamSheet.getRange((i % lastTeamMemberIndex) + 1, 2).getValues()[0];
    }
  }  
  // Send the invite
  createInvite_(teamMembers, settings);

  // Increment the settings
  incrementSettings_(settingsSheet, teamSheet, settings);

  // Update the team document with some indicator of the new state
  // updateTeamSheet_(teamSheet, settings);
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
 * @param  {array}  teamMembers An array of email addresses taken from the spreadsheet
 * @param  {object} settings    The settings object
 * @return {void}
 */
function createInvite_(teamMembers, settings) {
  var title       = settings.eventTitle.value,
      description = settings.eventDescription.value + '\n\nCreated via RoboScheduler 3000';

  settings.teamMember = teamMembers; // This feels wrong. :-(
  description = template_(description, settings);
  
  createEventInvitePeople_(teamMembers.join(','), title, description, settings);
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

  if (settings.debug.value == 1) {
    Logger.log('Date: ' + start);
    Logger.log('E-Mail: ' + emailString);
    Logger.log('Title: ' + title);
    Logger.log('Description: ' + description);
  } else {
    var event = cal.createEvent(title, start, end, {
        description : description,
        guests : emailString,
        sendInvites : 'true'
    });
    var reminders = settings.remindersInMinutes.value.split(',');
    for (var i, len = reminders.length; i < len; i++) {
      event.addPopupReminder(reminders[i].trim());
    }
    Logger.log('Event ID: ' + event.getId());
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
            template = template.replace(new RegExp('\\{\\{' + p + '\\}\\}', 'g'), data[p]);
          }
      }
  }
  return template;
}