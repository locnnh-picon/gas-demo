const spreadsheetId = '{{change it}}';
const backupFolderId =  '{{change it}}';
const now = new Date();
const timeExtendNum = 31 * 24 * 60 * 60 * 1000;
const actionTypeEnum = { CREATED: '作成', UPDATED: '更新', DELETED: '削除' };

const momentURL = "https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.30.1/moment.min.js";
let currentLogSheetName = '';
let oldLogSheetName = '';
let backupFileName = '';
const boxSpreadsheet = SpreadsheetApp.openById(spreadsheetId);

function main() {
  eval(UrlFetchApp.fetch(momentURL).getContentText());
  currentLogSheetName = `log_${moment().format('YYYYMMDD')}`;
  oldLogSheetName = `log_${moment().add(-1, 'year').format('YYYYMMDD')}`;
  backupFileName = `log_${moment().add(-1, 'year').format('YYYY')}年分`;
  const calendars = CalendarApp.getAllCalendars();
  for (const caldr of calendars) {
    updatePasswordToCalendar(caldr)
  }
}

function updatePasswordToCalendar(caldr) {
  const calendarId = caldr.getId();
  const calendarName = caldr.getName();
  console.log('Start calendar: %s (%s)', calendarName, calendarId);
  const optionalArgs = {
    timeMin: moment().toISOString(),
    showDeleted: true,
    singleEvents: true,
    orderBy: 'startTime'
  };

  const boxSetting = getBoxSettingValuesFromSpeadsheet(calendarName);
  if (!boxSetting || !boxSetting['boxSettingRow']) {
    console.log(`Calendar ${calendarName} not found setting password and description!`);
    return
  }
  const headerRow = boxSetting['headerRow'];
  const boxSettingRow = boxSetting['boxSettingRow'];
  let boxDesc = boxSettingRow[3];
  let boxPassword = undefined;


  try {
    const response = Calendar.Events.list(calendarId, optionalArgs);
    const events = response.items;
    if (response.nextSyncToken != null) {
      properties.setProperty('syncToken', response.nextSyncToken)
    }
    if (events.length === 0) {
      console.log('Finished! no events found.');
      return;
    }
    // Update password to calendar events
    for (const event of events) {
      const eventStartTime = moment(event.start.dateTime);
      console.log(`event ${event.summary}, start time ${eventStartTime.format("YYYY-MM-DD HH:mm:SS")} process update....`);

      const passwordUpdatedTag = event.extendedProperties && event.extendedProperties.private['updatedPwdbox'] == 'true';
      if (!passwordUpdatedTag) {

        const dayOfYear = getDayOfYear(moment(event.start.date || event.start.dateTime).toDate())
        for (let i = 4; i < headerRow.length; i++) {
          const dayArr = headerRow[i].split('-');
          if (dayOfYear >= parseInt(dayArr[0]) && dayOfYear <= parseInt(dayArr[1])) {
            boxPassword = boxSettingRow[i];
            break;
          }
        }
        const desc = (event.getDescription() || '') + `\n Password for box: ${boxPassword}` + `\n 詳細: ${boxDesc}`;
        event.description = desc;
        if (!event.extendedProperties) {
          event.extendedProperties = { "private": { 'updatedPwdbox': 'true' } };
        } else {
          event.extendedProperties.private['updatedPwdbox'] = 'true';
        }
      }

      // write log event if change
      insertEventChangeLog(calendarName, event);

      Calendar.Events.update(
        event,
        calendarId,
        event.id,
        { 'sendUpdates': 'all' }
      );
      console.log(`event ${event.summary}, start time ${event.start.dateTime} update password success!`);
    }

  } catch (err) {
    console.log('Execute update for %s error: %s', calendarName, err.message);
  }
  console.log('End calendar: %s', calendarName);
}

function getBoxSettingValuesFromSpeadsheet(calendarName) {
  const passwordSheet = boxSpreadsheet.getSheetByName('シート1');
  const values = passwordSheet.getDataRange().getValues();
  const headerRow = values[3];
  for (let i = 4; i < values.length; i++) {
    if (values[i][1].trim() === calendarName) {
      return { "headerRow": headerRow, "boxSettingRow": values[i] };
    }
  }

  return undefined;
}

function insertEventChangeLog(calendarName, event) {
  
  let logSheet = boxSpreadsheet.getSheetByName(currentLogSheetName);
  if(!logSheet){

    // rotation prev year for log
    const oldLogSheet = boxSpreadsheet.getSheetByName(oldLogSheetName);
    if(oldLogSheet){
      const backupFolder = DriveApp.getFolderById(backupFolderId);
      let backupFile = undefined;
      const searchBkFiles = backupFolder.getFilesByName(backupFileName);
      if(searchBkFiles.hasNext()){
        backupFile = searchBkFiles.next();
      }
      if (!backupFile){
        // DriveApp.createFile(backupFileName,  MimeType.GOOGLE_SHEETS)
        backupFile = SpreadsheetApp.create(backupFileName);
        oldLogSheet.copyTo(backupFile);
        DriveApp.getFileById(backupFile.getId()).moveTo(backupFolder);
      }else{
        backupFile = SpreadsheetApp.openById(backupFile.getId());
        oldLogSheet.copyTo(backupFile);
      }
      boxSpreadsheet.deleteSheet(oldLogSheet);
    }

    let templateSheet = boxSpreadsheet.getSheetByName('log_template')
    logSheet = boxSpreadsheet.insertSheet(currentLogSheetName, {template: templateSheet}).showSheet();
  }
  const logValues = logSheet.getDataRange().getValues();
  const lastRow = logValues[logSheet.getLastRow() - 1];
  let lastidx = 0;
  try {
    lastidx = parseInt(lastRow[0]) + 1;
  } catch {
    lastidx += 1;
  }
  if(isNaN(lastidx)){
    lastidx = 1;
  }
  const logCreateTagExist = event.extendedProperties && event.extendedProperties.private['createdLogTag'];
  if (!logCreateTagExist) {
    logSheet.appendRow([lastidx, moment().format('YYYY-MM-DD HH:mm:SS'), moment(event.start.dateTime).format('YYYY-MM-DD HH:mm:SS'), calendarName, event.summary, actionTypeEnum.CREATED]);
    if (!event.extendedProperties) {
      event.extendedProperties = { "private": { 'createdLogTag': true } };
    } else {
      event.extendedProperties.private['createdLogTag'] = true;
    }
    saveHistoryForProperties(event);
  } else {

    if (event.status === 'cancelled') {
      const deletedLogTag = event.extendedProperties && event.extendedProperties.private['deletedLogTag'];
      if (deletedLogTag) {
        return
      }
      logSheet.appendRow([lastidx, moment(event.updated).format('YYYY-MM-DD HH:mm:SS'), moment(event.start.dateTime).format('YYYY-MM-DD HH:mm:SS'), calendarName, event.summary, actionTypeEnum.DELETED]);
      if (!event.extendedProperties) {
        event.extendedProperties = { "private": { 'deletedLogTag': true } };
      } else {
        event.extendedProperties.private['deletedLogTag'] = true;
      }
    } else {
      const historyTag = event.extendedProperties && event.extendedProperties.private['history'];

      if (historyTag) {
        const history = JSON.parse(historyTag);
        if (!cmpDiffEventHistory(event, history)) {
          return;
        }
      }
      logSheet.appendRow([lastidx, moment(event.updated).format('YYYY-MM-DD HH:mm:SS'), moment(event.start.dateTime).format('YYYY-MM-DD HH:mm:SS'), calendarName, event.summary, actionTypeEnum.UPDATED]);
      saveHistoryForProperties(event);
    }
  }

}

function saveHistoryForProperties(event) {
  const lastHistory = {
    'startDateTime': event.start.dateTime,
    'endDateTime': event.end.dateTime,
    'attachments': !event.attachments ? '' : event.attachments.map((item) => item.title).sort((att1, att2) => att1 - att2),
    'attendeesEmails': !event.attachments ? '' : event.attendees.map((item) => item.email).sort((email1, email2) => email1 - email2),
    'summary': event.summary,
    'description': event.description,
    'location': event.location,
    'hangoutLink': event.hangoutLink,
  }
  if (!event.extendedProperties) {
    event.extendedProperties = { "private": { 'history': JSON.stringify(lastHistory) } };
  } else {
    event.extendedProperties.private['history'] = JSON.stringify(lastHistory);
  }
}

function cmpDiffEventHistory(event, history) {

  const currentAttendeesEmails = !event.attendees ? '' : event.attendees.map((item) => item.email).sort((email1, email2) => email1 - email2);
  const currentAttachments = !event.attachments ? '' : event.attachments.map((item) => item.title).sort((att1, att2) => att1 - att2);
  return event.start.dateTime !== history.startDateTime ||
    event.end.dateTime !== history.endDateTime ||
    JSON.stringify(currentAttendeesEmails) == history.attendeesEmails ||
    event.summary !== history.summary ||
    event.description !== history.description ||
    event.location !== history.location ||
    event.hangoutLink !== history.hangoutLink ||
    JSON.stringify(currentAttachments) == history.attachments

}

function getDayOfYear(targetDate) {
  const start = new Date(targetDate.getFullYear(), 0, 0);
  const diff = (targetDate - start) + ((start.getTimezoneOffset() - targetDate.getTimezoneOffset()) * 60 * 1000);
  const oneDay = 1000 * 60 * 60 * 24;
  const day = Math.floor(diff / oneDay);
  return day
}


