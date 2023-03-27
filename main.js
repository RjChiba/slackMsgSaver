
// set your `Bot User OAuth Access Token`
const TOKEN = '';

// set your time zone
const timeZone = 'Asia/Tokyo';

let userCache = {};
let channelCache = {};
let exceptionChannel = [];

function main(){
  // setup sheet
  const setUpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setups');

  // get user data 
  importSlackUsernameToCache();

  // get channel data
  importSlackChannelnameToCache();

  // get exception channel list
  getExceptionChannel(setUpSheet);

  // get range of data to load
  const oldest = getOldestUnixTime(setUpSheet);

  const now = new Date();

  // get log and write sheet
  // except exceptionChannel
  Object.keys(channelCache).forEach((id) => {
    if(!exceptionChannel.includes(id)){
      importSlackDataToSheet(id, oldest);
    }
  }, channelCache);

  // set next date
  setNextDate(now, setUpSheet);
}

function replaceText(text) {
  const regex = /<@(.*?)>/g;
  const replacedText = text.replace(regex, (match, userId) => {
    if (userCache[userId]) {
      return `@${userCache[userId]}`;
    } else {
      return match;
    }
  });

  return replacedText;
}

function getExceptionChannel(sheet){
  const exceptions = sheet.getRange('B3:B8').getValues();
  exceptions.forEach((channel) => {
    if(channel[0] !== ''){
      exceptionChannel.push(channel[0]);
    }
  });
}

function setNextDate(ts, sheet){
  // write spreadsheet
  // 1. operated time
  sheet.getRange('B1').setValue(Utilities.formatDate(ts, timeZone, 'yyyy-MM-dd hh:mm:ss'));
  
  // 2. next operation time
  let delta = sheet.getRange('B3').getValue();
  if(delta == ''){ delta = 30; };
  
  ts.setDate(ts.getDate() + delta);
  sheet.getRange('B2').setValue(Utilities.formatDate(ts, timeZone, 'yyyy-MM-dd hh:mm:ss'));
  sheet.getRange('C2').setValue(Date.parse(ts));


  // set trigger

  // 1. delete all triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // 2. set new trigger
  ScriptApp.newTrigger("main")
    .timeBased()
    .at(ts)
    .inTimezone(timeZone)
    .create();
}

function getOldestUnixTime(sheet){
  let date = sheet.getRange('C2').getValue();
  if( date == '' ){
    date = 0;
  }
  return date
}

function importSlackUsernameToCache(){
  var apiUrl = "https://slack.com/api/users.list";

  const headers = {
    'Authorization': 'Bearer ' + TOKEN,
    "Content-Type": "application/json; charset=utf-8"
  };

  const options = {
    "method": "get",
    "headers": headers,
    "muteHttpExceptions": true
  };

  var res = UrlFetchApp.fetch(apiUrl, options);
  var userData = JSON.parse(res.getContentText());

  if(userData.ok){
    userData.members.forEach((user) => {
      userCache[user.id] = user.profile.display_name || user.name;
    });
  }else{
    console.log("In `importSlackUsernameToCache`");
    console.log(userData.error);
    return 0;
  }
}

function importSlackChannelnameToCache(){
  const apiUrl = "https://slack.com/api/conversations.list";

  const headers = {
    'Authorization': 'Bearer ' + TOKEN,
    "Content-Type": "application/json; charset=utf-8"
  };

  const options = {
    "method": "get",
    "headers": headers,
    "muteHttpExceptions": true
  };

  var res = UrlFetchApp.fetch(apiUrl, options);
  var channelData = JSON.parse(res.getContentText());

  if(!channelData.ok){
    console.log("In `importSlackChannelnameToCache`");
    console.log(channelData.error);
    return 0;
  }

  channelData.channels.forEach( (channel) => {
    channelCache[channel.id] = channel.name;
  })

}

function importSlackDataToSheet(CHANNEL, oldest) {
  var msgData = importSlackMsg(CHANNEL, oldest);
  if(!msgData.ok){
    console.log(`In 'importSlackDataToSheet' ${channelCache[CHANNEL]}`);
    console.log(msgData.error);
    return 0;
  };
  var rows = [];

  msgData.messages.forEach( (message) => {
    rows.push([
      Utilities.formatDate(new Date(message.ts * 1000), timeZone, 'yyyy-MM-dd hh:mm:ss'),
      userCache[message.user],
      replaceText(message.text),
      `'${message.thread_ts || ''}`,
    ]);
  });

  if(rows.length == 0){ return 0; }

  var sheetName = channelCache[CHANNEL];
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  // sheet setups
  var lastRow = sheet.getLastRow();
  if(lastRow == 0){
    sheet.appendRow(['Timestamp', 'Author', 'Message', 'Thread TS']);
    lastRow = 1;
  }

  // append top messages
  sheet.getRange(1+lastRow, 1, rows.length, rows[0].length).setValues(rows);
  
  // get thread messages
  const threadTsList = getTsGreaterThanThreshold(sheet, oldest);
  var threadRows = [];

  threadTsList.forEach( (thread_ts) => {
    var threadlist = importSlackThread(CHANNEL, thread_ts, oldest);
    threadlist.messages.forEach( (message) => {
      threadRows.push([
        Utilities.formatDate(new Date(message.ts * 1000), timeZone, 'yyyy-MM-dd hh:mm:ss'),
        userCache[message.user],
        replaceText(message.text),
        `'${message.thread_ts || ''}`
      ])
    });
  })

  if(threadRows.length == 0){ return 0; }

  lastRow = sheet.getLastRow();
  sheet.getRange(1+lastRow, 1, threadRows.length, threadRows[0].length).setValues(threadRows);

  // clearn up
  sheet.getDataRange().removeDuplicates();
}

function importSlackMsg(CHANNEL, oldest){
  var apiUrl = `https://slack.com/api/conversations.history?channel=${CHANNEL}&inclusive=true&oldest=${oldest}`;
  const httpHeaders = {
    'Authorization': 'Bearer ' + TOKEN,
    "Content-Type": "application/json; charset=utf-8"
  };

  const options = {
    "method": "get",
    "headers": httpHeaders,
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(apiUrl, options);
  var data = JSON.parse(response.getContentText());

  return data;
}

function importSlackThread(CHANNEL, ts, oldest){
  var apiUrl = `https://slack.com/api/conversations.replies?channel=${CHANNEL}&ts=${ts}&oldest=${oldest}`;
  const httpHeaders = {
    'Authorization': 'Bearer ' + TOKEN,
    "Content-Type": "application/json; charset=utf-8"
  };

  const options = {
    "method": "get",
    "headers": httpHeaders,
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(apiUrl, options);
  var data = JSON.parse(response.getContentText());
  
  return data;
}

function getTsGreaterThanThreshold(sheet, ts_thres) {
  const timestamps = sheet.getRange('D:D').getValues().flat().filter((timestamp) => timestamp != "'");
  const filteredTimestamps = timestamps.filter((timestamp) => Number(timestamp.replace("'","")) > ts_thres);
  console.log(filteredTimestamps);
  return filteredTimestamps;
}