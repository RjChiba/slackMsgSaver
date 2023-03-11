
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
      return userCache[userId];
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
    .at(date)
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
    console.log(userData.error);
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

  channelData.channels.forEach( (channel) => {
    channelCache[channel.id] = channel.name;
  })

}

function importSlackDataToSheet(CHANNEL, oldest) {
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

  if(!data.ok){
    console.log(data.error);
    return 0;
  };

  var headers = ['Timestamp', 'Author', 'Message', 'Thread ID'];
  var rows = [];

  data.messages.forEach(function(message) {
    rows.push([
      new Date(message.ts * 1000), // msec -> sec
      userCache[message.user],
      replaceText(message.text),
      message.thread_ts || '',
    ]);
  });

  // Insert the headers and rows into the sheet
  if(rows.length !== 0){
    // Create a new sheet and insert the message data into it
    var sheetName = channelCache[CHANNEL];
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

    let lastRow = sheet.getLastRow();

    if(lastRow == 0){
      sheet.appendRow(headers);
      lastRow = 1;
    }

    sheet.getRange(1+lastRow, 1, rows.length, rows[0].length).setValues(rows);
  }
}
