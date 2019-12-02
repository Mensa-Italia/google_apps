//  1. Run > setup
//
//  2. Publish > Deploy as web app 
//    - enter Project Version name and click 'Save New Version' 
//    - set security level and enable service (most likely execute as 'me' and access 'anyone, even anonymously) 
//
//  3. Copy the 'Current web app URL' and post this in your form/script action 
//

var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service
  
function doPost(e){
  return handleResponse(e);
}
 
function doGet(e){
  return handleResponse(e);
}

function log(op, obj){
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var logs_sheet = doc.getSheetByName('Log');
    var logs_headRow = 1;
    var logs_nextRow = logs_sheet.getLastRow()+1; // get next row
    var logs_row = [new Date(), op, JSON.stringify(obj)]; 
    logs_sheet.getRange(logs_nextRow, 1, 1, logs_row.length).setValues([logs_row]);
}
 
function handleResponse(e) {
  // shortly after my original solution Google announced the LockService[1]
  // this prevents concurrent access overwritting data
  // [1] http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  // we want a public lock, one that locks for all invocations
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
   
  // If you are passing JSON in the body of the request uncomment this block
  jsonString = e.postData.getDataAsString();
  e.parameter = JSON.parse(jsonString);
  
  try {
    // next set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    
    
    log('log creation', new Date());
    var logs_sheet = doc.getSheetByName('Log');
    var logs_headRow = 1;
    var logs_nextRow = logs_sheet.getLastRow()+1; // get next row
    var logs_row = [new Date(), 'run', jsonString]; 
    logs_sheet.getRange(logs_nextRow, 1, 1, logs_row.length).setValues([logs_row]);
    log('end log creation', new Date());
    
    log('confs creation', new Date());
    // confs
    var confs_sheet = doc.getSheetByName('Conferenze');
    var confs_headRow = 1;
    var confs_headers = confs_sheet.getRange(1, 1, 1, confs_sheet.getLastColumn()).getValues()[0];
    var conf;
    var confs_nextRow;
    var confs_row;
    // loop through the header columns
    log('confs', e.parameter.confs);
    for (r in e.parameter.confs){
      conf = e.parameter.confs[r];
      log('conf', conf);
      confs_nextRow = confs_sheet.getLastRow()+1; // get next 
      confs_row = [];
      for (i in confs_headers){
        if (confs_headers[i] == "Timestamp"){ // special case if you include a 'Timestamp' column
          confs_row.push(new Date());
        } else if (confs_headers[i] == "Author"){ // special case if you include a 'Timestamp' column
          confs_row.push(e.parameter['author']);
        } else { // else use header name to get data
          confs_row.push(conf[confs_headers[i]]);
        }
      }
      confs_sheet.getRange(confs_nextRow, 1, 1, confs_row.length).setValues([confs_row]);
    }
    log('end confs creation', new Date());
    
    log('fairs creation', new Date());
    // confs
    var fairs_sheet = doc.getSheetByName('Fiere');
    var fairs_headRow = 1;
    var fairs_headers = fairs_sheet.getRange(1, 1, 1, fairs_sheet.getLastColumn()).getValues()[0];
    var fair;
    var fairs_nextRow;
    var fairs_row;
    // loop through the header columns
    log('fairs', e.parameter.fairs);
    for (r in e.parameter.fairs){
      fair = e.parameter.fairs[r];
      log('fair', fair);
      fairs_nextRow = fairs_sheet.getLastRow()+1; // get next 
      fairs_row = [];
      for (i in fairs_headers){
        if (fairs_headers[i] == "Timestamp"){ // special case if you include a 'Timestamp' column
          fairs_row.push(new Date());
        } else if (fairs_headers[i] == "Author"){ // special case if you include a 'Timestamp' column
          fairs_row.push(e.parameter['author']);
        } else { // else use header name to get data
          fairs_row.push(fair[fairs_headers[i]]);
        }
      }
      fairs_sheet.getRange(fairs_nextRow, 1, 1, fairs_row.length).setValues([fairs_row]);
    }
    log('end fair creation', new Date());
    
    log('brain creation', new Date());
    // confs
    var brains_sheet = doc.getSheetByName('Brain');
    var brains_headRow = 1;
    var brains_headers = brains_sheet.getRange(1, 1, 1, brains_sheet.getLastColumn()).getValues()[0];
    var brain;
    var brains_nextRow;
    var brains_row;
    // loop through the header columns
    log('brains', e.parameter.brain);
    for (r in e.parameter.brain){
      brain = e.parameter.brain[r];
      log('brain', brain);
      brains_nextRow = brains_sheet.getLastRow()+1; // get next 
      brains_row = [];
      for (i in brains_headers){
        if (brains_headers[i] == "Timestamp"){ // special case if you include a 'Timestamp' column
          brains_row.push(new Date());
        } else if (brains_headers[i] == "Author"){ // special case if you include a 'Timestamp' column
          brains_row.push(e.parameter['author']);
        } else { // else use header name to get data
          brains_row.push(brain[brains_headers[i]]);
        }
      }
      brains_sheet.getRange(brains_nextRow, 1, 1, brains_row.length).setValues([brains_row]);
    }
    log('end brain creation', new Date());
    
    log('test creation', new Date());
    var test_sheet = doc.getSheetByName('Test');
    var test_headRow = 1;
    var test_nextRow = test_sheet.getLastRow()+1; // get next row
    var test_row = [new Date(), e.parameter['author'], e.parameter['tests'], e.parameter['tests_note']]; 
    test_sheet.getRange(test_nextRow, 1, 1, test_row.length).setValues([test_row]);
    log('end test creation', new Date());
    
    
    // return json success results
    return ContentService
      .createTextOutput(JSON.stringify({"result":"success"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e){
    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}
 
function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    SCRIPT_PROP.setProperty("key", doc.getId());
}
