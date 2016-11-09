//
// Version: 0.0.2
// Developer: gertheine@aol.com
// Date: 2016.10.30
//
function myFunction() {
  IMPORTRANGE("https://docs.google.com/document/d/1nYMuyZ89LxkrQYfTP2XocW9tWPzzGa0-kZUdbSnVR4k", "sheet1!N2:O2")

IMPORTRANGE(N2,O2)
}

var country_pref = {};
var MAIN_SP_NAME = "PHP Project";
var MAIN_FOLDER = "MAIN";

//
// Open the Spreadsheet
//
function onOpen(e){
  
  toast("Welcome to this script. from gertheine@aol.com");
  fill_case_id();
  
}

function fill_case_id(){
  var mainSheet = SpreadsheetApp.getActiveSheet();
  var data = mainSheet.getDataRange().getValues();
  var rows = data.length;
  
  var key = "";
  for(var i=rows-1;i>0;i--){
    if( data[i][1-1] == "" || data[i][1-1] == undefined || data[i][1-1].length != 7) // XXXX-XX
    {
      key = makeid(6);
      //mainSheet.getRange(i+1, 1).setValue(key.substring(0,4)+"-"+key.substring(4,2));
      mainSheet.getRange(i+1, 1).setValue(key.substring(0,4)+"-"+key.substring(4,6));
    }
  }
}
//
// Insert the new Row
//
function onChange(e){
  //toast(e.changeType.toString());
  
  //if(e.changeType.toString() == "EDIT")
  {
    fill_case_id();
  }
}

function updateSpreadsheet(private_company, contact_claimant_country){
  MAIN_SP_NAME = SpreadsheetApp.getActive().getName();
  private_company = "COMPANY"; //"PRIVATE PERSON";
  contact_claimant_country = "be";
  
  initCountryPrefix();
  
  var country = country_pref[contact_claimant_country];
  var filename = contact_claimant_country + "_" + MAIN_SP_NAME;
  
  //
  // Checking pre-exist Spreadsheet
  //
  var folders = DriveApp.getFoldersByName(MAIN_FOLDER);
  var folder;
  if(folders.hasNext()){
    
    folder = folders.next();
    folders = folder.getFoldersByName(country);
    if(folders.hasNext()){
      folder = folders.next();
      folders = folder.getFoldersByName(private_company);
      if(folders.hasNext()){
        
        folder = folders.next();
        
      }else
      {
        alert ("no folder:" + private_company);
        return;
      }
    }
    else
    {
      alert ("no folder:" + country);
      return;
    }
  }else{
    alert ("no folder:" + MAIN_FOLDER);
    return;
  }
  
  var files = folder.getFilesByName(filename);
  //
  // Move to Trash pre-exist file
  //
  toast("Trahsh pre-existed files...");
  while(files.hasNext()){
    
    files.next().setTrashed(true);
  }
  toast("Copying template file...");
  //
  // Main spreadsheet
  //
  var file = DriveApp.getFilesByName(MAIN_SP_NAME).next();
  //
  // copy
  //
  file.makeCopy(filename, folder);
  //
  // Remove un-corresponding rows
  //
  toast("Removing unnecessary rows...");
  files = folder.getFilesByName(filename);
  if(files.hasNext()){
    var tempSpreadsheet = SpreadsheetApp.open(files.next());
    var tempSheet = tempSpreadsheet.getSheets()[0];
    
    
    var data = tempSheet.getDataRange().getValues();
    var rows = data.length;
    
    for(var i=rows;i>0;i--){
      if( data[i-1][2-1].toLowerCase() == private_company.toLowerCase() &&
         data[i-1][15-1].toLowerCase() == contact_claimant_country.toLowerCase()) {
        continue;
      }
      tempSheet.deleteRow(i);
    }
    toast("Success!");
  }else{
    toast("Failed.");
  }
}

function initCountryPrefix(){
  country_pref["be"] = "BELGIUM";
  country_pref["us"] = "USA";
  country_pref["ge"] = "GERMANY";
  country_pref["fr"] = "FRANCE";
  country_pref["sp"] = "SPAIN";
}

function makeid(length)
{
    var text = "";
    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    for( var i=0; i < length; i++ )
        text += possible.charAt(Math.floor(Math.random() * possible.length));

    return text;
}

function alert(str){
  SpreadsheetApp.getUi().alert(str);
}
function toast(str) {
  var myToast = new Toaster(str, "Information");
  myToast.display();
}

/**
 * "Class" Toaster
 *
 * From http://stackoverflow.com/a/33552904/1677912
 *
 * Wrapper for Spreadsheet.toast() with support for multi-line messages.
 *
 * Constructor:    new Toaster( message, title, timeoutSeconds );
 *
 * @param message         {String}    Toast message, possibly with newlines (`\n`)
 * @param title           {String}    (optional) Toast title
 * @param timeoutSeconds  {Number}    (optional) Duration of display, default 3s
 *
 * @returns {Toaster}                 Toaster instance.
 */
var Toaster = function(message, title, timeoutSeconds) {
  if (typeof message == 'undefined')
    throw new TypeError( "missing message" );

  this.message = this.parseMessage(message);
  this.title = title || '';
  this.timeoutSeconds = timeoutSeconds || 3;
  this.ss = SpreadsheetApp.getActiveSpreadsheet();
};

/**
 * Display Toaster message using previously set parameters.
 */
Toaster.prototype.display = function() {
  this.ss.toast(this.message,this.title,this.timeoutSeconds);
}

/**
 * This is where the magic happens. Prepares multi-line messages for display.
 *
 * @param {String} msg    Toast message, possibly with newlines (`\n`)
 *
 * @returns{String}       Message, ready to display.
 */
Toaster.prototype.parseMessage = function( msg ) {
  var maxWidth = 52;             // Approx. number of non-breaking spaces required to span toast popup.
  var knob = 1.85;               // Magical approx. ratio of avg char width : non-breaking space width
  var parsedMessage = '';

  var lines = msg.split('\n');   // Break lines at newline chars

  // Rebuild message with padded lines
  for (var i=0; i<lines.length; i++) {
    var len = lines[i].length;
    // Build padding string of non-breaking spaces sandwiched with normal spaces.
    var padding = ' '
                + len < (maxWidth / knob) ?
                  Array(Math.floor(maxWidth-(lines[i].length * knob))).join(String.fromCharCode(160)) + ' ' : '';
    parsedMessage += lines[i] + padding;
  }
  return parsedMessage;
}