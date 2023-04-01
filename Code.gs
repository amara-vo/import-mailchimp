/* 
* Modified by Amara Vo for RelishCareers
* This program gets the most recently uploaded CSV file in a specified Google folder, and then adds it to a specified Google spreadsheet
*/
function onSubmitForm(){
  
    
  /*
  * REPLACE AS NEEDED: (Folder ID is found at the end of URL)
  * csvFolder = Folder where CSV reports are saved
  */
  var csvFolder = DriveApp.getFolderById("1YqkpWgcPXAvUQJHNTt-EulErYnQJuE_X");                    // Example: "[insert ID here]" KEEP THE QUOTATION MARKS
  
  /*
  * REPLACE AS NEEDED: (Spreadsheet ID is found at the end of URL)
  * spreadsheet = Spreadsheet that holds the sheet that will be updated with new data
  */
  var spreadsheet = SpreadsheetApp.openById("17hw_xPptw_KYbfuUJYwkyGbXIoTUxRFQd3Elm68-7f8");       // Example: "[insert ID here]"
  
  /*
  * REPLACE AS NEEDED (Keep quotations)
  * sheet = Sheet within the Spreadsheet where cells will be updated with new data
  */
  var sheet = spreadsheet.getSheetByName("2020 Student Digest By Campaign");                                                // Example: "[insert Sheet Name here]"
  
  var csvFile = csvFolder.getFilesByName(getNewestFileInFolder(csvFolder)); 
  
  getCSV(csvFile, sheet);
  
}

/*
* Retrieves specific data cells from MailChimp CSV.
* Returns a dictionary of relevant CSV data.
* Source: https://stackoverflow.com/questions/26854563/how-to-automatically-import-data-from-uploaded-csv-or-xls-file-into-google-sheet
*/
function getCSV(fi, ss){
    var date, month, digestName, audience, sent, opens, openRate, totalClicks, uniqueClicks, unsubs;
  
    if ( fi.hasNext() ) {                          
      var file = fi.next();
      var csv = file.getBlob().getDataAsString();
      var csvFullData = CSVToArray(csv);
      
      for ( var i=0; i < csvFullData.length; i++ ) {  // loop through csvFullData array
        
        // get Title
        if (csvFullData[i].indexOf("Title:") != (-1)) {
          var regex = /\(([^)]+)\)/;   
          digestName = csvFullData[i][1].split(' (')[0];
          audience = csvFullData[i][1].match(regex)[1];
          
          //csvData.push(digestName);
          //csvData.push(audience);
          
           //csvData["Name"] = digestName;
          continue;
        }
        
        // get Date
        if (csvFullData[i].indexOf("Delivery Date/Time:") != (-1)) {
          var fullDate = new Date (Date.parse(csvFullData[i][1]));
          date = fullDate.toDateString();
          month = new Intl.DateTimeFormat("en-US", { month: "long" }).format(fullDate);
          
          //csvData.push(date);
          //csvData.push(month);
          continue;
        }
        
        // get # of Recipients
        if (csvFullData[i].indexOf("Total Recipients:") != (-1)) {
          sent = csvFullData[i][1];
          
          //csvData.push(sent);
          continue;
        }
        
        // get # of Opens
        if (csvFullData[i].indexOf("Recipients Who Opened:") != (-1)) {
          opens = csvFullData[i][1].split(' ')[0];
          
          //csvData.push(opens);
          continue;
        }
        
        // get # of Total Clicks
        if (csvFullData[i].indexOf("Total Clicks:") != (-1)) {
          totalClicks = csvFullData[i][1];
          
          //csvData.push(totalClicks);
          continue;
        }
        
        // get # of Unique Clicks
        if (csvFullData[i].indexOf("Recipients Who Clicked:") != (-1)) {
          uniqueClicks = csvFullData[i][1].split(' ')[0];
          
          //csvData.push(uniqueClicks);
          continue;
        }
        
        // get # of Unsubscribers
        if (csvFullData[i].indexOf("Total Unsubs:") != (-1)) {
          unsubs = csvFullData[i][1];
          
          //csvData.push(uniqueClicks);
          continue;
        }
        
        
      }
      
      ss.appendRow([Utilities.formatDate(new Date(), "GMT-4", "MM/dd, HH:mm aaa"), "date", digestName, audience, sent, opens, "open rate", totalClicks, uniqueClicks, "CTR", "engagement rate", unsubs, "unsubscribe rate"]);
      // Format date
      ss.getRange(ss.getLastRow(), 2).setValue(date).setNumberFormat("M/DD/YYYY");
      
      // Open Rate formula: opens/sent
      ss.getRange(ss.getLastRow(), 7).setFormula("=F"+ss.getLastRow()+"/E"+ss.getLastRow());
      
      // CTR formula: uniqueClicks/sent
      ss.getRange(ss.getLastRow(), 10).setFormula("=I"+ss.getLastRow()+"/E"+ss.getLastRow());
      
      // Engagement Rate formula: uniqueClicks/open
      ss.getRange(ss.getLastRow(), 11).setFormula("=I"+ss.getLastRow()+"/F"+ss.getLastRow());
      
      // Unsubsribe Rate formula: unscribers/sent
      ss.getRange(ss.getLastRow(), 13).setFormula("=L"+ss.getLastRow()+"/E"+ss.getLastRow());
      
    }
  
};


/*
* Returns file name of most recently uploaded file in csvFolder
* Source: https://stackoverflow.com/questions/28323703/get-the-newest-file-in-a-google-drive-folder
*/
function getNewestFileInFolder(csvFolder) {
  var arryFileDates,file,fileDate,files,folder,folders,newestDate,newestFileID,objFilesByDate, newestFileName;

  arryFileDates = [];
  objFilesByDate = {};

  folder = csvFolder;     // pass in CSV folder
  
  files = folder.getFilesByType(MimeType.CSV);     // get only CSV files
  fileDate = "";
  
  while (files.hasNext()) {  //If no files are found then this won't loop
    file = files.next();
    //Logger.log('xxxx: file data: ' + file.getLastUpdated());
    //Logger.log('xxxx: file name: ' + file.getName());
    //Logger.log('xxxx: mime type: ' + file.getMimeType())
    //Logger.log(" ");
    
    fileDate = file.getDateCreated();
    objFilesByDate[fileDate] = file.getName(); // Create an object of file names by file Name
    
    arryFileDates.push(file.getDateCreated());
  }
  
  if (arryFileDates.length === 0) {         //The length is zero so there is nothing
    //to do
    return;
  }
  
  // sort by most recent
  arryFileDates.sort(function(a,b){return b-a});
  //Logger.log(arryFileDates);
  
  newestDate = arryFileDates[0];
  //Logger.log('Newest date is: ' + newestDate);
    
  newestFileName = objFilesByDate[newestDate];
  Logger.log('newestFileName: ' + newestFileName);
  
  return newestFileName;
  
};

// http://www.bennadel.com/blog/1504-Ask-Ben-Parsing-CSV-Strings-With-Javascript-Exec-Regular-Expression-Command.htm
// This will parse a delimited string into an array of
// arrays. The default delimiter is the comma, but this
// can be overriden in the second argument.

function CSVToArray( strData, strDelimiter ) {
  // Check to see if the delimiter is defined. If not,
  // then default to comma (Mailchimp uses comma).
  strDelimiter = (strDelimiter || ",");

  // Create a regular expression to parse the CSV values.
  var objPattern = new RegExp(
    (
      // Delimiters.
      "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

      // Quoted fields.
      "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

      // Standard fields.
      "([^\"\\" + strDelimiter + "\\r\\n]*))"
    ),
    "gi"
  );

  // Create an array to hold our data. Give the array
  // a default empty first row.
  var arrData = [[]];

  // Create an array to hold our individual pattern
  // matching groups.
  var arrMatches = null;

  // Keep looping over the regular expression matches
  // until we can no longer find a match.
  while (arrMatches = objPattern.exec( strData )){

    // Get the delimiter that was found.
    var strMatchedDelimiter = arrMatches[ 1 ];

    // Check to see if the given delimiter has a length
    // (is not the start of string) and if it matches
    // field delimiter. If id does not, then we know
    // that this delimiter is a row delimiter.
    if (
      strMatchedDelimiter.length &&
      (strMatchedDelimiter != strDelimiter)
    ){

      // Since we have reached a new row of data,
      // add an empty row to our data array.
      arrData.push( [] );

    }

    // Now that we have our delimiter out of the way,
    // let's check to see which kind of value we
    // captured (quoted or unquoted).
    if (arrMatches[ 2 ]){

      // We found a quoted value. When we capture
      // this value, unescape any double quotes.
      var strMatchedValue = arrMatches[ 2 ].replace(
        new RegExp( "\"\"", "g" ),
        "\""
      );

    } else {

      // We found a non-quoted value.
      var strMatchedValue = arrMatches[ 3 ];

    }

    // Now that we have our value string, let's add
    // it to the data array.
    arrData[ arrData.length - 1 ].push( strMatchedValue );
  }

  // Return the parsed data.
  return( arrData );
};
