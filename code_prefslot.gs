//Returns all data stored in sheetName as an array
function getData(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  var data = sheet.getDataRange().getValues();
  return data;
}

//Writes val into the cell with row no. = r and column no. = c 
function Write(sheetName, r, c, val){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  sheet.getRange(r, c).setValue(val)
}

//Returns the col. number/Index of the column whose header has the name colName
function getIndexByColumnName(sheetName, colName) {
    const wk = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const headers = wk.getRange("A1:1").getValues()[0]
    const colNum = headers.indexOf(colName)
    return colNum
}

//Deletes the sheet with name sheetName
function deleteSheet(sheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var itt = ss.getSheetByName(sheetName);
  if (itt) {
    ss.deleteSheet(itt);
  }
}

//Makes a copy of the sheet Form Responses first and then updates it such that it no longer have duplicate entries corresponding to the same name and email address. Only the first response is considered in such a case. 
function removeDuplicates() {
  deleteSheet("NoDuplicates")
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = source.getSheetByName("Form Responses");
  sheet.copyTo(source).setName("NoDuplicates");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NoDuplicates");
  sheet.deleteColumn(1);
  var data = sheet.getDataRange().getValues();
  var numRows = sheet.getLastRow();
  var numCols = sheet.getLastColumn(); 
  var uniqueData = [];
  var uniqueRows = [];

  for (var i = 0; i < numRows; i++)
  {
        var duplicate = false;
    for (var j = 0; j < uniqueData.length; j++) {
      var isEqual = true;
      for (var k = 0; k < 2; k++) {            //Taking duplicate entry if BOTH name and email are same
        if (data[i][k] !== uniqueData[j][k]) {
          isEqual = false;
          break;
        }
      }
      if (isEqual) {
        duplicate = true;
        break;
      }
    }
    if (!duplicate) {
      uniqueData.push(data[i]);               //If current row is not present in uniqueData, push it to uniqueData
      uniqueRows.push(i + 1);
    }
  }
  if (uniqueData.length < numRows) {
    sheet.clearContents();
    sheet.getRange(1, 1, uniqueData.length, numCols).setValues(uniqueData);               //Updating the sheet to only contain the uniqueData entries
    
  } 

}


function locationIndex(location){
  var locationarr = ["Delhi", "Mumbai", "Bangalore", "Chennai", "Kolkata"];
  for(var i = 0; i < 5; i=i+1){
    if(location == locationarr[i])
      return i;
  }

}

function slotsIndex(slot){
  var preferenceArr = ["No Preference", "10:00 - 10:30 AM", "10:30 - 11:00 AM", "11:00 - 11:30 AM", "11:30 AM - 12:00 PM", "12:00 - 12:30 PM", "12:30 - 1:00 PM"];
  for(var i = 0; i < 7; i++){
    if(slot == preferenceArr[i])
      return i;
  }
}

function TokenAssigner(sheetName){
   var total_slots = 6;
   var total_locations = 5;

   var emailData = getData(sheetName);
   var headerRow = emailData.shift(); 
   var tokens = [[0, 0, 0, 0, 0, 0],          
                  [0, 0, 0, 0, 0, 0],
                  [0, 0, 0, 0, 0, 0],
                  [0, 0, 0, 0, 0, 0],
                  [0, 0, 0, 0, 0, 0]];

   var locationData = [[],[],[],[],[]];         
   //locationData[i] contains all data corresponding to location[i]. Helps in implementing location-specific queues

   
   
   var row_num = 1;
   emailData.forEach(function (row) {

    var location = locationIndex(row[getIndexByColumnName(sheetName, "Preferred Location")]);
    locationData[location].push([row_num].concat(row));
    row_num = row_num + 1;
   });
  //Adding row num as well to locationData to later on help in writing Assigned Token at that row in original sheet. Not required if we just want to send emails. 

   var location = 0;
   for(location = 0; location < total_locations; location = location + 1){
    var data = locationData[location];
    var unallocated = []
    
    //If people who have applied for that particular location are less in number than the slots available. In this case everyone is assigned a slot. 
    if(data.length <= total_slots){
      
      for(var i = 0; i < data.length; i=i+1){

        var preferred_slot = slotsIndex(data[i][getIndexByColumnName(sheetName, "Preferred Slot") + 1]) - 1;
        if(preferred_slot == -1) //No preference, move to unallocated
          unallocated.push(data[i]);
        
        else if(tokens[location][preferred_slot] == 0){
          tokens[location][preferred_slot] = preferred_slot + 1;
          Write(sheetName, data[i][0] + 1, getIndexByColumnName(sheetName, "Assigned Token")+1, preferred_slot+1);
        }
        else
          unallocated.push(data[i]);
      }

      for(var i = 0; i < unallocated.length; i= i+1){
        for(var slot = 0; slot < total_slots; slot = slot + 1){
          if( tokens[location][slot] == 0){
            tokens[location][slot] = slot + 1;
            Write(sheetName , unallocated[i][0] + 1, getIndexByColumnName(sheetName, "Assigned Token")+1, slot+ 1)
            break;
          }
        }
      }

    }

    //If people who have applied for that particular location are more in number than the slots available. In this case only the first 6 persons are assigned a slot. 
    else{
      for(var i = 0; i < total_slots; i=i+1){

        var preferred_slot = slotsIndex(data[i][getIndexByColumnName(sheetName, "Preferred Slot") + 1]) - 1;
        if(preferred_slot == -1)
          unallocated.push(data[i]);
        else if(tokens[location][preferred_slot] == 0){
          tokens[location][preferred_slot] = preferred_slot + 1;
          Write(sheetName, data[i][0] + 1, getIndexByColumnName( sheetName, "Assigned Token")+1, preferred_slot+1);
        }
        else
          unallocated.push(data[i]);
      }

      for(var i = 0; i < unallocated.length; i= i+1){
        for(var slot = 0; slot < total_slots; slot = slot + 1){
          if( tokens[location][slot] == 0){
            tokens[location][slot] = slot + 1;
            Write(sheetName, unallocated[i][0] + 1, getIndexByColumnName(sheetName, "Assigned Token")+1, slot+ 1);
            break;
          }
        }
      }

      for(var i = total_slots; i < data.length; i = i+1)
        Write(sheetName, data[i][0] + 1, getIndexByColumnName(sheetName, "Assigned Token")+1, "-");
    }

    
}

}


//main function
function sendEmails() {
  var sheetName = "NoDuplicates";
  removeDuplicates();
  Write(sheetName, 1, 6, "Assigned Token");
  TokenAssigner(sheetName);
  var emailData = getData(sheetName);
  var headerRow = emailData.shift();
 
  var templateData = getData("Templates");
  var emailSubject = templateData[1][0]; //Cell A2 (contains the email subject)
  var emailBody1 = templateData[4][0]; //Cell A5 (contains the email body if some token is assigned)
  var emailBody2 = templateData[7][0]; //Cell A8 (contains the email body if NO token is assigned)
  var row_num = 1;
  var timeslots = ["-", "10:00 - 10:30 AM", "10:30 - 11:00 AM", "11:00 - 11:30 AM", "11:30 AM - 12:00 PM", "12:00 - 12:30 PM", "12:30 - 1:00 PM"];
  emailData.forEach(function (row) {
    var email = row[getIndexByColumnName(sheetName, "Email Address")]
    var name = row[getIndexByColumnName(sheetName, "Name")]
    var location = row[getIndexByColumnName(sheetName, "Preferred Location")]
    var preferred_slot = row[getIndexByColumnName(sheetName, "Preferred Slot")]
    var assigned_token = row[getIndexByColumnName(sheetName, "Assigned Token")]
    var curr_date = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy")
    if(assigned_token != "-"){
    time = timeslots[assigned_token]

    MailApp.sendEmail({
     to: email,
     subject: emailSubject,
     htmlBody: Utilities.formatString(emailBody1, name, curr_date, time, assigned_token, location) ,

   });
    }
    else{
      time = "Not Available";
    MailApp.sendEmail({
     to: email,
     subject: emailSubject,
     htmlBody: Utilities.formatString(emailBody2, time, assigned_token),

   });
    }
    
    row_num = row_num + 1;
  })
};
