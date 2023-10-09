function FetchMonthlySession(month) {
  // Get all the values of the sheets we need
  var ss1 = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet
  var WorkSheet = ss1.getSheetByName('2023'); // The main sheet (Sessions organization)
  var ss2 = SpreadsheetApp.openById("1aUU6i5Sh93OKUYClxE8445JMC4qAxCdHDHI2Gj6pI-0");
  var SessionsSource = ss2.getSheetByName('Sessions organization');
  
  var StartDate = new Date('2023-' + month + '-01'); // Construct the start date
  var EndDate = new Date('2023-' + month + '-31'); // Construct the end date
  var i = 8;
  var startline = 0;
  var endline = 0;
  var Sessions_List = [];
  
  // Find the start line
  while (new Date(SessionsSource.getRange(i, 3).getValue()) >= StartDate) {
    i++;
  }
  startline = i-1;

  while (new Date(SessionsSource.getRange(i, 3).getValue()) <= EndDate) {
    i--;
  }
  endline = i+1;
   
   for (i = startline; i >= endline; i--) {
    if (!judge_week(i, SessionsSource)) {
      Sessions_List.push(SessionsSource.getRange(i, 9).getValue());
    }
  }

  return Sessions_List;
}



function judge_week(i,source){
  var week = source.getRange(i,1).getValue();
  var regExp = new RegExp ("^WEEK");
  var res = regExp.exec(week);
  if(res!=null)
  {
     return true;
  }
  return false;
}

function myFunction() {
  // Get all the values of the sheets we need
  var ss1 = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet
  var WorkSheet = ss1.getSheetByName('2023'); // The main sheet (Sessions organization)
  var i;
  var j;
  var Find = [];
  var SessionCount = 0;
  var TraineeCount_Target = 0;
  var TraineeCount_PTS = 0;
  var TraineeCount_NOPTS = 0;

  for (i = 5; i <= 13; i++) {
    if (WorkSheet.getRange(48, i).getValue() === "") {
      var Session_Array = FetchMonthlySession(i - 4);
      SessionCount = 0; // Reset counts for each month
      TraineeCount = 0;
      for (j = 0; j < Session_Array.length; j++) {
        Find = FindRecord(Session_Array[j]);
        if (Find[0] > 0) {
          SessionCount++;
          TraineeCount_Target += Find[0];
          TraineeCount_PTS += Find[1];
          TraineeCount_NOPTS += Find[2];
        }
      }
      if (SessionCount > 0) {
          WorkSheet.getRange(48, i).setValue(SessionCount);
          WorkSheet.getRange(49, i).setValue(TraineeCount_Target);
          WorkSheet.getRange(50, i).setValue(TraineeCount_PTS);
          WorkSheet.getRange(51, i).setValue(TraineeCount_NOPTS);
           Find = [];
           SessionCount = 0;
           TraineeCount_Target = 0;
           TraineeCount_PTS = 0;
          TraineeCount_NOPTS = 0;
      }
    }
    
  }
}


function FindRecord(session_number) {
  var ss = SpreadsheetApp.openById("1mOEB6b17LMxAOv4agEPwxXs8eLO1kWCmgbv_bhsdmic");
  var recordSheet = ss.getSheetByName('Session_Completed');
  var ss2 = SpreadsheetApp.openById("1aUU6i5Sh93OKUYClxE8445JMC4qAxCdHDHI2Gj6pI-0");
  var SessionsSheet = ss2.getSheetByName('Sessions organization');
  var recordData = recordSheet.getDataRange().getValues(); // Get all data in the sheet
  var sessionsData = SessionsSheet.getDataRange().getValues(); // Get all data in the sheet
  
  var recordList = [];
  
  for (var i = 1; i < recordData.length; i++) { // Start from row 2 (assuming row 1 contains headers)

    if (recordData[i][0] == session_number) { 
       
       for (var j = 7; j < sessionsData.length; j++) { // Start from row 8 
        if (sessionsData[j][8] == session_number) { 

            // Get the cell value
            var cellValue = sessionsData[j][11];
            
            // Split the cell value by the word "of"
            var parts = cellValue.split("of");
            
            // Extract the second part, which should be "12"
            var valueOnRight = parts[1].trim();
            // Convert the extracted string to an integer
            var valueOnRightAsInteger = parseInt(valueOnRight, 10);

            recordList[0] = valueOnRightAsInteger;
          break;
        }
       }
       recordList[1]=recordData[i][1];
       recordList[2]=recordData[i][2];  
       break;
    }
  }

  return recordList;
}

