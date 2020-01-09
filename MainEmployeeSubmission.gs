function onOpen(e) {
 var menu = SpreadsheetApp.getUi().createMenu('Scripts')
 menu.addItem('CopyMasterListForToday', 'CopyMasterListForToday')
 menu.addItem('Send Data To E', 'SendToAdmin')
 .addToUi(); 

};

function CopyMasterListForToday(){
    
   var dateObj = new Date();
   var month = dateObj.getMonth() + 1; //months from 1-12
   var day = dateObj.getDate();
   var year = dateObj.getFullYear();
   var date = month + "/" + day + "/" + year;
   CopyMasterList(date);
 
}
;
function DeleteEverySheetBut(DontDeleteThis){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsCount = ss.getNumSheets();
  var sheets = ss.getSheets();
    for (var i = 0; i < sheetsCount; i++){
      var sheet = sheets[i]; 
      var sheetName = sheet.getName();
      Logger.log(sheetName);
      if (sheetName.indexOf(DontDeleteThis.toString()) === -1){
        Logger.log("DELETE!");
        ss.deleteSheet(sheet);
      }
    } 
  
  };

function CopyMasterList(date){
   var source = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1syWWdkZvl5Qak_lyk5CGxkkeATAYtpZgZ8utzq3WTXk/edit#gid=723494406');
   var sheet = source.setActiveSheet(source.getSheetByName('TipOutSheet_Master'), true);
   var destination = SpreadsheetApp.getActiveSpreadsheet();
   sheet.copyTo(destination);
   
   
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsCount = ss.getNumSheets();
  var sheets = ss.getSheets();
  var sheet = sheets[0]; 
  ss.deleteSheet(sheet);
   

   destination.getActiveSheet().setName(date);
   sheets = destination.getSheets();
   DeleteEverySheetBut(date);
   
   
   
   destination.getActiveSheet().setName(date);
   destination.getRange('M2').activate();
   destination.getActiveRangeList().setValue(date);
   return 0

};



function SendToAdmin(){
   var source = SpreadsheetApp.getActiveSpreadsheet();
   var sourceSheet = source.getActiveSheet();
   var rangeData = sourceSheet.getDataRange();
   var searchRange = sourceSheet.getRange(1,1, rangeData.getLastRow(), rangeData.getLastColumn());
   var cell = searchRange.getValues();
   Logger.log(cell[1][12]);
   var date = cell[1][12].toString();
   Logger.log(date);
   var sheet = source.setActiveSheet(source.getSheetByName(date), true);
   var destination = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1syWWdkZvl5Qak_lyk5CGxkkeATAYtpZgZ8utzq3WTXk/edit#gid=723494406');
   var FindSheet = destination.getSheetByName(date);
   Logger.log(FindSheet);
   Logger.log(source.getSheetByName(date));
   if(FindSheet != null){
       Logger.log('Made');
      destination.deleteSheet(FindSheet);
      Logger.log('End');
     }
   sheet.copyTo(destination);
   destination = destination.setActiveSheet(destination.getSheetByName("Copy of " + date), true);
   destination.setName(date);
  
}
