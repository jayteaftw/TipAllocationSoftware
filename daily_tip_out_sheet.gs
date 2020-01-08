function onOpen(e) {
 var menu = SpreadsheetApp.getUi().createMenu('Scripts')
 menu.addItem('TransferData', 'TransferData')
 //.addSeparator()
 //.addSubMenu(SpreadsheetApp.getUi().createMenu('Sub Menu')
 //.addItem('One sub-menu item', 'subFunction1')
 //.addItem('Another sub-menu item', 'subFunction2'))
 .addToUi(); 
};


function TransferData(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeData = sheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow();
  var searchRange = sheet.getRange(1,1, lastRow, lastColumn);
  var cell = searchRange.getValues();
  for(i = 1; i < lastColumn; i++)
  {
    if(cell[i][1] == true)
      addTips(cell[i][0],cell[1][12],cell[i][3],cell[i][4],cell[i][5],cell[i][6],cell[i][7],cell[i][8],cell[i][9],cell[i][10],cell[i][11],cell[i][2]);
  }
  return 0;
};


function addTips(name, Date, Cash, Charge, Total, Host, Expo, Bar, Kit, Out, Home, Pos) {
  //var sheet = SpreadsheetApp.getActive()
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1csFfCG1eTX6RiWhoamn9-oUxYhR3t86oTO_4wzZDWd0/edit#gid=877735631');
  var nameofSheet = sheet.getSheetByName(name);
   if (nameofSheet == null) { //If Employee doesnt have a sheet, then new sheet with his or her name will be added plus tip out.
    Logger.log('Need new Employee');
    sheet.setActiveSheet(sheet.getSheetByName('Employee_Tip_Out_Master'), true);
    sheet.getRange('A1').activate();
    sheet.duplicateActiveSheet();
    sheet.getCurrentCell().setValue(name);
    sheet.getActiveSheet().setName(name);
    addUserToEndYear(name);
    sheet.appendRow(['', Date, Cash, Charge, Total, Host, Expo, Bar, Kit, Out, Home, Pos]);
   }
   else {
     sheet.setActiveSheet(nameofSheet, true);
     sheet.appendRow(['', Date, Cash, Charge, Total, Host, Expo, Bar, Kit, Out, Home, Pos]);
   }
   UpdateYearly(name);
   return 0;
};




function UpdateYearly(name){
   var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1csFfCG1eTX6RiWhoamn9-oUxYhR3t86oTO_4wzZDWd0/edit#gid=877735631');
   var sheet = ss.setActiveSheet(ss.getSheetByName(name), true);
   var rangeData = sheet.getDataRange();
   var lastColumn = rangeData.getLastColumn();
   var lastRow = rangeData.getLastRow();
   var searchRange = sheet.getRange(1,1, lastRow, lastColumn);
   var cell = searchRange.getValues();
   sendToYearly(name, cell[1][2],cell[1][3],cell[1][4], cell[1][5], cell[1][6], cell[1][7], cell[1][8], cell[1][9], cell[1][10]);
   return 0;
};

function sendToYearly(name, Cash, Charge, Total, Host, Expo, Bar, Kit, Out, Home){
  var value = [name, Cash, Charge, Total, Host, Expo, Bar, Kit, Out, Home];
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1po8BFkXjKHjyyxA_3eorUThfBo4RVpMfEMIVf8Nif0Y/edit#gid=1629813004');
  var dateObj = new Date();
  var year = dateObj.getFullYear();
  var sheet = ss.setActiveSheet(ss.getSheetByName(year), true);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var SheetArray = sheet.getSheetValues(2, 1, 10, 10);
  Logger.log(SheetArray[0][0]);
  Logger.log(value[2]);
  for(i =1; i < lastColumn; i++){

     if(SheetArray[i - 1][0] == name) {
       
       for(j = 1; j < lastColumn; j++){
       sheet.getRange(i + 1,j+1,lastRow,lastColumn).activate();
       sheet.getCurrentCell().setValue(value[j]);
       }
       i = lastColumn;
     }
     
    }
    
    return 0;
  
};

function addUserToEndYear(name){
   var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1po8BFkXjKHjyyxA_3eorUThfBo4RVpMfEMIVf8Nif0Y/edit#gid=0');
   var dateObj = new Date();
   var year = dateObj.getFullYear();
   sheet.setActiveSheet(sheet.getSheetByName(year), true);
   sheet.appendRow([name]);
};






