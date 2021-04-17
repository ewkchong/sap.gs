// the only sheet that should be edited by the script is the one with the name "Addresses from Shopify"
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Orders");
const mapSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Map");

/*

// produces column number for a given header (s) in row 1
function colNumber(s) {
  for (var i = 1; i <= sheet.getLastColumn(); i++) {
   var cell = sheet.getRange(1,i).getCell(1,1)
   
   if (cell.getValue() == s) {
     return cell.getColumn();
   }
  }
}


// define column for each set of information, with 1 being column A 
const orderNumColumn = colNumber('Order Number');

const addressColumn = colNumber('Address');
const unitColumn = colNumber('Unit');
const phoneColumn = colNumber('Phone Number');
const statusColumn = colNumber('Status');

*/
const orderNumColumn = 1;
const addressColumn = 21;
const unitColumn = 22;
const phoneColumn = 23;
const statusColumn = 19;

const height = sheet.getLastRow();
const length = sheet.getLastColumn();

const begList = sheet.createTextFinder('Pending Delivery').findNext().getRow() - 1;

function onOpen(e){
  SpreadsheetApp.getUi()
    .createMenu("Address Filler")
    .addItem("Fill Addresses", 'quickFill')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Reset")
                  .addItem("Reset All Orders", 'resetAllOrders')
                  .addItem("Reset To Do List", 'resetToDo')
                  .addItem("Reset Map", 'resetMap'))
    .addSeparator()
    .addItem("Open Sidebar", 'sidebar')
    .addItem("Make a Copy", 'saveAsSpreadsheet')
    .addToUi();
}
function quickFill() {
  let rowNum = sheet.getLastRow()
  let colNum = sheet.getLastColumn()
  let range = sheet.getRange(2, 1, rowNum, colNum)
  const username = 'USERNAME' ;
  const password = 'API_KEY' ;
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Please wait... this will take about 15 seconds");
  
  // Shopify API Request for the last 150 unfulfilled orders
  const orderListReq = UrlFetchApp.fetch('https://------.myshopify.com/admin/api/2020-10/orders.json?limit=150', {
            "method": "get",
            "headers": {
                "Authorization": "Basic " + Utilities.base64Encode(username + ':' + password)
                }
            });
  
  const orderList = JSON.parse(orderListReq);

  var orderRange = sheet.getRange(begList, orderNumColumn, height).getValues();  
  var addressRange = sheet.getRange(begList, addressColumn, height);  
  var unitRange = sheet.getRange(begList, unitColumn, height);  
  var phoneRange = sheet.getRange(begList, phoneColumn, height);
  
  var unitRangeValues = sheet.getRange(begList, unitColumn, height).getValues();  
  var addressRangeValues = sheet.getRange(begList, addressColumn, height).getValues();
 
  var ordersInRange = orderRange.map(num => orderList.orders.find(order => order.order_number == num[0]))
  
  function splitArray(someArray) {
    var tempArray = [];
    var arrayLength = someArray.length
    
    for (var index = 0; index < arrayLength; index += 1) {
      chunk = someArray.slice(index, index + 1);
      tempArray.push(chunk);
    }
    return tempArray;
  }

  function getAddress(order){
    if (order == undefined){
      return null;
    }
    else if (order.shipping_address == undefined){
      return null;
    }
    else{
      return order.shipping_address.address1 
      + ", " 
      + order.shipping_address.city 
      + ", " 
      + order.shipping_address.zip;
    }
  }
  function getAddress(order){
    if (order == undefined){
      return null;
    }
    else if (order.shipping_address == undefined){
      return null;
    }
    else{
      return order.shipping_address.address1 
      + ", " 
      + order.shipping_address.city 
      + ", " 
      + order.shipping_address.zip;
    }
  }
  
  function getPhoneNumber(order){
    
    function fixNum(s){
      var num = s.split("");
      var mapped = num.map(x => parseInt(x));
      var filtered = mapped.filter(n => Number.isInteger(n));
      
      if (filtered[0] == 1){
        filtered.shift();
      }
      
      return "(" 
      + filtered[0]
      + filtered[1]
      + filtered[2]
      + ")" + " " +
        + filtered[3]
      + filtered[4]
      + filtered[5]
      + "-"
      + filtered[6]
      + filtered[7]
      + filtered[8]
      + filtered[9]
    }
    if (order == undefined){
      return null;
    }
    else if (order.shipping_address == undefined){
      return null;
    }
    else if (order.shipping_address.phone == undefined) {
      return null;
    }
    else{
      return fixNum(order.shipping_address.phone);
    }
  }
  
  function getUnitInfo(order){
    if (order == undefined){
      return null;
    }
    else if (order.shipping_address == undefined){
      return null;
    }
    else if (order.shipping_address.address2 == undefined){
      return null;
    }
    else{
      return order.shipping_address.address2;
    }
  }
  
  function replace(arr, editArr) {
    for(i = 0; i < arr.length; i++){
      if(editArr[i] !== ""){
        arr[i] = editArr[i]
      }
    }
    return arr
  }
  
  function unsplitArray(arr){
    var unsplit = arr.map(element => element[0]);
    return unsplit;
  }
  
  var unitInfo = ordersInRange.map(order => getUnitInfo(order));
  var phoneNumbers = ordersInRange.map(order => getPhoneNumber(order));
  var shippingAddresses = ordersInRange.map(order => getAddress(order));
  
  var finalUnitInfo = replace(unitInfo, unsplitArray(unitRangeValues));
  var finalAddresses = replace(shippingAddresses, unsplitArray(addressRangeValues))
  
  addressRange.setValues(splitArray(finalAddresses));
  SpreadsheetApp.flush();

  unitRange.setValues(splitArray(finalUnitInfo));
  SpreadsheetApp.flush();

  phoneRange.setValues(splitArray(phoneNumbers));
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast("Done!");

  /*
  sheet.autoResizeColumns(1, length);
  mapSheet.autoResizeColumns(1, 5);
  mapSheet.setColumnWidth(6, 200);
  */
}

function resetAllOrders() {
  const source = SpreadsheetApp.getActiveSpreadsheet();
  const rangeToCopy = source.getSheetByName("Copy of All Orders").getRange('A:Y');

  rangeToCopy.copyTo(source.getSheetByName("All Orders").getRange('A:Y'));

}

function resetToDo() {
  const source = SpreadsheetApp.getActiveSpreadsheet();
  const rangeToCopy = source.getSheetByName("Copy of To Do List").getRange('A:F');

  rangeToCopy.copyTo(source.getSheetByName("To Do List").getRange('A:F'))

}

function resetMap() {
  const source = SpreadsheetApp.getActiveSpreadsheet();
  const rangeToCopy = source.getSheetByName("Copy of Map").getRange('A:H');

  rangeToCopy.copyTo(source.getSheetByName("Map").getRange('A:H'))
}

function resizeColumns() {
  mapSheet.setColumnWidth(2, 180);
  SpreadsheetApp.flush();

  mapSheet.autoResizeColumns(3, 5);
  SpreadsheetApp.flush();

  mapSheet.setColumnWidth(6, 200);
}

function sidebar() {
  const sb = HtmlService.createHtmlOutputFromFile('sidebar');
  SpreadsheetApp.getUi().showSidebar(sb);

}

function makeSpreadsheet(name, email) {
    const destFolder = DriveApp.getFolderById("1IJYlLuWxXDcMl1HeiWD4i5p9DIMLe1Mr");
    DriveApp.getFileById('1wp3E1HktWsOo77_KRUxde0PykWFiWgxjfTJjZHXnDj0').makeCopy("[" + name + "]" + " Delivery Mapper", destFolder)
      .addEditor(email);
}

function processForm(form) {
  if (form.name == "") {
    SpreadsheetApp.getUi().alert("Please enter a name and try again.");
  } else {
    if (form.email == "") {
      SpreadsheetApp.getUi().alert("Please enter an e-mail address and try again.");
    } else {
    makeSpreadsheet(form.name, form.email);
    SpreadsheetApp.getUi().alert("Spreadsheet Created!");
    }
  }
}

function saveAsSpreadsheet() {
  const dialog = HtmlService.createHtmlOutputFromFile('dialog');
  dialog.setWidth(400);
  dialog.setHeight(210);
  SpreadsheetApp.getUi().showDialog(dialog);
}

