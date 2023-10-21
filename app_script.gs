function copySheet(name){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = spreadsheet.getSheetByName("template"); // ss = source sheet 
  ss.copyTo(spreadsheet).setName(name); // copy sheet to target spreadsheet
}
function convertNumberToChinese(num) {
 const chineseNumbers = ['零', '壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖'];
 const unit = ['', '拾', '佰', '仟', '萬'];
 
 num = num.toString().split('').reverse().join('');
 let result = '';
 
 for (let i = 0; i < num.length; i++) {
   result += unit[i] + chineseNumbers[num[i]];
 }
 
 return result.split('').reverse().join('') + '元整';
}
 
function isNumeric(s) {
 return !isNaN(parseFloat(s)) && isFinite(s);
}
function askForMonth() {
 var ui = SpreadsheetApp.getUi();
 var response = ui.prompt('輸入月份', '輸入月份(EX:1):', ui.ButtonSet.OK_CANCEL);
 var button = response.getSelectedButton();
 var month = response.getResponseText();
 if (button == ui.Button.OK) {
 } else if (button == ui.Button.CANCEL ||
            button == ui.Button.CLOSE) {
   ui.alert('未選擇月份');
   month=-1;
 }
 if (isNumeric(month) &&
     parseInt(month, 10) >= 1 &&
     parseInt(month, 10) <= 12) {
 } else {
   var msg = '月份不合理:' + month;
   ui.alert(msg);
   month=-1;
 }
 return month;
}

function askForInput(total) {
 var ui = SpreadsheetApp.getUi();
 var startResponse = ui.prompt('輸入起始資料以及最後資料', '總共有' + total + '筆資料，由1開始，最後一筆資料為' +(total)+' 建議一次100筆\n EX:要區間1~10的話，請填寫:1,10', ui.ButtonSet.OK_CANCEL);
 var button = startResponse.getSelectedButton();
 var response = startResponse.getResponseText();
 if (button == ui.Button.OK) {
 } else if (button == ui.Button.CANCEL ||
            button == ui.Button.CLOSE) {
   ui.alert('未選擇資料');
   return [-1,-1];
 }

 var start = response.match(/(\d+),(\d+)/)[1];
 var end = response.match(/(\d+),(\d+)/)[2];
 if (isNumeric(start) &&
     parseInt(start, 10) >= 1 &&
     parseInt(start, 10) <= total &&
     parseInt(end, 10) > start &&
     parseInt(end, 10) <= total ) {
 } else {
   var msg = '輸入不合理:' + start + "~" +end;
   ui.alert(msg);
   start=-1;
 }
 
 return [ start, end];
}
function createReceiptLists() {
  var inputMonth = askForMonth();
  createReceipt(inputMonth, true, false)
}
function createMonthReport() {
  var inputMonth = askForMonth();
  createReceipt(inputMonth, false, true)
}

function createReceipt(inputMonth, printReceipt, printReport) {
 if(inputMonth == -1){
   return;
 }
 // Get the active spreadsheet
 var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 var targetSheet = spreadsheet.getSheetByName("response");
 
 // Get the range of data
 var dataRange = targetSheet.getDataRange();
 
 // Get the values of the data
 var data = dataRange.getValues();
 
 const caseNameColumn = 1;
 const dateColumn = 2;
 const typeColumn = 3;
 const nameColumns = [4,6,8,10,12,14,16,18,20,];
 const classColumns = [5,7,9,11,13,15,17,19];
 var currentDateee = new Date();
 var currentTime = currentDateee.toLocaleTimeString(); // "12:35 PM", for instance
 console.log(currentTime);
 // Create an array to store the unique values
 const map = new Map();
 const docMap = new Map();
 // Iterate through the data and store the unique values in the array
 for (var i = 1; i < data.length; i++) {
   var caseName = data[i][caseNameColumn];
   var cur_date = data[i][dateColumn];
   var cur_type = data[i][typeColumn];
   const dateee = new Date(cur_date);
   const month = dateee.getMonth() + 1;
   if(month == inputMonth){
     const doc_name = data[i].find((word, index) => word.length != 0 && (index>=4 && )nameColumns.includes(index));
     const classes = data[i].find((word, index) => word.length != 0 && classColumns.includes(index));
     const price = classes.match(/\d+/)[0];
     const month_date = month + "/" + dateee.getDate();
     /* Case */
     var info = {
         doctorName : doc_name + " " + cur_type,
         classes : classes,
         price : price
     };
     var class_therpist_price = doc_name + classes+ price;
     if(!map.has(caseName)){
       map.set(caseName, new Map());
     }
     if(!map.get(caseName).has(class_therpist_price)){
       map.get(caseName).set(class_therpist_price, [info]);
     }
     map.get(caseName).get(class_therpist_price).push(month_date);
     /* Doc */
     if(!docMap.has(doc_name)){
       docMap.set(doc_name, new Map());
     }
     if(!docMap.get(doc_name).has(classes)){
       docMap.get(doc_name).set(classes, []);
     }
     docMap.get(doc_name).get(classes).push(month_date);
   }
 }
 if(printReceipt){
  var caseNum = map.size;
  var getStartEnd = askForInput(caseNum);
  var caseStart = getStartEnd[0];
  var caseEnd = getStartEnd[1];
  if(caseStart == -1 || caseEnd==-1){
    return;
  }
  var currentDate = new Date();
  var year = currentDate.getFullYear();
  var month = currentDate.getMonth() + 1;  // January is 0
  var day = currentDate.getDate();
  var dateStr = '民國'+(year-1911)+'年'+month+'月'+day+'日';
  // Create a sheet for each unique value

  // Set template
  const templateSheet = spreadsheet.getSheetByName("template");
  templateSheet.getRange("A2").setValue(inputMonth + '月收據');
  templateSheet.getRange("A4").setValue("開立日期 ：" + dateStr);
  
  var caseCnt = 1;
  for (var caseName of map.keys()) {
    if(caseCnt >= caseStart && caseCnt <= caseEnd){
      if(caseCnt %10==0){
        console.log(caseCnt);
      }
        var newSheetName = caseName +"_" + year +"_"+month +"月_case" + (caseCnt-1);
        const checkSheet = spreadsheet.getSheetByName(newSheetName);
        if (checkSheet) {
          spreadsheet.deleteSheet(checkSheet);
        }
        copySheet(newSheetName);
        const newSheet = SpreadsheetApp.getActive().getSheetByName(newSheetName);
        //newSheet.activate();
        var total = 0;
        var reciptMap = map.get(caseName);
        var num = 0;
        newSheet.getRange("A3").setValue('姓名：' + caseName);
        for( const class_therpist_price of reciptMap.keys()){
          var valueOfMap = reciptMap.get(class_therpist_price);
          var info = valueOfMap[0];
          var classNum = valueOfMap.length-1;
          const total_price = info.price * classNum;
          var classes = info.classes;
          let matches = classes.match(/^([^(]+)/);
          if (matches) {
            classes = matches[1];
          }
          //newSheet.appendRow([classes, info.doctorName, valueOfMap.slice(1).reverse().join(","), classNum, info.price, total_price]);
          var curRow = 6+num;
          newSheet.getRange("A" + curRow +":F"+curRow).setValues([[classes, info.doctorName, valueOfMap.slice(1).reverse().join(","), classNum, info.price, total_price]]);
          total+=total_price;
          num++;
        }
        newSheet.getRange("B12:F12").setValue(total);
        newSheet.getRange("B13:F13").setValue(convertNumberToChinese(total));
    }
    caseCnt++;
  }
 }
 /* DocMap report */
 if(printReport){
  var reportSheetName = "營運報表" + inputMonth + "月";
  const checkSheet = spreadsheet.getSheetByName(reportSheetName);
  if (checkSheet) {
    spreadsheet.deleteSheet(checkSheet);
  }
  var reportSheet = spreadsheet.insertSheet(reportSheetName);
  reportSheet.appendRow(['治療師', '治療內容','堂數', '單價','小計']);
  for (var docName of docMap.keys()) {
    var classesMap = docMap.get(docName);
    reportSheet.appendRow([docName]);
    //console.log(docName);
    var total_num = 0;
    for(var className of classesMap.keys()){
      const price = className.match(/\d+/)[0];
      const dates = classesMap.get(className);
      const prices = price * dates.length;
      reportSheet.appendRow(["", className, dates.length, price, prices]);
      total_num+=prices;
      //console.log("    " + className + ":( " + dates.join(",")+" ) = " + prices);
    }
    reportSheet.appendRow(["當月收入", "", "", "", total_num]);
    reportSheet.appendRow([" "]);
  }
  currentTime = new Date().toLocaleTimeString(); // "12:35 PM", for instance
  console.log(currentTime);
 }
}
 
function onOpen() {
 var spreadsheet = SpreadsheetApp.getActive();
 var menuItems = [
   {name: '生成收據', functionName: 'createReceiptLists'},
   {name: '生成營業報表', functionName: 'createMonthReport'}
 ];
 spreadsheet.addMenu('自訂程式', menuItems);
}
 



