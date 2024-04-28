var TEMPLATE_ID = '1ny3XrbECqL7L7gZtuVrYNQd7FVUVnTrmBeIIrZIv7Rc'     // id google slide เกียรติบัตร******************
 
 
 
var TOTAL_SCORE = 10     // ระบุจำนวนข้อมสอบที่มี *********************************************************************
 
 
 
var PASS_PERCENT = 70      // กำหนดเกณฑ์ผ่าน***********************************************************************
 
 
 
var SAVE_FOLDER_ID = '1VX89FyCWPknbTSsEHlBhAidvVl6lLMgw';  //สร้าง folder เก็บ pdf ตั้งชื่อ เปิดแชร์ ให้คนมีลิงก์ ดูได้ เอา id มาวาง**********************************************************************************************************
 
 
 
// สร้าง sheet 1 แผ่น แค่ไปที่ sheet กด + มุมล่างซ้าย
 
 
 
// กด save scirpt นี้ กดบ่อย ๆ ก็ดี
 
 
 
// กำหนด trigger เป็น เมื่อส่ง form
 
 
 
//--------------------------------จบ-------------------------------------------
 
 
 
 
 
var email_column = 'ที่อยู่อีเมล'
 
 
 
var date = new Date(); 
 
 
 
var YEAR = date.getFullYear();
 
 
 
function createPdf(event) {
 
 
 
  if (TEMPLATE_ID === '') {  
 
      throw new Error('TEMPLATE_ID needs to be defined in Code.gs')
 
  }
 
  
 
  var activeSheet
 
  
 
  var activeRowIndex
 
  
 
  var range
 
  
 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 
  
 
  var sheets =  ss.getSheets()
 
  
 
  var certify_sheet = sheets[1]
 
  
 
 // var running_number_sheet = sheets[2]
 
  
 
  if (typeof event === 'undefined') {
 
    
 
    activeSheet = SpreadsheetApp.getActiveSheet()
 
   
 
 
 
    if (activeSheet === null) {
 
      throw new Error('Select a cell in the row that you want to use')
 
    }
 
    
 
    range = activeSheet.getActiveRange()
 
    
 
    if (activeSheet === null) {
 
      throw new Error('Select a cell in the row that you want to use')
 
    }
 
    
 
    activeRowIndex = range.getRowIndex()
 
 
 
  } else {
 
    
 
    range = event.range
 
    
 
    activeSheet = range.getSheet()
 
    
 
    activeRowIndex = range.getRowIndex()
 
    
 
  }
 
  
 
  //var numberOfColumns = activeSheet.getLastColumn()
 
  
 
  var numberOfColumns = 6
 
  
 
  var activeRow = activeSheet.getRange(activeRowIndex, 1, 1, numberOfColumns).getValues()
 
  
 
  var headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues()
 
  
 
  //Logger.log(headerRow)
 
  
 
  var columnIndex = 0
 
  
 
  var headerValue
 
  
 
  var activeCell
 
  
 
  var ID = null
 
  
 
  var recipient = null
 
  
 
  var user_email  = activeRow[0][1]
 
  
 
  var user_score = activeRow[0][2]
 
  
 
  var user_name  = activeRow[0][3]


 
  // var user_position  = activeRow[0][4]
 
  
 
  // var school_name  = activeRow[0][5]
 
 
 
  var percentage = (user_score/TOTAL_SCORE) * 100
 
   
 
  Logger.log(user_name, percentage, )
 
 
 
  var slide_file = DriveApp.getFileById(TEMPLATE_ID).makeCopy()
 
  
 
  var copyFile = slide_file.makeCopy('เกียรติบัตรผ่านการอบรมของ '+user_name);
 
  
 
  var copyId = copyFile.getId()
 
  
 
  var copyDoc = SlidesApp.openById(copyId);
 
  
 
  var slides = copyDoc.getSlides();
 
  
 
  var templateSlide = slides[0];
 
  
 
  var shapes = templateSlide.getShapes(); 
 
     
 
  var count,any_file,all_files,save_pdf_folder;//Define variables without assigning a value
 
 
 
  save_pdf_folder = DriveApp.getFolderById(SAVE_FOLDER_ID);
 
  
 
  //-------------------------------
 
  if (percentage >= PASS_PERCENT){
 
 
 
    //------------------------------------------------------------
 
    
 
    //--------------------------------------------
 
  
 
    shapes.forEach(function (shape) {
 
      
 
   //   shape.getText().replaceAllText("{{id}}", thaiNumber(certify_running_id));
 
      
 
      shape.getText().replaceAllText("{{name}}", user_name);
 
      
 
      // shape.getText().replaceAllText("{{position}}", user_position);
 
      
 
      // shape.getText().replaceAllText('{{school}}', school_name); 
 
         
 
      shape.getText().replaceAllText("{{email}}", user_email);
 
      
 
      shape.getText().replaceAllText("{{score}}", thaiNumber(percentage));
 
 
 
      shape.getText().replaceAllText("{{date}}", getThaiDate()); 
 
      
 
      shape.getText().replaceAllText('{{month}}',getThaiMonth());
 
      
 
      shape.getText().replaceAllText('{{year}}', getThaiYear());  
 
      
 
    });
 
    
 
    copyDoc.saveAndClose()
 
    
 
    var pdf_file = DriveApp.createFile(copyFile.getAs("application/pdf"));   
 
    
 
    var pdf_download_url = pdf_file.getDownloadUrl()
 
 
 
    certify_sheet.appendRow([ user_email, percentage, pdf_download_url ]);
 
    
 
    save_pdf_folder.addFile(pdf_file);
 
      
 
    DriveApp.removeFile(pdf_file) 
 
    
 
  } else { 
 
    
 
    Logger.log("ขอแสดงความเสียใจ คุณสอบไม่ผ่าน") 
 
    
 
  }
 
  
 
  slide_file.setTrashed(true);
 
  
 
  copyFile.setTrashed(true);
 
   
 
}
 
 
 
function getThaiDate() {
 
 
 
  var date = new Date();
 
 
 
  var DATE = date.getDate(); 
 
 
 
  return thaiNumber(DATE);
 
 
 
}
 
 
 
function getThaiMonth() {
 
 
 
  var date = new Date();
 
 
 
  var DATE = date.getDate();
 
 
 
  var MONTH = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
 
 
 
  var THAI_MONTH = MONTH[date.getMonth()]; 
 
 
 
  return THAI_MONTH
 
 
 
}
 
 
 
function getThaiYear() { 
 
 
 
  var date = new Date(); 
 
 
 
  var YEAR = date.getFullYear();
 
 
 
  var THAI_YEAR = YEAR + 543; 
 
 
 
  return thaiNumber(THAI_YEAR);
 
 
 
}
 
 
 
function thaiNumber(num){
 
 
 
 var array = {"1":"1", "2":"2", "3":"3", "4" : "4", "5" : "5", "6" : "6", "7" : "7", "8" : "8", "9" : "9", "0" : "0"};
 
 
 
 var str = num.toString();
 
 
 
 for (var val in array) {
 
 
 
  str = str.split(val).join(array[val]);
 
 
 
 }
 
 
 
 return str;
 
 
 
}
 
