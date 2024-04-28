
function doGet() {
  return HtmlService.createTemplateFromFile('index')
  .evaluate()
  .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}
function submitData(searchEmail){ //searchEmail คือข้อมูลที่ได้จาก form ในท่่นี้คือ email ที่ต้องการค้นหา
    var ss = SpreadsheetApp.openById("1CkGKYZ_EiQVikar1vqXQec2XGVWOEtsG7EFdVi3XLB0");  // id google sheet การตอบกลับ**********
    
    var sheet = ss.getSheetByName("ชีต1");    //ระบุชีตที่ต้องการ
    
    var lr = sheet.getLastRow();    //ค้นหาข้อมูลแถวสุดท้าย

    var  flag  =  1 ;    //กำหนดตัวแปร flag ให้มีค่าเริ่มต้น = 1 และนิยามว่า ไม่มีข้อมูล

    var data="<table class='table table-boardered table-striped table-hover'>";
    //สร้างข้อมูลตอบกลับ ในรูปแบบตาราง

          for(var i = 1;i <= lr;i++){             //วนลูป จาก 1 ไปจนถึงจำนวนข้อมูล

    var email = sheet.getRange(i, 1).getValue();
    // สร้างตัวแปรชื่อ email มีค่าเท่ากับข้อมูล คอลัม A1 , A2 ,... An ตามจำนวนของข้อมูล ในที่นี้คือ email

      if(email == searchEmail){ // ตรวจสอบว่า email (email ใน google sheet) ตรงกับ email ที่ค้นหาหรือไม่

      flag = 0;   // ถ้าตรง เปลี่ยนค่าตัวแปร flag เป็น 0 นิยามว่า มีข้อมูล

    var LinkCer = sheet.getRange(i, 3).getDisplayValue(); // กำหนดตัวแปรชื่อ LinkCer เพื่อเก็บ ข้อมูลจาก คอลัม C คือ link

  data +="<tr><td>อีเมล:</td><td>"+email+"</td></tr>"; //สร้างข้อมูลแถวที่ 1 แสดง Email

  data+="<tr><td>ลิงก์ดาวน์โหลด:</td><td><a href='"+LinkCer+"' target='_blank' class='btn btn-warning'>ดาวน์โหลดใบประกาศ</a></td></tr>"; 
  // สร้างแถวข้อมูลแถวที่ 2

    }//จบคำสั่ง วน ลูป จากบรรทัดที่ 21
   }
   data+='</table>';//สร้างข้อมูลตอบกลับ ปิด tag table

if(flag==1){
  // ตรวจสอบว่า flag มีค่าเท่ากับ 1 หรือไม่ // โดยอ้างอิง flag จากบรรทัดที่ 16 กรณีที่หาข้อมูลไม่พบ หากพบข้อมูล จะอ้างอิง flag ที่บรรทัดที่ 28
  
  var data ="<div class='alert alert-danger'>ไม่พบข้อมูล.</div>"; // ถ้าใช่จะส่งข้อความ ไม่พบข้อมูลกลับ 
    }// ออกจากการตรวจสอบเงื่อนไข
return data; //ส่งข้อมูลไปกลับ

    };
