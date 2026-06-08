
const OUTBOUND_SHEET_NAME = "Outbound_Books";
const CONTROL_SHEET_NAME = "Control_Ledger";
const EVIDENCE_FOLDER_ID = "root";
// นำบรรทัดนี้ไปเพิ่มไว้ด้านบนสุดของไฟล์
const DB_ID = '1nhhXnDwYw1Gaysp50Vew2_-LFsT23oGRkXKIrxQoS2w';

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function doGet() {
  ensureDatabaseExists();
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ระบบทะเบียนรับและติดตามหนังสือ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ตรวจสอบและสร้างฐานข้อมูลทั้งหมดของระบบ (เพิ่มโครงสร้างคลังน้ำมัน ยานพาหนะ และระบบใช้รถ)
function ensureDatabaseExists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ตาราง Book_Register
  var sheetRegister = ss.getSheetByName('Book_Register');
  var registerHeaders = ['book_id', 'วันที่', 'ทะเบียนรับ', 'กห', 'ที่', 'ลงวันที่', 'จาก', 'เรื่อง', 'status', 'current_dept', 'price', 'book_type'];
  if (!sheetRegister) {
    sheetRegister = ss.insertSheet('Book_Register');
    sheetRegister.appendRow(registerHeaders);
    sheetRegister.getRange(1, 1, 1, registerHeaders.length).setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 2. ตาราง Book_Timeline
  var sheetTimeline = ss.getSheetByName('Book_Timeline');
  var timelineHeaders = ['timeline_id', 'book_id', 'department_from', 'department_to', 'receiver_name', 'signature_data', 'timestamp', 'note'];
  if (!sheetTimeline) {
    sheetTimeline = ss.insertSheet('Book_Timeline');
    sheetTimeline.appendRow(timelineHeaders);
    sheetTimeline.getRange(1, 1, 1, timelineHeaders.length).setFontWeight('bold').setBackground('#065f46').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 3. ตาราง Users
  var sheetUsers = ss.getSheetByName('Users');
  var usersHeaders = ['username', 'password', 'department', 'name_title', 'profile_pic'];
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet('Users');
    sheetUsers.appendRow(usersHeaders);
    sheetUsers.appendRow(['admin', '1234', 'ผกบ.', 'แอดมิน ผกบ.', '']);
    sheetUsers.getRange(1, 1, 1, usersHeaders.length).setFontWeight('bold').setBackground('#9333ea').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 4. ตาราง Settings_Depts
  var sheetDepts = ss.getSheetByName('Settings_Depts');
  if (!sheetDepts) {
    sheetDepts = ss.insertSheet('Settings_Depts');
    sheetDepts.appendRow(['department_name']);
    var defaultDepts = [['ผู้บังคับบัญชา'], ['ผกบ.'], ['ฝพธ.'], ['ฝสภ.'], ['แหล่งรวมรถ'], ['โรงซักรีด'], ['โรงประกอบเลี้ยงคนไข้']];
    sheetDepts.getRange(2, 1, defaultDepts.length, 1).setValues(defaultDepts);
    sheetDepts.getRange(1, 1, 1, 1).setFontWeight('bold').setBackground('#374151').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 5. ตาราง Outbound_Books
  var sheetOutbound = ss.getSheetByName(OUTBOUND_SHEET_NAME);
  var outboundHeaders = ['เลขที่', 'วันที่', 'เรื่อง', 'ไปยัง', 'ประเภท', 'จำนวนรายการ', 'จำนวนราคา', 'ผู้รับผิดชอบ', 'ผู้ตรวจ', 'Timestamp', 'ไฟล์หลักฐาน', 'out_id', 'สป_สาย'];
  if (!sheetOutbound) {
    sheetOutbound = ss.insertSheet(OUTBOUND_SHEET_NAME);
    sheetOutbound.appendRow(outboundHeaders);
    sheetOutbound.getRange(1, 1, 1, outboundHeaders.length).setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 6. ตาราง Control_Ledger
  var sheetControl = ss.getSheetByName(CONTROL_SHEET_NAME);
  var controlHeaders = ['เลขที่', 'วันที่', 'เรื่อง', 'ไปยัง', 'ประเภท', 'จำนวนรายการ', 'จำนวนราคา', 'ผู้รับผิดชอบ', 'ผู้ตรวจ', 'ไฟล์หลักฐาน', 'วันที่ขึ้นบัญชี', 'out_id', 'สป_สาย'];
  if (!sheetControl) {
    sheetControl = ss.insertSheet(CONTROL_SHEET_NAME);
    sheetControl.appendRow(controlHeaders);
    sheetControl.getRange(1, 1, 1, controlHeaders.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 7. ตาราง Book_Procurement
  var sheetProc = ss.getSheetByName('Book_Procurement');
  var procHeaders = ['proc_id', 'book_id', 'step_name', 'amount', 'date', 'note', 'responsible', 'timestamp'];
  if (!sheetProc) {
    sheetProc = ss.insertSheet('Book_Procurement');
    sheetProc.appendRow(procHeaders);
    sheetProc.getRange(1, 1, 1, procHeaders.length).setFontWeight('bold').setBackground('#0d9488').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // ตรวจสอบโครงสร้างคลังน้ำมัน ยานพาหนะ และระบบใช้รถประจำวัน
  getOrCreateSheet('Fuel_Depot');
  getOrCreateSheet('Vehicles');
  getOrCreateSheet('Fuel_Log');
  getOrCreateSheet('Vehicle_Usage');
}


// ==========================================
// ฟังก์ชันสำหรับสั่งสร้างฐานข้อมูลครั้งแรก (รันแค่ครั้งเดียว)
// ==========================================
function setupDatabase() {
  try {
    getOrCreateSheet('Fuel_Depot');
    getOrCreateSheet('Vehicles');
    getOrCreateSheet('Fuel_Log');
    getOrCreateSheet('Vehicle_Usage');
    Logger.log("สร้างฐานข้อมูลสำเร็จทั้ง 4 ชีต!");
  } catch (error) {
    Logger.log("เกิดข้อผิดพลาด: " + error);
  }
}

// ฟังก์ชันหลักที่ใช้สร้างชีต
function getOrCreateSheet(sheetName) {
  // ดึงไฟล์ Database ตาม DB_ID ที่ประกาศไว้บนสุดของโปรเจกต์
  var ss = SpreadsheetApp.openById(DB_ID); 
  
  // (ถ้าคุณฝังสคริปต์ไว้ในไฟล์ Google Sheets ให้ใช้บรรทัดล่างนี้แทน)
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    if (sheetName === 'Fuel_Depot') {
      sheet.appendRow(["รหัสใบเบิก", "วันที่", "ปีงบประมาณ", "อัตรา/เครดิต", "ชนิดน้ำมัน", "จำนวน (ลิตร)", "ไฟล์หลักฐาน", "ผู้รับผิดชอบ", "สถานะ"]);
      sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
      
    } else if (sheetName === 'Vehicles') {
      sheet.appendRow(["รหัสรถ", "ประเภทรถ", "ยี่ห้อ", "ทะเบียนราชการ", "ทะเบียนพลเรือน", "น้ำมันอัตราพิกัด", "อัตราความสิ้นเปลือง", "เลขไมล์ปัจจุบัน", "ผู้รับผิดชอบ", "เลขไมล์เริ่มต้น"]);
      sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
      
    } else if (sheetName === 'Fuel_Log') {
      sheet.appendRow(["รหัสการเติม", "วันที่", "รหัสรถ", "ทะเบียนรถ", "ชนิดน้ำมัน", "จำนวน (ลิตร)", "เลขไมล์ตอนเติม", "ผู้บันทึก"]);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#059669").setFontColor("#ffffff");
      
    } else if (sheetName === 'Vehicle_Usage') {
      sheet.appendRow(["รหัสการใช้รถ", "รหัสรถ", "ทะเบียนรถ", "พลขับ", "ภารกิจ", "เลขไมล์ก่อนออก", "เลขไมล์กลับ", "เวลาออก", "เวลากลับ", "สถานะ"]);
      sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    }
    
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

// ----------------------------------------------------
// ระบบลงทะเบียน และยืนยันตัวตน
// ----------------------------------------------------
function getDepartments() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings_Depts');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var depts = [];
  for (var i = 1; i < data.length; i++) { if (data[i][0]) depts.push(data[i][0]); }
  return depts;
}

function registerUser(d) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = s.getDataRange().getValues();

  // 1. เช็ค Username ซ้ำ
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === d.username) return { success: false, message: "Username นี้มีผู้ใช้งานแล้วครับ" };
  }

  // 2. จัดการรูปภาพ (สร้างไฟล์ลงโฟลเดอร์)
  var profileUrl = "";
  if (d.base64Profile && d.base64Profile !== "") {
    try {
      var bytes = Utilities.base64Decode(d.base64Profile.split(',')[1]);
      var blob = Utilities.newBlob(bytes, d.mimeType || 'image/jpeg', d.fileName || 'profile.jpg');
      
      var folderId = "1poakHrgkZ6Nd5xPr3YF9dhgwwipz1mXq"; 
      var folder = DriveApp.getFolderById(folderId);
      
      // สร้างไฟล์ลงโฟลเดอร์ (รูปเข้า Drive ตอนนี้)
      var file = folder.createFile(blob);
      
      // 🎯 แก้ไข: ดึง ID และสร้างลิงก์ทันทีที่สร้างไฟล์เสร็จ (ก่อนที่มันจะ error)
      profileUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w500";
      
      // 🎯 แก้ไข: แยกการแชร์ไฟล์ไว้ใน try-catch อีกชั้น 
      // ถ้าบัญชีองค์กรบล็อกการแชร์ โค้ดก็จะไม่พัง และลิงก์ก็ยังถูกบันทึกได้ปกติ
      try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(shareErr) {
          // ปล่อยผ่านไปเลย เพราะโฟลเดอร์หลักเปิดแชร์ไว้อยู่แล้ว
      }
      
    } catch (e) {
      // 🎯 เพิ่มการแจ้งเตือน หากรูปอัปโหลดไม่ผ่านจริงๆ จะได้รู้ว่าติดที่อะไร
      return { success: false, message: "อัปโหลดรูปภาพมีปัญหา: " + e.toString() }; 
    }
  }

  // 3. บันทึกลงตาราง 
  s.appendRow([d.username, d.password, d.department, d.name, profileUrl]);
  
  return { success: true, message: "สมัครสมาชิกสำเร็จ!" };
}

function verifyLogin(u, p) {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == u && data[i][1] == p) return { success: true, username: data[i][0], password: data[i][1], department: data[i][2], name: data[i][3], profile_pic: data[i][4] };
  }
  return { success: false, message: "ชื่อผู้ใช้/รหัสผ่านผิด" };
}

function updateProfile(d) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === d.username) {
      s.getRange(i + 1, 2).setValue(d.password); s.getRange(i + 1, 4).setValue(d.name_title);
      if (d.profile_pic) s.getRange(i + 1, 5).setValue(d.profile_pic);
      return { success: true, message: "อัปเดตสำเร็จ", name: d.name_title, profile_pic: d.profile_pic || data[i][4] };
    }
  }
}

function getBooks(dept) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Book_Register');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  // -------------------------------------------------------------
  // 🟢 ส่วนที่แทรกเพิ่ม: ไปดึง "ขั้นตอนล่าสุด" เตรียมรอเอาไว้ก่อน
  // -------------------------------------------------------------
  var stepSheet = ss.getSheetByName('Book_Procurement_Steps');
  var latestStepsObj = {}; 
  
  if (stepSheet) {
    var stepsData = stepSheet.getDataRange().getValues();
    for (var s = 1; s < stepsData.length; s++) {
      // คอลัมน์ B (index 1) = book_id, คอลัมน์ C (index 2) = ชื่อขั้นตอน
      var stepBookId = String(stepsData[s][1]).trim(); 
      var stepName = String(stepsData[s][2]).trim(); 
      
      if (stepBookId !== "" && stepName !== "") {
        latestStepsObj[stepBookId] = stepName; // วนทับไปเรื่อยๆ จะได้บรรทัดล่างสุดเสมอ
      }
    }
  }
  // -------------------------------------------------------------

  var books = [];
  for (var i = 1; i < data.length; i++) {
    if (dept === 'ผกบ.' || data[i][9] === dept) {
      var dateVal = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], "GMT+7", "yyyy-MM-dd") : (data[i][1] || "");
      var bookDateVal = (data[i][5] instanceof Date) ? Utilities.formatDate(data[i][5], "GMT+7", "yyyy-MM-dd") : (data[i][5] || "");
      
      var currentBookId = String(data[i][0]).trim(); // ดึงรหัสหนังสือเพื่อเอาไปเทียบ

      books.push({
        book_id: data[i][0] || "", 
        วันที่: dateVal, 
        ทะเบียนรับ: data[i][2] || "", 
        กห: data[i][3] || "", 
        ที่: data[i][4] || "", 
        ลงวันที่: bookDateVal,
        จาก: data[i][6] || "", 
        เรื่อง: data[i][7] || "", 
        status: data[i][8] || "รอดำเนินการ", 
        current_dept: data[i][9] || "", 
        price: data[i][10] || "", 
        book_type: data[i][11] || "ทั่วไป",
        
        // 🎯 สิ่งที่เพิ่มมามีแค่บรรทัดนี้บรรทัดเดียวครับ นอกนั้นของคุณเหมือนเดิมเป๊ะ!
        latest_step: latestStepsObj[currentBookId] || 'ยังไม่มีการดำเนินการ'
      });
    }
  }
  return books.reverse();
}

function forwardBook(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet = ss.getSheetByName('Book_Register');
  var data = regSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == d.book_id) {
      var targetStatus = d.is_last ? 'เสร็จสิ้น' : (d.is_return ? 'ส่งคืนแก้ไข' : 'อยู่ระหว่างดำเนินการ');
      regSheet.getRange(i + 1, 9).setValue(targetStatus); regSheet.getRange(i + 1, 10).setValue(d.department_to);
      break;
    }
  }
  var prefixNote = d.is_return ? "[ส่งคืนแก้ไข] " : "";
  var timelineNote = prefixNote + "เลขทะเบียนรับหน่วย: " + (d.regNum || "-") + " | วันที่รับ: " + (d.regDate || "-");
  if (d.note) { timelineNote += " | หมายเหตุ: " + d.note; }
  ss.getSheetByName('Book_Timeline').appendRow(['TL-' + new Date().getTime(), d.book_id, d.department_from, d.department_to, d.receiver_name, d.signature_data, new Date(), timelineNote]);
  return true;
}

function toggleReturnStatus(bookId, isReturn, userDept, userName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet = ss.getSheetByName('Book_Register');
  var data = regSheet.getDataRange().getValues();
  var newStatus = isReturn ? 'ส่งคืนแก้ไข' : 'อยู่ระหว่างดำเนินการ';
  for (var i = 1; i < data.length; i++) { if (data[i][0] == bookId) { regSheet.getRange(i + 1, 9).setValue(newStatus); break; } }
  ss.getSheetByName('Book_Timeline').appendRow(['TL-' + new Date().getTime(), bookId, userDept, userDept, userName, '', new Date(), isReturn ? '[สลับสวิตช์ตาราง] เปลี่ยนสถานะเป็นส่งคืนกลับไปแก้ไข' : '[สลับสวิตช์ตาราง] ยกเลิกสถานะส่งคืนแก้ไข (กลับเข้าสู่กระบวนการ)']);
  return true;
}

function registerBook(d) { 
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Book_Register').appendRow([
    'BK-'+new Date().getTime(), // A: ID
    d.date,                     // B: วันที่
    d.regNum,                   // C: ทะเบียน
    d.kh,                       // D: กห
    d.bookNum,                  // E: ที่
    d.bookDate,                 // F: ลงวันที่
    d.fromUnit,                 // G: จาก
    d.subject,                  // H: เรื่อง
    'รอดำเนินการ',                // I: สถานะ
    d.current_dept,             // J: แผนก
    d.price,                    // K: ราคา
    d.bookType,                 // L: ประเภท
    d.refId || ""               // M: 🆕 รหัสอ้างอิงเรื่องเดิม
  ]); 
}
function updateBook(d) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Book_Register'); 
  var data = s.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == d.book_id) {
      s.getRange(i + 1, 2).setValue(d.date); 
      s.getRange(i + 1, 3).setValue(d.regNum); 
      s.getRange(i + 1, 4).setValue(d.kh);
      s.getRange(i + 1, 5).setValue(d.bookNum); 
      s.getRange(i + 1, 6).setValue(d.bookDate); 
      s.getRange(i + 1, 7).setValue(d.fromUnit);
      s.getRange(i + 1, 8).setValue(d.subject); 
      s.getRange(i + 1, 11).setValue(d.price); 
      s.getRange(i + 1, 12).setValue(d.bookType);
      
      // 🆕 เพิ่มบรรทัดนี้: อัปเดตคอลัมน์ที่ 13 (ref_id)
      s.getRange(i + 1, 13).setValue(d.refId || ""); 
      
      break;
    }
  }
  return true;
}

function deleteBook(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ลบจาก Book_Register
  var s = ss.getSheetByName('Book_Register');
  var data = s.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) { // ไล่จากล่างขึ้นบน เพื่อป้องกันการขยับแถวแล้วพลาด
    if (data[i][0] == id) {
      s.deleteRow(i + 1);
      break;
    }
  }
  
  // 2. ลบจาก Book_Timeline
  var t = ss.getSheetByName('Book_Timeline');
  var td = t.getDataRange().getValues();
  for (var j = td.length - 1; j >= 1; j--) {
    if (td[j][1] == id) {
      t.deleteRow(j + 1);
    }
  }
  
  return true;
}

function getTimeline(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Book_Timeline'); if (!sheet) return [];
  var data = sheet.getDataRange().getValues(); var timelines = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] == id) {
      var ts = data[i][6] ? Utilities.formatDate(new Date(data[i][6]), "GMT+7", "dd/MM/yyyy HH:mm") : "";
      timelines.push({ timeline_id: data[i][0], book_id: data[i][1], department_from: data[i][2], department_to: data[i][3], receiver_name: data[i][4], signature_data: data[i][5], timestamp: ts, note: data[i][7] });
    }
  }
  return timelines;
}

function getUserWorkload() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Book_Timeline');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var counts = {};
  for (var i = 1; i < data.length; i++) {
    var name = data[i][4];
    var dept = data[i][2];
    if (name && dept) {
      var key = name + " [" + dept + "]";
      counts[key] = (counts[key] || 0) + 1;
    }
  }
  var list = [];
  for (var key in counts) {
    list.push({ label: key, value: counts[key] });
  }
  list.sort(function(a, b) { return b.value - a.value; });
  return list;
}

function getGlobalTimeline() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timelineSheet = ss.getSheetByName("Book_Timeline") || ss.getSheets()[1];
    var registerSheet = ss.getSheetByName("Book_Register"); 
    
    if (!timelineSheet) return [];
    var timelineData = timelineSheet.getDataRange().getValues();
    if (timelineData.length <= 1) return [];

    // ----------------------------------------------------------------
    // 1. สร้างคู่มือค้นหา (ดิกชันนารี) จากชีต Book_Register
    // ----------------------------------------------------------------
    var subjectMap = {};
    if (registerSheet) {
      var regData = registerSheet.getDataRange().getValues();
      
      // อิงจากภาพ Book_Register: A=book_id(0), H=เรื่อง(7)
      for (var i = 1; i < regData.length; i++) {
        var bookId = regData[i][0];       // คอลัมน์ A (book_id)
        var subjectName = regData[i][7];  // คอลัมน์ H (เรื่อง)
        
        if (bookId) {
          subjectMap[bookId.toString().trim()] = subjectName;
        }
      }
    }

    // ----------------------------------------------------------------
    // 2. ดึงข้อมูล Timeline และจับคู่ชื่อเรื่อง
    // ----------------------------------------------------------------
    var result = [];
    var startRow = Math.max(1, timelineData.length - 15); // ดึง 15 รายการล่าสุด
    
    for (var j = timelineData.length - 1; j >= startRow; j--) {
      var row = timelineData[j];
      
      // อิงจากภาพ Book_Timeline: book_id อยู่คอลัมน์ B (Index 1)
      var timelineBookId = row[1] ? row[1].toString().trim() : "";
      
      // นำ book_id ไปค้นหาชื่อเรื่อง
      var realSubject = subjectMap[timelineBookId];
      var displaySubject = realSubject ? realSubject : "อัปเดตสถานะเดินหนังสือ";

      result.push({
        department_from: row[2] || "ระบบ", // คอลัมน์ C
        department_to: row[3] || "-",      // คอลัมน์ D
        receiver_name: row[4] || "-",      // คอลัมน์ E
        timestamp: row[6] ? Utilities.formatDate(new Date(row[6]), "GMT+7", "dd/MM/yyyy HH:mm") : "", // คอลัมน์ G
        note: row[7] || "",                // คอลัมน์ H
        subject: displaySubject            // 🎯 แสดงชื่อเรื่องที่ดึงมาได้
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

// ฟังก์ชันหาเลขทะเบียนรับล่าสุด และรีเซ็ตเป็นเลข 1 เมื่อขึ้นปีใหม่
// ฟังก์ชันหาเลขทะเบียนรับล่าสุด และรีเซ็ตเป็นเลข 1 เมื่อขึ้นปีใหม่
function getNextRegisterNumber() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Book_Register"); 
    
    if (!sheet) return 1;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 1;

    const headers = data[0];
    
    // 🔴 จุดที่แก้ไข: ปรับให้ค้นหาคำว่า "วันที่" ตามหน้าชีตของท่าน
    let dateColIndex = headers.findIndex(h => String(h).trim() === "วันที่" || String(h).trim() === "วันที่รับ");
    let regColIndex = headers.findIndex(h => String(h).trim() === "ทะเบียนรับ");
    
    // ตั้งค่า Default เผื่อหาหัวตารางไม่เจอ (ตามภาพของท่าน วันที่อยู่คอลัมน์ B=1, ทะเบียนรับอยู่คอลัมน์ C=2)
    if (regColIndex === -1) regColIndex = 2; 
    if (dateColIndex === -1) dateColIndex = 1; 

    // หา "ปีปัจจุบัน"
    const currentYear = new Date().getFullYear();
    let maxNum = 0;

    // วนลูปเช็คข้อมูลทีละบรรทัด
    for (let i = 1; i < data.length; i++) {
      let rowDate = data[i][dateColIndex];
      let rowYear = 0;
      
      // สกัด "ปี" ออกมาจากช่องวันที่
      if (rowDate instanceof Date) {
        rowYear = rowDate.getFullYear();
      } else if (typeof rowDate === 'string' && rowDate.trim() !== "") {
        let dateMatch = rowDate.match(/\d{4}/);
        if (dateMatch) {
          rowYear = parseInt(dateMatch[0], 10);
          if (rowYear > 2500) rowYear -= 543;
        }
      }

      // ถ้าปีตรงกัน ให้เทียบหาเลขสูงสุด
      if (rowYear === currentYear) {
        let rawValue = String(data[i][regColIndex]).replace(/\D/g, ''); 
        let currentNum = parseInt(rawValue, 10);
        
        if (!isNaN(currentNum) && currentNum > maxNum) {
          maxNum = currentNum;
        }
      }
    }

    return maxNum + 1; // คืนค่า เลขล่าสุด + 1

  } catch (error) {
    return 1;
  }
}

// ----------------------------------------------------
// ระบบสารบรรณขึ้นบัญชีคุม สป. ทั่วไป
// ----------------------------------------------------
function getOutboundBooks() {
  // รบกวนเช็คชื่อชีตให้ตรงกับที่คุณใช้จริงนะครับ (เช่น 'Outbound_Books' หรือตัวแปร OUTBOUND_SHEET_NAME)
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Outbound_Books');
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    // ข้ามบรรทัดที่ว่างเปล่า
    if (data[i][0] === "" && data[i][2] === "") continue;

    // 1. แปลงคอลัมน์วันที่ (Index 1) เป็นข้อความ
    var dStr = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], "Asia/Bangkok", "yyyy-MM-dd") : String(data[i][1] || "");
    
    // 🔴 2. จุดแก้ปัญหา! แปลงคอลัมน์ Timestamp (Index 9) เป็นข้อความ
    var tsStr = (data[i][9] instanceof Date) ? Utilities.formatDate(data[i][9], "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss") : String(data[i][9] || "");

    result.push({
      no: data[i][0], 
      date: dStr, 
      subject: data[i][2], 
      to: data[i][3], 
      type: data[i][4],
      itemCount: data[i][5], 
      price: data[i][6], 
      responsible: data[i][7], 
      approver: data[i][8],
      timestamp: tsStr, // <--- ส่งเป็นข้อความแทน Date Object
      file_url: data[i][10], 
      out_id: data[i][11], 
      sapo: data[i][12]
    });
  }
  
  return result;
}

function saveOutboundBook(d) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OUTBOUND_SHEET_NAME);
  if (d.out_id) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][11] == d.out_id) {
        sheet.getRange(i+1, 2, 1, 8).setValues([[d.date, d.subject, d.to, d.type, d.itemCount, d.price, d.responsible, d.approver]]);
        sheet.getRange(i+1, 13).setValue(d.sapo);
        return { success: true };
      }
    }
  } else {
    var nextNo = sheet.getLastRow();
    var newId = 'OUT-' + new Date().getTime();
    sheet.appendRow([nextNo, d.date, d.subject, d.to, d.type, d.itemCount, d.price, d.responsible, d.approver, new Date(), '', newId, d.sapo]);
    return { success: true };
  }
}

function deleteOutboundBook(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OUTBOUND_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][11] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

function getControlLedger() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTROL_SHEET_NAME);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var dStr = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], "GMT+7", "yyyy-MM-dd") : data[i][1];
    var tsStr = (data[i][10] instanceof Date) ? Utilities.formatDate(data[i][10], "GMT+7", "dd/MM/yyyy HH:mm") : data[i][10];
    result.push({
      no: data[i][0], date: dStr, subject: data[i][2], to: data[i][3], type: data[i][4],
      itemCount: data[i][5], price: data[i][6], responsible: data[i][7], approver: data[i][8],
      file_url: data[i][9], timestamp: tsStr, out_id: data[i][11], sapo: data[i][12]
    });
  }
  return result.reverse();
}

function updateControlLedgerEntry(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control_Ledger");
    if (!sheet) throw new Error("ไม่พบชีต Control_Ledger");

    // 1. จัดการอัปโหลดไฟล์หลักฐานลง Google Drive
    let fileUrl = data.old_file_url || ""; 

    if (data.base64File && data.base64File !== "") {
      const base64Data = data.base64File.split(',')[1];
      const mimeType = data.mimeType || 'application/pdf';
      const fileName = data.fileName || ('Evidence_' + new Date().getTime());
      
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
      const newFile = DriveApp.createFile(blob);
      fileUrl = newFile.getUrl(); 
    }

    const sheetData = sheet.getDataRange().getDisplayValues();
    let rowIndex = -1;

    // 2. ค้นหาแถวที่จะแก้ไข (เปลี่ยนไปค้นหาจาก คอลัมน์ L / Index 11 แทน)
    if (data.out_id && data.out_id !== "") {
      for (let i = 1; i < sheetData.length; i++) {
        if (sheetData[i][11] === data.out_id) { // คอลัมน์ L คือ out_id
          rowIndex = i + 1;
          break;
        }
      }
    }

    const timestamp = new Date();

    if (rowIndex > -1) {
      // ---- กรณีแก้ไขข้อมูลเดิม (Update) ----
      sheet.getRange(rowIndex, 1).setValue(data.no);           // A: เลขที่
      sheet.getRange(rowIndex, 2).setValue(data.date);         // B: วันที่
      sheet.getRange(rowIndex, 3).setValue(data.subject);      // C: เรื่อง
      sheet.getRange(rowIndex, 4).setValue(data.to || "");     // D: ไปยัง
      sheet.getRange(rowIndex, 5).setValue(data.type);         // E: ประเภท
      sheet.getRange(rowIndex, 6).setValue(data.itemCount);    // F: จำนวนรายการ
      sheet.getRange(rowIndex, 7).setValue(data.price);        // G: จำนวนราคา
      sheet.getRange(rowIndex, 8).setValue(data.responsible);  // H: ผู้รับผิดชอบ
      sheet.getRange(rowIndex, 9).setValue(data.approver);     // I: ผู้ตรวจ
      sheet.getRange(rowIndex, 10).setValue(fileUrl);          // J: ไฟล์หลักฐาน
      sheet.getRange(rowIndex, 11).setValue(timestamp);        // K: วันที่ขึ้นบัญชี
      // L: out_id (ไม่ต้องแก้รหัสเดิม)
      sheet.getRange(rowIndex, 13).setValue(data.sapo);        // M: สป_สาย
      
      return { success: true, message: "อัปเดตบัญชีคุมและไฟล์สำเร็จ" };
      
    } else {
      // ---- กรณีเพิ่มข้อมูลใหม่ (Create) ----
      const newId = "CTRL-" + new Date().getTime();
      const docNo = data.no || sheet.getLastRow(); // ถ้าไม่มีลำดับ ให้รันต่อจากแถวล่าสุด
      
      // เรียง Array ให้ตรงกับคอลัมน์ A ถึง M เป๊ะๆ
      const newRow = [
        docNo,             // A: เลขที่
        data.date,         // B: วันที่
        data.subject,      // C: เรื่อง
        data.to || "",     // D: ไปยัง
        data.type || "ทั่วไป", // E: ประเภท
        data.itemCount,    // F: จำนวนรายการ
        data.price,        // G: จำนวนราคา
        data.responsible,  // H: ผู้รับผิดชอบ
        data.approver,     // I: ผู้ตรวจ
        fileUrl,           // J: ไฟล์หลักฐาน
        timestamp,         // K: วันที่ขึ้นบัญชี
        newId,             // L: out_id (รหัสซ่อน)
        data.sapo          // M: สป_สาย
      ];
      
      sheet.appendRow(newRow);
      return { success: true, message: "บันทึกบัญชีคุมพร้อมไฟล์สำเร็จ" };
    }
  } catch (error) {
    throw new Error(error.toString());
  }
}

function deleteControlLedgerEntry(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTROL_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][11] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

function registerToControlLedger(outId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetOut = ss.getSheetByName(OUTBOUND_SHEET_NAME);
  var sheetCtrl = ss.getSheetByName(CONTROL_SHEET_NAME);
  var dataOut = sheetOut.getDataRange().getValues();
  for (var i = 1; i < dataOut.length; i++) {
    if (dataOut[i][11] == outId) {
      var nextNo = sheetCtrl.getLastRow();
      sheetCtrl.appendRow([nextNo, dataOut[i][1], dataOut[i][2], dataOut[i][3], dataOut[i][4], dataOut[i][5], dataOut[i][6], dataOut[i][7], dataOut[i][8], dataOut[i][10], new Date(), outId, dataOut[i][12]]);
      return { success: true };
    }
  }
  return { success: false, message: "ไม่พบข้อมูลต้นทาง" };
}

// ----------------------------------------------------
// ระบบคลังน้ำมันและยานพาหนะ (Fuel System)
// ----------------------------------------------------
function getFuelDepots() {
  try {
    var sheet = getOrCreateSheet('Fuel_Depot');
    var data = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var dStr = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], "GMT+7", "yyyy-MM-dd") : data[i][1];
      result.push({
        id: String(data[i][0]), date: dStr, year: String(data[i][2]), credit: String(data[i][3]),
        type: String(data[i][4]), liter: String(data[i][5]), file: String(data[i][6]),
        responsible: String(data[i][7]), status: String(data[i][8])
      });
    }
    return result.reverse();
  } catch (e) { return []; }
}

function saveFuelDepot(d) {
  var sheet = getOrCreateSheet('Fuel_Depot');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == d.id) {
      sheet.getRange(i+1, 2, 1, 7).setValues([[d.date, d.year, d.credit, d.type, d.liter, d.file || '', d.responsible]]);
      sheet.getRange(i+1, 9).setValue(d.status);
      return { success: true };
    }
  }
  sheet.appendRow([d.id, d.date, d.year, d.credit, d.type, d.liter, d.file || '', d.responsible, d.status]);
  return { success: true };
}

function deleteFuelDepot(id) {
  var sheet = getOrCreateSheet('Fuel_Depot');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

function updateFuelDepotStatus(id, newStatus) {
  var sheet = getOrCreateSheet('Fuel_Depot');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) { sheet.getRange(i + 1, 9).setValue(newStatus); return { success: true }; }
  }
  return { success: false };
}

// ระบบจัดการยานพาหนะพร้อมอัตราสิ้นเปลือง และคำนวนน้ำมันคงถังสะสมแบบเรียลไทม์
function getVehicles() {
  try {
    var sheetV = getOrCreateSheet('Vehicles');
    var sheetL = getOrCreateSheet('Fuel_Log');
    var vData = sheetV.getDataRange().getValues();
    var lData = sheetL.getDataRange().getValues();
    
    var result = [];
    for (var i = 1; i < vData.length; i++) {
      if (!vData[i][0]) continue;
      var vId = String(vData[i][0]);
      
      // 1. คำนวณยอดเติมน้ำมันสะสม
      var totalFill = 0;
      for (var j = 1; j < lData.length; j++) {
        if (String(lData[j][2]) === vId) {
          totalFill += Number(lData[j][5]) || 0;
        }
      }

      // 2. จัดการเลขไมล์เริ่มต้น (ดึงจากคอลัมน์ J)
      var currMile = Number(vData[i][7]) || 0; // คอลัมน์ H (ไมล์ปัจจุบัน)
      var startMileRaw = vData[i][9];          // คอลัมน์ J (ไมล์เริ่มต้น)
      var startMile = Number(startMileRaw);
      
      // ถ้าไม่มีเลขไมล์เริ่มต้น โค้ดจะดึงไมล์ปัจจุบันไปใช้ และเขียนลงฐานข้อมูลให้อัตโนมัติ
      if (startMileRaw === "" || startMileRaw === undefined) {
        startMile = currMile;
        sheetV.getRange(i + 1, 10).setValue(startMile);
      }

      var distance = Math.max(0, currMile - startMile);
      
      // 3. ประเมินน้ำมันคงเหลือในถัง
      var consumeRate = Number(vData[i][6]) || 1;
      var estimatedConsume = distance / consumeRate;
      var fuelInTank = Math.max(0, Math.round(totalFill - estimatedConsume));

      // ส่งข้อมูลออกไปหน้าเว็บ
      result.push({
        id: vId, type: String(vData[i][1]), brand: String(vData[i][2]),
        govReg: String(vData[i][3]), civilReg: String(vData[i][4]),
        quota: String(vData[i][5]), consume: String(vData[i][6]),
        mileage: String(vData[i][7]), driver: String(vData[i][8]),
        startMileage: String(startMile), distance: distance, fuelInTank: fuelInTank
      });
    }
    return result;
  } catch (e) { return []; }
}

function saveVehicle(d) {
  try {
    var sheet = getOrCreateSheet('Vehicles');
    var data = sheet.getDataRange().getValues();
    
    // ตั้งค่าเลขไมล์เริ่มต้น: ถ้าหน้าเว็บส่ง d.startMileage มาให้ใช้ค่านั้น ถ้าไม่ส่งมาให้ใช้ d.mileage แทน
    var startMile = d.startMileage ? d.startMileage : d.mileage;
    
    // 1. กรณี: แก้ไขข้อมูลรถเดิม
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == d.id) {
        // อัปเดตรวดเดียวตั้งแต่คอลัมน์ที่ 2 (ประเภทรถ) จนถึงคอลัมน์ที่ 10 (เลขไมล์เริ่มต้น)
        sheet.getRange(i+1, 2, 1, 9).setValues([[
          d.type, d.brand, d.govReg, d.civilReg, d.quota, d.consume, d.mileage, d.driver, startMile
        ]]);
        return { success: true };
      }
    }
    
    // 2. กรณี: เพิ่มรถใหม่
    var id = "V" + new Date().getTime();
    sheet.appendRow([
      id, d.type, d.brand, d.govReg, d.civilReg, d.quota, d.consume, d.mileage, d.driver, startMile
    ]);
    
    return { success: true };
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteVehicle(id) {
  var sheet = getOrCreateSheet('Vehicles');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

// ==========================================
// ส่วนจัดการข้อมูล: ประวัติการเติมน้ำมัน (Fuel_Log) ปรับปรุงใหม่
// ==========================================

/**
 * ดึงข้อมูลประวัติการเติมน้ำมันทั้งหมด (เชื่อมต่อผ่าน DB_ID)
 */


/**
 * บันทึก/แก้ไขประวัติการเติมน้ำมัน
 */
function saveFuelLog(d) {
  try {
    var sheet = getOrCreateSheet('Fuel_Log');
    var isEdit = false;
    
    // ตรวจสอบว่าเป็นการแก้ไขข้อมูลเดิมหรือไม่
    if (d.logId) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] == d.logId) {
          sheet.getRange(i+1, 2, 1, 7).setValues([[d.date, d.vehicleId, d.vehicleReg, d.type, d.liter, d.mileage, d.recorder]]);
          isEdit = true; 
          break;
        }
      }
    }
    
    // หากไม่ใช่การแก้ไข ให้เพิ่มแถวใหม่
    if (!isEdit) {
      var logId = "L" + new Date().getTime(); // สร้าง ID อัตโนมัติ
      sheet.appendRow([logId, d.date, d.vehicleId, d.vehicleReg, d.type, d.liter, d.mileage, d.recorder]);
    }
    
    // อัปเดตเลขไมล์กลับไปที่ฐานข้อมูล Vehicles
    var sheetV = getOrCreateSheet('Vehicles');
    var vData = sheetV.getDataRange().getValues();
    for (var k = 1; k < vData.length; k++) {
      if (vData[k][0] == d.vehicleId) {
        var currentMileage = Number(vData[k][7] || 0);
        // เช็กไมล์ตอนเติม ต้องมากกว่าเลขไมล์ปัจจุบันในระบบถึงจะอัปเดต
        if (Number(d.mileage) > currentMileage) {
          sheetV.getRange(k + 1, 8).setValue(Number(d.mileage));
        }
        break;
      }
    }
    
    // 🎯 1. [เพิ่มตรงนี้] บังคับให้ Google Sheet เซฟข้อมูลทั้งหมดลงเซลล์ทันที
    SpreadsheetApp.flush(); 

    // 🎯 2. [เพิ่มตรงนี้] เรียกฟังก์ชันคำนวณยอดน้ำมัน (เพื่อเอาตัวเลขที่เพิ่งตัดยอดใหม่ๆ)
    // ⚠️ หมายเหตุ: เปลี่ยน getFuelDashboardData เป็นชื่อฟังก์ชันคำนวณน้ำมันที่คุณใช้อยู่จริง
    var freshStats = getFuelDashboardData(); 
    
    // 🎯 3. [แก้ตรงนี้] แนบ freshStats ส่งกลับไปให้หน้าเว็บด้วย
    return { 
      success: true, 
      message: 'บันทึกประวัติการเติมน้ำมันเรียบร้อย',
      updatedStats: freshStats 
    };

  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

/**
 * ลบประวัติการเติมน้ำมัน
 */
function deleteFuelLog(id) {
  try {
    var sheet = getOrCreateSheet('Fuel_Log');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) { 
        sheet.deleteRow(i + 1); 
        return { success: true, message: 'ลบประวัติการเติมน้ำมันสำเร็จ' }; 
      }
    }
    return { success: false, message: 'ไม่พบรายการที่ต้องการลบ' };
  } catch (error) {
    return { success: false, message: 'ลบข้อมูลไม่สำเร็จ: ' + error.toString() };
  }
}





function getFuelDashboardOverview() {
  try {
    var stats = { totalVehicles: 0, totalFuelAdded: 0, totalFuelDisbursed: 0, pendingRequests: 0, totalFuelInDepot: 0 };
    var sheetVehicles = getOrCreateSheet('Vehicles');
    stats.totalVehicles = Math.max(0, sheetVehicles.getLastRow() - 1);

    var sheetDepot = getOrCreateSheet('Fuel_Depot');
    var depotData = sheetDepot.getDataRange().getValues();
    for (var i = 1; i < depotData.length; i++) {
      var status = String(depotData[i][8]);
      if (status === 'รออนุมัติจ่าย' || status === 'รอดำเนินการ') {
        stats.pendingRequests++;
      }
      if (status === 'รับน้ำมันแล้ว' || status === 'เสร็จสิ้น') {
        stats.totalFuelAdded += Number(depotData[i][5]) || 0;
      }
    }

    var sheetLog = getOrCreateSheet('Fuel_Log');
    var logData = sheetLog.getDataRange().getValues();
    for (var j = 1; j < logData.length; j++) {
      stats.totalFuelDisbursed += Number(logData[j][5]) || 0;
    }

    stats.totalFuelInDepot = Math.max(0, stats.totalFuelAdded - stats.totalFuelDisbursed);
    return stats;
  } catch (e) {
    return { totalVehicles: 0, totalFuelAdded: 0, totalFuelDisbursed: 0, pendingRequests: 0, totalFuelInDepot: 0 };
  }
}

// ----------------------------------------------------
// ระบบขอใช้รถยนต์ประจำวัน (Vehicle Daily Usage)
// ----------------------------------------------------
function getVehicleUsages() {
  try {
    var sheet = getOrCreateSheet('Vehicle_Usage');
    var data = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var depTime = data[i][7] ? Utilities.formatDate(new Date(data[i][7]), "GMT+7", "yyyy-MM-dd'T'HH:mm") : "";
      var retTime = data[i][8] ? Utilities.formatDate(new Date(data[i][8]), "GMT+7", "yyyy-MM-dd'T'HH:mm") : "";
      result.push({
        usageId: String(data[i][0]), vehicleId: String(data[i][1]), vehicleReg: String(data[i][2]),
        driver: String(data[i][3]), purpose: String(data[i][4]), startMileage: String(data[i][5]),
        endMileage: String(data[i][6]), departureTime: depTime, returnTime: retTime, status: String(data[i][9])
      });
    }
    return result.reverse();
  } catch (e) { return []; }
}

function saveVehicleUsage(d) {
  var sheet = getOrCreateSheet('Vehicle_Usage');
  var usageId = "USG" + new Date().getTime();
  sheet.appendRow([usageId, d.vehicleId, d.vehicleReg, d.driver, d.purpose, d.startMileage, '', d.departureTime, '', 'อยู่ระหว่างภารกิจ']);
  
  // อัปเดตไมล์เริ่มต้นให้รถทันที
  var sheetV = getOrCreateSheet('Vehicles');
  var vData = sheetV.getDataRange().getValues();
  for (var i = 1; i < vData.length; i++) {
    if (vData[i][0] == d.vehicleId) {
      sheetV.getRange(i + 1, 8).setValue(Number(d.startMileage)); break;
    }
  }
  return { success: true };
}

function returnVehicleUsage(d) {
  var sheet = getOrCreateSheet('Vehicle_Usage');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == d.usageId) {
      sheet.getRange(i + 1, 7).setValue(Number(d.endMileage));
      sheet.getRange(i + 1, 9).setValue(d.returnTime);
      sheet.getRange(i + 1, 10).setValue('เสร็จสิ้นภารกิจ');
      
      // ส่งไมล์สะสมล่าสุดกลับไปที่ Vehicles
      var sheetV = getOrCreateSheet('Vehicles');
      var vData = sheetV.getDataRange().getValues();
      for (var k = 1; k < vData.length; k++) {
        if (vData[k][0] == d.vehicleId) {
          sheetV.getRange(k + 1, 8).setValue(Number(d.endMileage)); break;
        }
      }
      return { success: true };
    }
  }
  return { success: false };
}

function deleteVehicleUsage(id) {
  var sheet = getOrCreateSheet('Vehicle_Usage');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

// ฟังก์ชันสำหรับแก้ไขข้อมูล
function updateRecord(sheetName, idColumnIndex, recordId, newData) {
  try {
    let ss = SpreadsheetApp.openById(DB_ID);
    let sheet = ss.getSheetByName(sheetName);
    let data = sheet.getDataRange().getValues();
    
    // วนลูปหาแถวที่มี ID ตรงกัน (เริ่มที่ 1 เพื่อข้ามหัวตาราง)
    for(let i = 1; i < data.length; i++) {
       if(data[i][idColumnIndex] == recordId) {
          // i+1 เพราะ Array เริ่มที่ 0 แต่ Row เริ่มที่ 1
          sheet.getRange(i+1, 1, 1, newData.length).setValues([newData]);
          return {success: true, message: "อัปเดตข้อมูลสำเร็จ"};
       }
    }
    return {success: false, message: "ไม่พบรหัสข้อมูลที่ต้องการแก้ไข"};
  } catch(e) {
    return {success: false, message: "Error: " + e.message};
  }
}

// ฟังก์ชันสำหรับลบข้อมูล
function deleteRecord(sheetName, idColumnIndex, recordId) {
  try {
    let ss = SpreadsheetApp.openById(DB_ID);
    let sheet = ss.getSheetByName(sheetName);
    let data = sheet.getDataRange().getValues();
    
    for(let i = 1; i < data.length; i++) {
       if(data[i][idColumnIndex] == recordId) {
          sheet.deleteRow(i+1);
          return {success: true, message: "ลบข้อมูลสำเร็จ"};
       }
    }
    return {success: false, message: "ไม่พบรหัสข้อมูลที่ต้องการลบ"};
  } catch(e) {
    return {success: false, message: "Error: " + e.message};
  }
}

function getDashboardSummary() {
  try {
    let ss = SpreadsheetApp.openById(DB_ID);
    
    // ดึงข้อมูลจากชีตต่างๆ
    let vehicleData = ss.getSheetByName('Vehicles').getDataRange().getValues();
    let usageData = ss.getSheetByName('Vehicle_Usage').getDataRange().getValues();
    let fuelLogData = ss.getSheetByName('Fuel_Log').getDataRange().getValues();
    
    let dashboardStats = [];
    
    // ข้ามแถวที่ 0 (หัวตาราง) เริ่มดึงข้อมูลรถทีละคัน
    for (let i = 1; i < vehicleData.length; i++) {
      let vCode = vehicleData[i][0]; // รหัสรถ
      let vReg = vehicleData[i][3];  // ทะเบียนราชการ
      
      let totalDistance = 0;
      let totalFuelAdded = 0;
      
      // 1. คำนวณระยะทางรวมจากชีต Vehicle_Usage
      for (let j = 1; j < usageData.length; j++) {
        if (usageData[j][1] === vCode && usageData[j][6] !== "" && usageData[j][5] !== "") {
          // (เลขไมล์กลับ - เลขไมล์ก่อนออก)
          let distance = Number(usageData[j][6]) - Number(usageData[j][5]);
          if (distance > 0) totalDistance += distance;
        }
      }
      
      // 2. คำนวณน้ำมันที่เติมรวมจากชีต Fuel_Log
      for (let k = 1; k < fuelLogData.length; k++) {
        if (fuelLogData[k][2] === vCode && fuelLogData[k][5] !== "") {
          totalFuelAdded += Number(fuelLogData[k][5]);
        }
      }
      
      // 3. คำนวณอัตราสิ้นเปลือง (Km/L)
      let consumptionRate = 0;
      if (totalFuelAdded > 0) {
        consumptionRate = (totalDistance / totalFuelAdded).toFixed(2);
      }
      
      // เก็บข้อมูลเข้า Array เพื่อส่งไปแสดงที่หน้าเว็บ
      dashboardStats.push({
        vehicleCode: vCode,
        registration: vReg,
        distanceStr: totalDistance + " กม.",
        fuelStr: totalFuelAdded + " ลิตร",
        kmPerLiter: consumptionRate + " กม./ลิตร"
      });
    }
    
    return { success: true, data: dashboardStats };
    
  } catch(e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.message };
  }
}

// ฟังก์ชันดึงรายชื่อรถจากชีต Vehicles
function getVehicleList() {
  try {
    let ss = SpreadsheetApp.openById(DB_ID);
    let sheet = ss.getSheetByName('Vehicles');
    let data = sheet.getDataRange().getDisplayValues(); // ใช้ getDisplayValues เพื่อให้ได้ข้อความตรงตามที่เห็นในชีต
    
    let vehicleList = [];
    
    // เริ่มวนลูปที่ i = 1 เพื่อข้ามหัวตารางแถวแรก
    for (let i = 1; i < data.length; i++) {
      let vCode = data[i][0]; // คอลัมน์ A : รหัสรถ
      let vReg = data[i][3];  // คอลัมน์ D : ทะเบียนราชการ
      
      // ถ้ามีรหัสรถ ให้เก็บเข้า Array
      if (vCode !== "") {
        vehicleList.push({
          code: vCode,
          label: vCode + " (" + vReg + ")" // รูปแบบที่จะแสดงใน Dropdown
        });
      }
    }
    
    return vehicleList;
    
  } catch(e) {
    return []; // ถ้ามี Error ให้ส่ง Array ว่างกลับไป
  }
}

// ==========================================
// ส่วนจัดการข้อมูล: ประวัติการเติมน้ำมัน (Fuel_Log)
// ==========================================

/**
 * ดึงข้อมูลประวัติการเติมน้ำมันทั้งหมด
 */
// ==========================================
// ส่วนจัดการข้อมูล: ประวัติการเติมน้ำมัน (Fuel_Log)
// ==========================================

function getFuelLogs() {
  try {
    // บังคับดึงข้อมูลจาก DB_ID
    var sheet = getOrCreateSheet('Fuel_Log');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var dStr = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], "GMT+7", "yyyy-MM-dd") : data[i][1];
      result.push({
        logId: String(data[i][0]),
        date: dStr,
        vehicleId: String(data[i][2]),
        vehicleReg: String(data[i][3]),
        type: String(data[i][4]),
        liter: String(data[i][5]),
        mileage: String(data[i][6]),
        recorder: String(data[i][7])
      });
    }
    return result.reverse();
  } catch (e) {
    return [];
  }
}

function saveFuelLog(d) {
  try {
    var sheet = getOrCreateSheet('Fuel_Log');
    var isEdit = false;
    
    if (d.logId) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] == d.logId) {
          sheet.getRange(i+1, 2, 1, 7).setValues([[d.date, d.vehicleId, d.vehicleReg, d.type, d.liter, d.mileage, d.recorder]]);
          isEdit = true; 
          break;
        }
      }
    }
    
    if (!isEdit) {
      var logId = "L" + new Date().getTime();
      sheet.appendRow([logId, d.date, d.vehicleId, d.vehicleReg, d.type, d.liter, d.mileage, d.recorder]);
    }
    
    // อัปเดตเลขไมล์กลับไปที่ฐานข้อมูล Vehicles
    var sheetV = getOrCreateSheet('Vehicles');
    var vData = sheetV.getDataRange().getValues();
    for (var k = 1; k < vData.length; k++) {
      if (vData[k][0] == d.vehicleId) {
        var currentMileage = Number(vData[k][7] || 0);
        if (Number(d.mileage) > currentMileage) {
          sheetV.getRange(k + 1, 8).setValue(Number(d.mileage));
        }
        break;
      }
    }
    return { success: true, message: 'บันทึกประวัติการเติมน้ำมันเรียบร้อย' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteFuelLog(id) {
  try {
    var sheet = getOrCreateSheet('Fuel_Log');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) { 
        sheet.deleteRow(i + 1); 
        return { success: true, message: 'ลบข้อมูลสำเร็จ' }; 
      }
    }
    return { success: false, message: 'ไม่พบรายการที่ต้องการลบ' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// ฟังก์ชันดึงรายชื่อผู้ตรวจ (ดึงจากชีต Users)
// ==========================================
function getApprovers() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const approvers = [];
    
    // เริ่มอ่านข้อมูลบรรทัดที่ 2 เป็นต้นไป (ข้ามหัวตาราง)
    for (let i = 1; i < data.length; i++) {
      // ตรวจสอบว่ามี Username และชื่อ-สกุล
      if (data[i][0] && data[i][3]) {
        approvers.push({ 
          username: data[i][0], 
          name: data[i][3], 
          department: data[i][2] 
        });
      }
    }
    return approvers;
  } catch (error) {
    return [];
  }
}

function getUsers() {
  try {
    // ดึงข้อมูลจากชีตที่ชื่อ 'Users' ใน Active Spreadsheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users'); 
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var result = [];
    
    // ข้ามแถวที่ 1 (หัวตาราง) เริ่มดึงข้อมูลจากแถวที่ 2
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; 
      
      // ส่งข้อมูลโดยจับคู่ชื่อตัวแปรให้ถูกต้องและปลอดภัย (Column D เป็นชื่อจริง, คอลัมน์ C เป็นแผนก)
      result.push({
        username: String(data[i][0]),    // คอลัมน์ A (Index 0) - Username
        name: String(data[i][3] || '-'), // คอลัมน์ D (Index 3) - ยศ ชื่อ-สกุล
        department: String(data[i][2] || '-') // คอลัมน์ C (Index 2) - แผนก
      });
    }
    
    return result; 
  } catch (e) {
    return [];
  }
}

// ----------------------------------------------------
// ระบบขั้นตอนการดำเนินงานจัดซื้อจัดจ้าง (Procurement System)
// ----------------------------------------------------
function getProcSteps(bookId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Book_Procurement');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var steps = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == bookId) {
        var dateVal = (data[i][4] instanceof Date) ? Utilities.formatDate(data[i][4], "GMT+7", "yyyy-MM-dd") : (data[i][4] || "");
        steps.push({
          proc_id: String(data[i][0]),
          book_id: String(data[i][1]),
          step_name: String(data[i][2]),
          amount: data[i][3],
          date: dateVal,
          note: String(data[i][5]),
          responsible: String(data[i][6]),
          timestamp: data[i][7] ? Utilities.formatDate(new Date(data[i][7]), "GMT+7", "dd/MM/yyyy HH:mm") : ""
        });
      }
    }
    return steps;
  } catch (e) {
    return [];
  }
}

// ฟังก์ชันบันทึกและอัปเดตประวัติการจัดหาลงชีต
function saveProcStep(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Book_Procurement_Steps");
    if (!sheet) throw new Error("ไม่พบชีต Book_Procurement_Steps");
    
    const timestamp = new Date(); // บันทึกเวลาปัจจุบัน
    
    // กรณี: "แก้ไขข้อมูลเดิม" (มี stepId ส่งมาด้วย)
    if (data.stepId && data.stepId !== "") {
      const sheetData = sheet.getDataRange().getDisplayValues();
      let rowIndex = -1;
      
      // วนลูปหาว่ารหัส step_id อยู่บรรทัดที่เท่าไหร่
      for (let i = 1; i < sheetData.length; i++) {
        if (sheetData[i][0] === data.stepId) {
          rowIndex = i + 1; // +1 เพราะ Array เริ่มที่ 0 แต่แถวในชีตเริ่มที่ 1
          break;
        }
      }
      
      if (rowIndex > -1) {
        // อัปเดตข้อมูลทับบรรทัดเดิม (ข้ามคอลัมน์ A, B เพราะเป็น ID)
        sheet.getRange(rowIndex, 3).setValue(data.stepName);   // คอลัมน์ C
        sheet.getRange(rowIndex, 4).setValue(data.amount);     // คอลัมน์ D
        sheet.getRange(rowIndex, 5).setValue(data.responsible); // คอลัมน์ E
        sheet.getRange(rowIndex, 6).setValue(data.date);       // คอลัมน์ F
        sheet.getRange(rowIndex, 7).setValue(data.note);       // คอลัมน์ G
        sheet.getRange(rowIndex, 8).setValue(timestamp);       // คอลัมน์ H (อัปเดตเวลาล่าสุด)
        return { success: true, message: "อัปเดตประวัติเรียบร้อย" };
      }
    }
    
    // กรณี: "เพิ่มรายการใหม่"
    // สร้างรหัสอ้างอิงสุ่ม (step_id) ไม่ซ้ำกัน
    const newStepId = "STEP-" + new Date().getTime(); 
    
    // เรียงข้อมูลตามคอลัมน์: A(step_id), B(book_id), C(step_name), D(amount), E(responsible), F(date), G(note), H(timestamp)
    const newRow = [
      newStepId, 
      data.bookId, 
      data.stepName, 
      data.amount, 
      data.responsible, 
      data.date, 
      data.note, 
      timestamp
    ];
    
    sheet.appendRow(newRow); // แทรกแถวใหม่ด้านล่างสุดของชีต
    return { success: true, message: "เพิ่มขั้นตอนใหม่เรียบร้อย" };
    
  } catch (error) {
    throw new Error(error.toString());
  }
}

function deleteProcStep(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Book_Procurement');
    if (!sheet) return { success: false, message: "ไม่พบตารางข้อมูล Book_Procurement" };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "ลบขั้นตอนสำเร็จ" };
      }
    }
    return { success: false, message: "ไม่พบข้อมูลที่ต้องการลบ" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// Alias functions สำหรับ frontend ที่เรียกชื่อต่างกัน
// ฟังก์ชันสำหรับดึงประวัติเพื่อส่งกลับไปให้หน้าเว็บ
// ฟังก์ชันดึงประวัติการทำงาน (อิงจากชีต Book_Procurement_Steps)
// 2. ดึงประวัติ (ใช้โค้ดตัวเดิม)
function getProcHistory(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Book_Procurement_Steps"); 
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getDisplayValues();
    const historyList = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === id) { 
        historyList.push({
          stepId: data[i][0], stepName: data[i][2], amount: data[i][3], 
          user: data[i][4], date: data[i][5], note: data[i][6] 
        });
      }
    }
    return historyList;
  } catch (error) {
    return [];
  }
} 

// 3. ฟังก์ชันดึงรายชื่อสมาชิกเพื่อไปใส่ใน Dropdown ผู้รับผิดชอบ
// 3. ฟังก์ชันดึงรายชื่อผู้รับผิดชอบจากชีต Users (ค้นหาจากคอลัมน์ name_title)
function getMembersList() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if(!sheet) return ["ไม่พบชีต Users"];
    
    const data = sheet.getDataRange().getDisplayValues();
    if(data.length < 2) return []; // ถ้ามีแต่หัวข้อ ไม่มีข้อมูล
    
    const headers = data[0];
    let nameColIndex = -1;
    
    // ค้นหาตำแหน่งคอลัมน์ที่มีคำว่า "name_title"
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].toString().trim() === "name_title") {
        nameColIndex = i;
        break;
      }
    }
    
    // ถ้าหาหัวข้อไม่เจอ ให้สุ่มใช้คอลัมน์ B (Index 1) ไปก่อนเพื่อป้องกัน Error
    if (nameColIndex === -1) {
      nameColIndex = 1; 
    }
    
    let members = [];
    for(let i = 1; i < data.length; i++) {
      const name = data[i][nameColIndex]; 
      if(name && name.trim() !== "") {
        members.push(name.trim());
      }
    }
    
    // ส่งรายชื่อกลับไปโดยลบชื่อที่ซ้ำกันออก (ถ้ามี)
    return [...new Set(members)];
    
  } catch(e) {
    return ["เกิดข้อผิดพลาดในการดึงรายชื่อ"];
  }
}

// ฟังก์ชันดึงรายละเอียดเอกสารหลักมาโชว์ในฟอร์ม (เช่น ชื่อเรื่อง)
// 1. ดึงรายละเอียดหนังสือหลัก (ชื่อเรื่อง) จากชีต Book_Register
// 1. ดึงรายละเอียดเอกสาร (ดึงจำนวนเงินจากคอลัมน์ K 'price')
function getProcBookDetails(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Book_Register"); 
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) { // คอลัมน์ A (Index 0) คือ book_id
        return {
          title: data[i][7],  // คอลัมน์ H (Index 7) คือ เรื่อง
          amount: data[i][10] // คอลัมน์ K (Index 10) คือ price (จำนวนเงิน)
        };
      }
    }
    return null;
  } catch (error) {
    return null;
  }
}
function submitProcStep(d) { return saveProcStep(d); }

// ฟังก์ชันสำหรับลบประวัติขั้นตอนดำเนินการ
function deleteProcStepData(stepId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Book_Procurement_Steps");
    if (!sheet) throw new Error("ไม่พบชีต Book_Procurement_Steps");

    const data = sheet.getDataRange().getDisplayValues();
    
    // วนลูปหาแถวที่มี step_id ตรงกับที่ส่งมา (เริ่มจาก 1 เพราะข้าม Header)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === stepId) { // คอลัมน์ A (Index 0) คือ step_id
        sheet.deleteRow(i + 1); // +1 เพราะแถวในชีตจริงๆ เริ่มนับที่ 1
        return { success: true, message: "ลบประวัติสำเร็จ" };
      }
    }
    return { success: false, message: "ไม่พบข้อมูลที่ต้องการลบในระบบ" };
  } catch (error) {
    throw new Error(error.toString());
  }
}

// ฟังก์ชันสำหรับตรวจสอบการเข้าสู่ระบบ
function checkLogin(username, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sheet) throw new Error("ไม่พบชีต Users");

    // ดึงข้อมูลทั้งหมดในชีตมาตรวจสอบ
    const data = sheet.getDataRange().getDisplayValues();

    // วนลูปเช็คทีละบรรทัด (เริ่มจาก 1 เพื่อข้ามหัวคอลัมน์)
    for (let i = 1; i < data.length; i++) {
      
      // ดึงข้อมูลตามลำดับ: A=0(User), B=1(Pass), C=2(Dept), D=3(Name), E=4(Pic)
      let sheetUser = String(data[i][0]).trim();
      let sheetPass = String(data[i][1]).trim();
      let sheetDept = String(data[i][2]).trim();
      let sheetName = String(data[i][3]).trim();
      let sheetPic  = String(data[i][4]).trim();

      // ถ้า Username และ Password ตรงกัน
      if (sheetUser === String(username).trim() && sheetPass === String(password).trim()) {
        
        // กำหนดสิทธิ์เริ่มต้นเป็น User ธรรมดา
        let userRole = "User"; 
        
        // 🔴 เงื่อนไขแต่งตั้ง Admin อัตโนมัติ (ถ้า Username คือ admin หรืออยู่แผนก ผกบ.)
        if (sheetUser.toLowerCase() === "admin" || sheetDept.indexOf("ผกบ") !== -1) {
          userRole = "Admin";
        }

        // ส่งข้อมูลกลับไปให้หน้าเว็บ
        return {
          success: true,
          user: {
            username: sheetUser,
            password: sheetPass,
            department: sheetDept,
            name: sheetName,
            profileUrl: sheetPic,
            role: userRole // ส่งสิทธิ์การใช้งานกลับไปด้วย
          }
        };
      }
    }
    
    // ถ้าวนลูปจนจบแล้วยังไม่เจอที่ตรงกัน
    return { success: false, message: "Username หรือ Password ไม่ถูกต้อง" };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ดึงข้อมูลสรุปน้ำมันแยกตามประเภท


// ฟังก์ชันคำนวณยอดน้ำมันคงคลัง (ส่งให้การ์ดและกราฟโดนัท)
// ฟังก์ชันคำนวณยอดน้ำมันคงคลัง (ดึงประวัติมาบวกลบ แล้วแยกกลุ่มให้การ์ด+กราฟ)
function getFuelDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stock = {}; 

    // 🟢 1. ดึงยอด "รับเข้า" จากชีต Fuel_Depot
    var depotSheet = ss.getSheetByName('Fuel_Depot'); 
    if (depotSheet) {
      var depotData = depotSheet.getDataRange().getValues();
      for (var i = 1; i < depotData.length; i++) {
        var type = String(depotData[i][4]).trim();   // คอลัมน์ E (ชนิดน้ำมัน)
        var amount = Number(depotData[i][5]) || 0;   // คอลัมน์ F (จำนวนลิตร)
        var status = String(depotData[i][8]).trim(); // คอลัมน์ I (สถานะ)

        if (type !== "" && amount > 0 && status.indexOf('รับน้ำมัน') !== -1) {
          if (stock[type] === undefined) stock[type] = 0;
          stock[type] += amount; 
        }
      }
    }

    // 🔴 2. ดึงยอด "จ่ายออก" จากชีต Fuel_Log (อัปเดตตามภาพล่าสุด!)
    var logSheet = ss.getSheetByName('Fuel_Log');
    if (logSheet) {
      var logData = logSheet.getDataRange().getValues();
      for (var j = 1; j < logData.length; j++) {
        var typeLog = String(logData[j][4]).trim();  // คอลัมน์ E (ชนิดน้ำมัน)
        var amountLog = Number(logData[j][5]) || 0;  // คอลัมน์ F (จำนวนลิตรที่เติม)

        // ถ้ามีการเติมน้ำมัน ให้หักออกจากสต๊อก
        if (typeLog !== "" && amountLog > 0) {
          if (stock[typeLog] === undefined) stock[typeLog] = 0;
          stock[typeLog] -= amountLog; 
        }
      }
    }

    // 🔵 3. สรุปยอดแยกประเภท
    var dieselTotal = 0;
    var benzeneTotal = 0;
    
    for (var key in stock) {
      if (key.indexOf('ดีเซล') !== -1) {
        dieselTotal += stock[key];
      } else if (key.indexOf('เบนซิน') !== -1 || key.indexOf('แก๊สโซฮอล์') !== -1) {
        benzeneTotal += stock[key];
      }
    }

    // ส่งค่ากลับไปให้หน้าเว็บ
    return { 
      success: true, 
      dieselFinal: dieselTotal,    
      benzeneFinal: benzeneTotal,  
      stockData: stock             
    };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
function getTodayData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ชื่อชีต');
  var data = sheet.getDataRange().getValues();
  
  // 1. ดึงวันที่วันนี้ แล้วแปลงเป็นรูปแบบข้อความ (เช่น 02/06/2026)
  var today = new Date();
  var todayString = Utilities.formatDate(today, "Asia/Bangkok", "dd/MM/yyyy"); 

  var filteredData = [];

  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0]; // สมมติว่าคอลัมน์ A (Index 0) คือวันที่

    // เช็คว่าช่องนั้นมีค่าและเป็นรูปแบบวันที่จริงๆ
    if (rowDate && typeof rowDate.getTime === 'function') {
      
      // 2. แปลงวันที่ในตารางให้เป็นรูปแบบข้อความเหมือนกัน
      var rowDateString = Utilities.formatDate(rowDate, "Asia/Bangkok", "dd/MM/yyyy");

      // 3. นำข้อความมาเทียบกัน ถ้าตรงกันแปลว่าเป็น "วันนี้"
      if (rowDateString === todayString) {
        filteredData.push(data[i]); // เก็บข้อมูลบรรทัดนี้ไว้
      }
    }
  }

  return filteredData;
}

function getVehiclesData() {
  try {
    // 🔴 ดึงข้อมูลจากชีตที่ชื่อ 'Vehicles'
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Vehicles'); 
    
    if (!sheet) {
      return [];
    }

    var data = sheet.getDataRange().getValues();
    data.shift(); // ตัดแถวแรก (หัวตาราง) ทิ้งไป
    
    return data; // ส่งข้อมูลเป็น Array กลับไปให้หน้าเว็บ
    
  } catch (error) {
    return [];
  }
}

function autoConvertDriveImageToBase64() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users"); 
  var data = sheet.getDataRange().getValues();
  
  // สมมติว่าคอลัมน์ "รููปประจำตัว" อยู่ที่คอลัมน์ E (Index 4)
  var profileColIndex = 4; 
  
  for (var i = 1; i < data.length; i++) {
    var picUrl = data[i][profileColIndex].toString();
    
    // ตรวจสอบว่าเป็นลิงก์ Google Drive
    if (picUrl.indexOf("drive.google.com") > -1) {
      try {
        var fileIdMatch = picUrl.match(/[-\w]{25,}/); 
        if (fileIdMatch) {
          var fileId = fileIdMatch[0];
          
          // 🎯 จุดที่แก้: สั่งให้ไปดึงภาพย่อ (Thumbnail) ขนาด 150px เพื่อลดจำนวนตัวอักษรไม่ให้เกิน 50,000
          var thumbnailUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w150";
          var response = UrlFetchApp.fetch(thumbnailUrl);
          var imageBlob = response.getBlob();
          var base64Data = Utilities.base64Encode(imageBlob.getBytes());
          
          var finalBase64 = "data:image/jpeg;base64," + base64Data;
          
          // ตรวจสอบความปลอดภัยอีกชั้น ถ้าโค้ดยังเกิน 50,000 ให้เซฟเป็นลิงก์ Thumbnail แทน
          if (finalBase64.length > 50000) {
             sheet.getRange(i + 1, profileColIndex + 1).setValue(thumbnailUrl);
          } else {
             sheet.getRange(i + 1, profileColIndex + 1).setValue(finalBase64);
          }
        }
      } catch (e) {
        Logger.log("เกิดข้อผิดพลาดแถวที่ " + (i+1) + ": " + e.toString());
      }
    }
  }
}

function getVehicleChartData() {
  // จำลองข้อมูล (คุณต้องเปลี่ยนไปดึงจาก Spreadsheet ของคุณ)
  var labels = ["08:00", "10:00", "12:00", "14:00", "16:00"];
  var values = [2, 5, 3, 7, 4];
  
  return { labels: labels, values: values };
}

function getBooksData() { // ⚠️ เช็คชื่อฟังก์ชันนี้ให้ตรงกับของคุณด้วยนะครับ
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var bookSheet = ss.getSheetByName('Book_Register');
    var stepSheet = ss.getSheetByName('Book_Procurement_Steps');

    var booksData = bookSheet.getDataRange().getValues();
    var stepsData = stepSheet.getDataRange().getValues();

    // 🟢 1. ส่วนจัดอันดับขั้นตอน
    var stepRanking = {
      "เสนอความต้องการ": 1,
      "ขออนุมัติหลักการ": 2,
      "รายงานขอซื้อ/ขอจ้าง": 3,
      "ขออนุมัติซื้อ/จ้าง": 4,
      "ออกใบสั่งซื้อ/สั่งจ้าง": 5,
      "ตรวจรับ": 6,
      "เบิกเงิน": 7
    };

    var latestStepsObj = {}; 
    var maxRankObj = {};     

    // วนลูปชีต Steps (เริ่มที่ 1 เพื่อข้ามหัวตาราง)
    for (var i = 1; i < stepsData.length; i++) {
      // 🎯 ใช้ String(...).trim() เพื่อลบช่องว่างที่อาจซ่อนอยู่ทั้งหน้าและหลัง
      var stepBookId = String(stepsData[i][1]).trim(); 
      var rawStepName = String(stepsData[i][2]).trim();
      
      // 🎯 ตัดตัวเลขด้านหน้าออก (ถ้ามี เช่น "1. เสนอ..." ให้เหลือแค่ "เสนอ...") เพื่อให้เทียบ Ranking ได้
      var stepNameClean = rawStepName.replace(/^[0-9]+\.\s*/, '').trim(); 
      
      if (stepBookId !== "") {
        var currentRank = stepRanking[stepNameClean] || 0;
        
        if (!maxRankObj[stepBookId] || currentRank >= maxRankObj[stepBookId]) {
          maxRankObj[stepBookId] = currentRank;   
          latestStepsObj[stepBookId] = rawStepName; // เก็บชื่อดั้งเดิมเอาไว้โชว์
        }
      }
    }

    // 🔴 2. ดึงข้อมูลหนังสือ และเอาขั้นตอนมาประกบ
    var result = [];
    for (var j = 1; j < booksData.length; j++) {
      // 🎯 ลบช่องว่างของรหัสหนังสือ เพื่อให้ตรงกับชีต Steps เป๊ะๆ
      var bookId = String(booksData[j][0]).trim(); 
      
      if (bookId !== "") {
        result.push({
          book_id: bookId,
          วันที่: booksData[j][1],
          ทะเบียนรับ: booksData[j][2],
          กน: booksData[j][3],
          ที่: booksData[j][4],
          ลงวันที่: booksData[j][5],
          จาก: booksData[j][6],
          เรื่อง: booksData[j][7],
          status: booksData[j][8],
          current_dept: booksData[j][9],
          price: booksData[j][10], 
          book_type: booksData[j][11],
          
          // 🎯 ดึงขั้นตอนล่าสุดมาใส่ (ถ้าจับคู่ถูก ข้อมูลจะมาทันที)
          latest_step: latestStepsObj[bookId] || 'ยังไม่มีการดำเนินการ' 
        });
      }
    }

    return result; 

  } catch (error) {
    return [];
  }
}

function setupRepairDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "ทะเบียนส่งซ่อม";
  
  let sheet = ss.getSheetByName(sheetName);
  
  // ==========================================
  // กรณีที่ 1: ยังไม่มีชีต ให้สร้างใหม่ทั้งหมด
  // ==========================================
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    const headers = [
      "ที่ใบส่งซ่อม", 
      "วันเวลาที่บันทึก", 
      "วันที่ส่งซ่อม", 
      "สป. สาย", 
      "หน่วยส่งซ่อม", 
      "รายการสิ่งอุปกรณ์", 
      "หมายเลขสิ่งอุปกรณ์", 
      "จำนวน", 
      "หน่วยนับ", 
      "อาการชำรุด", 
      "สถานะงานซ่อม",
      "ผลการดำเนินการ"  // <-- เพิ่มคอลัมน์ใหม่มารองรับการตอบกลับ
    ];
    
    // กำหนดความกว้างคอลัมน์ให้สมดุล (เพิ่ม 250px สำหรับผลการดำเนินการ)
    const columnWidths = [100, 150, 100, 80, 150, 200, 120, 80, 80, 250, 150, 250];
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#0f172a"); 
    headerRange.setFontColor("#ffffff");
    headerRange.setHorizontalAlignment("center");
    
    sheet.setFrozenRows(1);
    
    columnWidths.forEach((width, index) => {
      sheet.setColumnWidth(index + 1, width);
    });
    
    console.log("สร้างฐานข้อมูล 'ทะเบียนส่งซ่อม' สำเร็จ!");
    return "สร้างฐานข้อมูลสำเร็จ!";
  } 
  
  // ==========================================
  // กรณีที่ 2: มีชีตอยู่แล้ว ให้เช็กและเติมคอลัมน์ที่ขาด
  // ==========================================
  else {
    const lastCol = sheet.getLastColumn();
    // ถ้าชีตว่างเปล่าไม่มีหัวตารางเลย
    if (lastCol === 0) return "ชีตว่างเปล่า กรุณาลบชีตนี้ทิ้งแล้วรันใหม่";
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // เช็กว่ามีคอลัมน์ "ผลการดำเนินการ" หรือยัง ถ้ายังให้สร้างเพิ่ม
    if (!headers.includes("ผลการดำเนินการ")) {
      const newColIndex = lastCol + 1;
      
      sheet.getRange(1, newColIndex).setValue("ผลการดำเนินการ");
      sheet.getRange(1, newColIndex).setBackground("#0f172a").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setColumnWidth(newColIndex, 250);
      
      console.log("อัปเกรดฐานข้อมูล: เพิ่มคอลัมน์ 'ผลการดำเนินการ' เรียบร้อยแล้ว");
      return "อัปเดตฐานข้อมูลสำเร็จ (เพิ่มคอลัมน์ใหม่)!";
    }
    
    console.log("โครงสร้างฐานข้อมูล 'ทะเบียนส่งซ่อม' เป็นเวอร์ชันล่าสุดอยู่แล้ว");
    return "มีฐานข้อมูลอยู่แล้วและสมบูรณ์ดี";
  }
}

/**
 * ระบบตรวจสอบบัญชีคุม สป. ทั่วไป (ส่วนระบบทะเบียนส่งซ่อม)
 * พัฒนาระบบ : ร.ท.ยงยุทธ์ เมืองกลาง
 * ปฏิบัติหน้าที่ ประจำแผนก ผกบ.กอ.ฯ โทร.ทบ. 23396
 */

const REPAIR_SHEET_NAME = 'ทะเบียนส่งซ่อม'; // ชื่อชีตที่จะใช้เก็บข้อมูล

// ==========================================
// 1. ฟังก์ชันสำหรับบันทึกข้อมูลใหม่ลง Google Sheets
// ==========================================
function saveRepairToSheet(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(REPAIR_SHEET_NAME);
    
    // หากยังไม่มีชีต "ทะเบียนส่งซ่อม" ให้ระบบสร้างให้ใหม่โดยอัตโนมัติ
    if (!sheet) {
      sheet = ss.insertSheet(REPAIR_SHEET_NAME);
      
      // กำหนดหัวตาราง (Headers)
      const headers = [
        "ที่ใบส่งซ่อม",
        "วันเวลาที่บันทึก", 
        "วันที่ส่งซ่อม", 
        "สป. สาย", 
        "หน่วยส่งซ่อม", 
        "รายการสิ่งอุปกรณ์", 
        "หมายเลขสิ่งอุปกรณ์", 
        "จำนวน", 
        "หน่วยนับ", 
        "อาการชำรุด", 
        "สถานะงานซ่อม" // คอลัมน์พิเศษสำหรับติดตามสถานะ
      ];
      
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#1e293b"); // สีพื้นหลังหัวตาราง
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
      
      sheet.setFrozenRows(1); // แช่แข็งแถวแรกไว้
      
      // ปรับความกว้างคอลัมน์เบื้องต้น
      sheet.setColumnWidth(1, 150); // วันเวลา
      sheet.setColumnWidth(6, 250); // รายการสิ่งอุปกรณ์
      sheet.setColumnWidth(10, 300); // อาการชำรุด
    }
    
    // เตรียมข้อมูลเป็น Array ตามลำดับหัวตาราง
    const timestamp = new Date();
    const initialStatus = "รอรับเรื่อง"; // สถานะเริ่มต้นเมื่อบันทึก
    
    const newRow = [
      data.docNo,      // ย้าย ที่ใบส่งซ่อม มาเป็นคอลัมน์ที่ 1
      timestamp,       // ย้าย วันเวลาที่บันทึก มาเป็นคอลัมน์ที่ 2
      data.date,
      data.category,
      data.unit,
      data.itemName,
      data.itemSerial,
      data.qty,
      data.unitType,
      data.symptom,
      initialStatus
    ];
    
    // บันทึกข้อมูลลงในแถวสุดท้าย
    sheet.appendRow(newRow);
    
    // ตีเส้นกรอบ (Border) ให้กับข้อมูลที่เพิ่งเพิ่มเข้าไป
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    sheet.getRange(lastRow, 1, 1, lastCol).setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
    
    return { success: true, message: 'บันทึกข้อมูลส่งซ่อมสำเร็จ' };
    
  } catch (error) {
    console.error("Error in saveRepairToSheet: ", error);
    return { success: false, message: 'เกิดข้อผิดพลาดฝั่งเซิร์ฟเวอร์: ' + error.message };
  }
}


// ==========================================
// 2. (เพิ่มเติม) ฟังก์ชันสำหรับดึงข้อมูลกลับไปแสดงในตารางหน้าเว็บ
// ==========================================
function getRepairRecords() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(REPAIR_SHEET_NAME);
    
    // ถ้ายังไม่มีชีต หรือไม่มีข้อมูล (มีแค่หัวตาราง) ให้ส่ง Array ว่างกลับไป
    if (!sheet || sheet.getLastRow() <= 1) return []; 
    
    // ดึงข้อมูลทั้งหมดมาเป็น Array 2 มิติ
    const data = sheet.getDataRange().getDisplayValues(); 
    const headers = data[0];
    const records = [];
    
    // แปลง Array 2 มิติ ให้กลายเป็น Array of Objects เพื่อให้ JavaScript จัดการง่าย
    // ไปที่ Code.gs หาฟังก์ชัน getRepairRecords แล้วแก้ตรงวนลูปครับ
for (let i = 1; i < data.length; i++) {
  let row = data[i];
  let obj = {};
  for (let j = 0; j < headers.length; j++) {
    // [จุดที่แก้] เพิ่ม .trim() เพื่อตัดช่องว่างหน้า/หลังหัวตารางทิ้ง
    let cleanHeader = String(headers[j]).trim(); 
    obj[cleanHeader] = row[j];
  }
  records.push(obj);
}
    
    // ส่งข้อมูลกลับไป โดยเรียงลำดับจากล่าสุดไปเก่าสุด (อิงจากแถว)
    return records.reverse(); 
    
  } catch (error) {
    console.error("Error in getRepairRecords: ", error);
    return [];
  }
}

// ฟังก์ชันแก้ไขอัปเดตข้อมูล
function updateRepairRecord(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ทะเบียนส่งซ่อม');
    const records = sheet.getDataRange().getDisplayValues();
    const headers = records[0];
    
    const docNoIdx = headers.indexOf('ที่ใบส่งซ่อม');
    if(docNoIdx === -1) return { success: false, message: 'โครงสร้างชีตไม่ถูกต้อง' };
    
    // วนลูปหาว่าข้อมูลเดิมอยู่แถวไหน
    for(let i=1; i<records.length; i++) {
      if(String(records[i][docNoIdx]) === String(data.originalDocNo)) {
        const row = i + 1; // ตำแหน่งแถวจริงใน Sheet
        
        // อัปเดตข้อมูลใหม่ทับลงไป
        sheet.getRange(row, headers.indexOf('ที่ใบส่งซ่อม') + 1).setValue(data.docNo);
        sheet.getRange(row, headers.indexOf('วันที่ส่งซ่อม') + 1).setValue(data.date);
        sheet.getRange(row, headers.indexOf('สป. สาย') + 1).setValue(data.category);
        sheet.getRange(row, headers.indexOf('หน่วยส่งซ่อม') + 1).setValue(data.unit);
        sheet.getRange(row, headers.indexOf('รายการสิ่งอุปกรณ์') + 1).setValue(data.itemName);
        sheet.getRange(row, headers.indexOf('หมายเลขสิ่งอุปกรณ์') + 1).setValue(data.itemSerial);
        sheet.getRange(row, headers.indexOf('จำนวน') + 1).setValue(data.qty);
        sheet.getRange(row, headers.indexOf('หน่วยนับ') + 1).setValue(data.unitType);
        sheet.getRange(row, headers.indexOf('อาการชำรุด') + 1).setValue(data.symptom);
        
        return { success: true, message: 'บันทึกการแก้ไขข้อมูลสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลเดิมที่ต้องการแก้ไข' };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ฟังก์ชันลบข้อมูล
function deleteRepairRecord(docNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ทะเบียนส่งซ่อม');
    const records = sheet.getDataRange().getDisplayValues();
    const headers = records[0];
    const docNoIdx = headers.indexOf('ที่ใบส่งซ่อม');
    
    // วนลูปหาข้อมูลเพื่อลบ
    for(let i=1; i<records.length; i++) {
      if(String(records[i][docNoIdx]) === String(docNo)) {
        sheet.deleteRow(i + 1); // ลบแถวทิ้ง
        return { success: true, message: 'ลบข้อมูลสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลที่ต้องการลบ' };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ฟังก์ชันสำหรับบันทึกผลการส่งซ่อม (สถานะ และ รายละเอียด)
// ฟังก์ชันสำหรับบันทึกผลการส่งซ่อม และอัปโหลดไฟล์หลักฐานลง Google Drive
// ฟังก์ชันสำหรับบันทึกผลการส่งซ่อม และอัปโหลดไฟล์หลักฐานลง Google Drive (เวอร์ชันแก้บั๊กลิงก์ไม่ลง Sheet)
// นำโค้ดนี้ไปวางทับฟังก์ชันเดิมใน Code.gs
function updateRepairOutcome(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ทะเบียนส่งซ่อม');
    const records = sheet.getDataRange().getDisplayValues();
    const headers = records[0];

    const docNoIdx = headers.findIndex(h => String(h).trim() === 'ที่ใบส่งซ่อม');
    const statusIdx = headers.findIndex(h => String(h).trim() === 'สถานะงานซ่อม');
    const detailIdx = headers.findIndex(h => String(h).trim() === 'ผลการดำเนินการ');
    const linkIdx = headers.findIndex(h => String(h).trim() === 'ลิงก์หลักฐาน');

    // จัดการไฟล์
    let uploadedFileUrl = "";
    if (data.file && data.file.base64) {
      const folder = DriveApp.getFolderById('116voRMbX3XOO4Zu8O_6wCjLjn2aR0_ij');
      const blob = Utilities.newBlob(Utilities.base64Decode(data.file.base64), data.file.mimeType, data.file.filename);
      const newFile = folder.createFile(blob);
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      uploadedFileUrl = newFile.getUrl();
    }

    // อัปเดตข้อมูล
    for (let i = 1; i < records.length; i++) {
      if (String(records[i][docNoIdx]) === String(data.docNo)) {
        const row = i + 1;
        sheet.getRange(row, statusIdx + 1).setValue(data.status);
        sheet.getRange(row, detailIdx + 1).setValue(data.detail);
        if (uploadedFileUrl !== "") {
          sheet.getRange(row, linkIdx + 1).setValue(uploadedFileUrl);
        }
        
        // [สำคัญ] ส่งค่า success กลับไปให้หน้าเว็บ
        return { success: true, message: 'บันทึกสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบเลขใบส่งซ่อมนี้' };

  } catch (error) {
    // [สำคัญ] ถ้าหลังบ้านพัง มันจะส่ง error นี้กลับไปให้หน้าเว็บแสดงผล
    return { success: false, message: 'Server Error: ' + error.toString() };
  }
}

function getBookRadarData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Book_Register");
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; // แถวแรกคือหัวตาราง
  
  // หาดัชนีคอลัมน์ต่างๆ
  const typeIdx = headers.indexOf("book_type"); 
  const titleIdx = headers.indexOf("เรื่อง");
  const idIdx = headers.indexOf("ทะเบียนรับ");
  const bookIdIdx = headers.indexOf("book_id"); // 🔴 เพิ่มการหาคอลัมน์ book_id

  if (typeIdx === -1) {
    return { error: "ไม่พบหัวตาราง 'book_type' ในชีต Book_Register" };
  }

  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx]) { 
      result.push({
        "book_id": data[i][bookIdIdx], // 🔴 ส่งค่า book_id กลับไปด้วย
        "ทะเบียนรับ": data[i][idIdx],
        "เรื่อง": data[i][titleIdx],
        "book_type": data[i][typeIdx] 
      });
    }
  }
  return result;
}

function getProcurementStepsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Book_Procurement_Steps"); 
  
  if (!sheet) {
    return { error: "หาชีต Book_Procurement_Steps ไม่พบครับ" };
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // ถ้ามีแต่หัวตาราง ให้ส่ง Array ว่างกลับไป

  const headers = data[0]; 
  
  // ค้นหาตำแหน่งคอลัมน์ (ใช้ .trim() ตัดเว้นวรรคทิ้ง ป้องกันปัญหาเผลอพิมพ์เว้นวรรคใน Sheet)
  const stepNameIdx = headers.findIndex(h => String(h).trim() === "step_name");
  const titleIdx = headers.findIndex(h => String(h).trim() === "เรื่อง");

  if (stepNameIdx === -1) {
    return { error: "ไม่พบหัวตาราง 'step_name' ในชีต Book_Procurement_Steps" };
  }

  const result = [];
  
  // เริ่มเก็บข้อมูลตั้งแต่บรรทัดที่ 2 เป็นต้นไป
  for (let i = 1; i < data.length; i++) {
    // เช็กว่าคอลัมน์ step_name ของบรรทัดนั้นไม่ได้ว่างเปล่า
    if (data[i][stepNameIdx]) { 
      
      // สร้างชุดข้อมูลเตรียมส่งกลับ
      let rowData = {
        "step_name": String(data[i][stepNameIdx]).trim()
      };
      
      // ถ้าในตารางมีคอลัมน์ "เรื่อง" ให้ดึงข้อมูลเรื่องกลับไปด้วย
      if (titleIdx !== -1 && data[i][titleIdx]) {
        rowData["เรื่อง"] = String(data[i][titleIdx]).trim();
      } else {
        rowData["เรื่อง"] = "ไม่ระบุชื่อเรื่อง";
      }

      result.push(rowData);
    }
  }
  
  return result;
}

// เปลี่ยนจาก .getProcurementTrackingData() เป็น .getProcurementStepsData()
function getProcurementTrackingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Book_Procurement_Steps"); 
  
  if (!sheet) return { error: "หาชีตไม่พบ" };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => String(h).trim()); // จัดการ header ให้สะอาด
  const stepNameIdx = headers.indexOf("step_name");
  const titleIdx = headers.indexOf("เรื่อง");
  
  // Debug: ถ้าหา "เรื่อง" ไม่เจอ ให้ดูใน Console ว่า header ที่อ่านได้คืออะไร
  console.log("Headers found:", headers);

  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][stepNameIdx]) { 
      let title = "ไม่ระบุชื่อเรื่อง";
      if (titleIdx !== -1 && data[i][titleIdx]) {
        title = String(data[i][titleIdx]).trim();
      }
      
      result.push({
        "step_name": String(data[i][stepNameIdx]).trim(),
        "เรื่อง": title
      });
    }
  }
  return result;
}

function getCombinedProcurementData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registerSheet = ss.getSheetByName("Book_Register");
  const stepsSheet = ss.getSheetByName("Book_Procurement_Steps");
  
  if (!registerSheet || !stepsSheet) return { error: "หาชีตไม่พบ" };

  const regData = registerSheet.getDataRange().getValues();
  const stepsData = stepsSheet.getDataRange().getValues();

  // 1. สร้าง Map เก็บชื่อเรื่องจาก Book_Register
  const titleMap = {};
  const regHeaders = regData[0];
  const bookIdIdxReg = regHeaders.indexOf("book_id");
  const titleIdxReg = regHeaders.indexOf("เรื่อง");

  for (let i = 1; i < regData.length; i++) {
    if (regData[i][bookIdIdxReg]) {
      titleMap[String(regData[i][bookIdIdxReg]).trim()] = regData[i][titleIdxReg] || "ไม่มีชื่อเรื่อง";
    }
  }

  const STEPS_ORDER = ["เสนอความต้องการ", "ขออนุมัติหลักการ", "รายงานขอซื้อ/ขอจ้าง", "ขออนุมัติซื้อ/จ้าง", "ออกใบสั่งซื้อ/สั่งจ้าง", "ตรวจรับ", "เบิกเงิน"];
  
  const stepsHeaders = stepsData[0];
  const bookIdIdxSteps = stepsHeaders.indexOf("book_id");
  const stepNameIdxSteps = stepsHeaders.indexOf("step_name");
  const responsibleIdxSteps = stepsHeaders.indexOf("responsible"); // ดึงคอลัมน์ผู้รับผิดชอบ
  
  let dateIdxSteps = stepsHeaders.findIndex(h => {
    let text = String(h).toLowerCase();
    return text.includes("วันที่") || text.includes("timestamp") || text.includes("date") || text.includes("เวลา");
  });
  if (dateIdxSteps === -1) dateIdxSteps = 0; 

  const bookStatus = {}; // เก็บสถานะล่าสุด (สำหรับการ์ด)
  const historyLog = []; // เก็บประวัติทุกบรรทัด (สำหรับ Popup)

  for (let i = 1; i < stepsData.length; i++) {
    const bId = String(stepsData[i][bookIdIdxSteps]).trim();
    const rawStepName = String(stepsData[i][stepNameIdxSteps]).trim();
    const cleanStepName = rawStepName.replace(/^\d+\.\s*/, "");
    const sIndex = STEPS_ORDER.indexOf(cleanStepName);
    const responsible = responsibleIdxSteps !== -1 ? String(stepsData[i][responsibleIdxSteps] || "-").trim() : "-";

    let rawDate = stepsData[i][dateIdxSteps];
    let formattedDate = "-";
    let timestamp = 0; 

    if (rawDate) {
      if (rawDate instanceof Date) {
        formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        timestamp = rawDate.getTime(); 
      } else {
        formattedDate = String(rawDate).split(" ")[0];
        let parsed = new Date(rawDate);
        timestamp = isNaN(parsed.getTime()) ? 0 : parsed.getTime();
      }
    }

    let title = titleMap[bId] || "ไม่พบชื่อเรื่อง";

    if (bId) {
      // 🔴 ส่วนที่ 1: เก็บข้อมูล "ทุกบรรทัด" ลงในประวัติ เพื่อส่งให้ Popup
      historyLog.push({
        title: title,
        stepName: rawStepName,
        date: formattedDate,
        timestamp: timestamp,
        responsible: responsible
      });

      // 🔴 ส่วนที่ 2: คัดกรอง "สถานะล่าสุด" เพื่อส่งให้การ์ดหน้าหลัก
      if (sIndex !== -1) {
        if (!bookStatus[bId] || sIndex >= bookStatus[bId].maxIndex) {
          bookStatus[bId] = { 
            title: title, 
            maxIndex: sIndex, 
            stepName: cleanStepName,
            date: formattedDate,
            timestamp: timestamp
          };
        }
      }
    }
  }

  // จัดเรียงข้อมูลทั้ง 2 ชุด จากล่าสุดไปเก่าสุด
  let summaryData = Object.keys(bookStatus).map(id => bookStatus[id]);
  summaryData.sort((a, b) => b.timestamp - a.timestamp);
  historyLog.sort((a, b) => b.timestamp - a.timestamp);

  // ส่งข้อมูลกลับไปเป็น Object 2 ก้อน
  return {
    summary: summaryData,
    history: historyLog
  };
}

function getRepairDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ทะเบียนส่งซ่อม"); 
  
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const catIdx = headers.indexOf("สป. สาย");
  const statIdx = headers.indexOf("สถานะงานซ่อม");

  let summary = {
    total: 0,
    writeoffCount: 0, // 🔴 ยอดจำหน่ายสำเร็จ
    categories: {}, 
    statuses: {}
  };

  for (let i = 1; i < data.length; i++) {
    let catText = String(data[i][catIdx]).trim() || "ไม่ระบุ";
    let statText = String(data[i][statIdx]).trim() || "ไม่ระบุ";
    if (!catText && !statText && String(data[i][0]).trim() === "") continue;

    summary.total++; 
    summary.categories[catText] = (summary.categories[catText] || 0) + 1;
    summary.statuses[statText] = (summary.statuses[statText] || 0) + 1;

    // 🔴 นับจำหน่ายสำเร็จแยกไว้ต่างหาก
    if (statText.includes("จำหน่ายสำเร็จ")) {
      summary.writeoffCount++;
    }
  }
  return summary;
}

// ฟังก์ชันหาเลขที่ส่งล่าสุด (ของ Outbound_Books) และรีเซ็ตเป็นเลข 1 เมื่อขึ้นปีใหม่
function getNextOutboundNumber() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Outbound_Books"); 
    if (!sheet) return 1;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 1;
    const currentYear = new Date().getFullYear();
    let maxNum = 0;
    for (let i = 1; i < data.length; i++) {
      let rowNum = data[i][0]; 
      let rowDate = data[i][1]; 
      let rowYear = 0;
      if (rowDate instanceof Date) rowYear = rowDate.getFullYear();
      else {
        let dateStr = String(rowDate);
        let match = dateStr.match(/\d{4}/);
        if (match) {
          rowYear = parseInt(match[0], 10);
          if (rowYear > 2500) rowYear -= 543;
        }
      }
      if (rowYear === currentYear) {
        let currentNum = parseInt(String(rowNum).replace(/\D/g, ''), 10);
        if (!isNaN(currentNum) && currentNum > maxNum) maxNum = currentNum;
      }
    }
    return maxNum + 1; // ส่งเลข 31 กลับไปเพียวๆ
  } catch (error) {
    return 1;
  }
}

// ฟังก์ชันหาเลขที่ใบส่งซ่อมล่าสุด
// ฟังก์ชันดึงเลขที่ใบส่งซ่อมล่าสุด
function getNextRepairNumber() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("ทะเบียนส่งซ่อม"); 
    if (!sheet) return 1;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 1;

    const currentYear = new Date().getFullYear();
    let maxNum = 0;

    // คอลัมน์ A (0) = เลขที่, คอลัมน์ C (2) = วันที่ส่งซ่อม
    for (let i = 1; i < data.length; i++) {
      let rowNum = data[i][0]; 
      let rowDate = data[i][2]; 
      
      let rowYear = 0;
      if (rowDate instanceof Date) {
        rowYear = rowDate.getFullYear();
      } else {
        let match = String(rowDate).match(/\d{4}/);
        if (match) {
          rowYear = parseInt(match[0], 10);
          if (rowYear > 2500) rowYear -= 543;
        }
      }

      if (rowYear === currentYear) {
        let currentNum = parseInt(String(rowNum).replace(/\D/g, ''), 10);
        if (!isNaN(currentNum) && currentNum > maxNum) maxNum = currentNum;
      }
    }
    return maxNum + 1; // ส่งเลขเพียวๆ กลับไปที่หน้าเว็บ
  } catch (e) {
    return 1;
  }
}

function getCombinedHistory(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var procSheet = ss.getSheetByName('Book_Procurement_Steps'); 
  var procData = procSheet.getDataRange().getValues();
  var searchId = String(id).trim(); // ID ที่ส่งมาจากหน้าเว็บ
  
  // สร้างรายการ ID ทั้งหมดที่มีในตารางมาให้ดู
  var allIdsInSheet = [];
  for (var k = 1; k < procData.length; k++) {
    allIdsInSheet.push(String(procData[k][1]).trim());
  }
  
  Logger.log("=== เริ่มค้นหาประวัติ ===");
  Logger.log("ID ที่ค้นหาคือ: [" + searchId + "]");
  Logger.log("รายการ ID ทั้งหมดในตาราง (คอลัมน์ B): " + allIdsInSheet.join(", "));
  
  var combinedHistory = [];
  
  // ค้นหาแบบเปรียบเทียบทีละตัว
  for (var j = 1; j < procData.length; j++) {
    var historyDocId = String(procData[j][1]).trim(); 
    
    if (historyDocId == searchId) {
      combinedHistory.push(procData[j]);
      Logger.log("พบรายการที่ Match! แถวที่ " + (j + 1));
    }
  }
  
  Logger.log("สรุป: เจอทั้งหมด " + combinedHistory.length + " แถว");
  return combinedHistory;
}


//----------------------------------------
//กำหนดสิทธิ์การใช้งาน
//---------------------------------------

