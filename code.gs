// กำหนดชื่อ Sheet ที่จะใช้งาน
const SHEET_USERS = 'Users';
const SHEET_STUDENTS = 'Students';

// 1. ฟังก์ชันแสดงหน้าเว็บเมื่อมีการเข้าถึง URL ของ Web App
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบ Phonics ร.ร.นิคมสร้างตนเองแว้ง')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // ทำให้รองรับมือถือ
}

// 2. ฟังก์ชันสร้างหัวตารางและฐานข้อมูลอัตโนมัติ (รันครั้งแรก)
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // สร้าง Sheet Users ถ้ายังไม่มี
  let userSheet = ss.getSheetByName(SHEET_USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(SHEET_USERS);
    userSheet.appendRow(['UserID', 'Username', 'Password', 'Role', 'Title', 'FirstName', 'LastName', 'Homeroom']);
    userSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#22c55e").setFontColor("white"); // แต่งสีหัวตาราง
    
    // สร้างข้อมูลผู้ใช้เริ่มต้น
    userSheet.appendRow(['U001', 'admin', 'password', 'admin', 'นาย', 'ผู้ดูแล', 'ระบบ', 'ทั้งหมด']);
    userSheet.appendRow(['U002', 'director', 'password', 'director', 'นาย', 'สมศักดิ์', 'ผู้อำนวยการ', 'ทั้งหมด']);
    userSheet.appendRow(['U003', 'teacher4', 'password', 'teacher', 'นาง', 'สมศรี', 'ครูป.4', 'ป.4/1']);
  }

  // สร้าง Sheet Students ถ้ายังไม่มี
  let studentSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!studentSheet) {
    studentSheet = ss.insertSheet(SHEET_STUDENTS);
    studentSheet.appendRow(['StudentID', 'Title', 'FirstName', 'LastName', 'Gender', 'Class', 'Room', 'PreScore', 'PostScore']);
    studentSheet.getRange("A1:I1").setFontWeight("bold").setBackground("#22c55e").setFontColor("white");
    
    // สร้างข้อมูลนักเรียนจำลองเบื้องต้น
    studentSheet.appendRow(['664001', 'เด็กชาย', 'รักเรียน', 'เพียรวิชา', 'ชาย', 'ป.4', '1', 3, 8]);
    studentSheet.appendRow(['664002', 'เด็กหญิง', 'ใจดี', 'มีสุข', 'หญิง', 'ป.4', '1', 4, '']);
    studentSheet.appendRow(['665001', 'เด็กชาย', 'กล้าหาญ', 'ชาญชัย', 'ชาย', 'ป.5', '1', 5, 9]);
    studentSheet.appendRow(['666001', 'เด็กหญิง', 'ใฝ่รู้', 'คู่คุณธรรม', 'หญิง', 'ป.6', '1', 2, 7]);
  }
  
  return "ตั้งค่าฐานข้อมูลสำเร็จ";
}

// 3. ฟังก์ชันเข้าสู่ระบบ
function verifyLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if(!sheet) return { success: false, message: "ไม่พบฐานข้อมูล" };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // เช็ค username และ password บรรทัดต่อบรรทัด
    if (data[i][1] == username && data[i][2] == password) {
      return {
        success: true,
        user: {
          id: data[i][0],
          username: data[i][1],
          role: data[i][3],
          title: data[i][4],
          firstName: data[i][5],
          lastName: data[i][6],
          name: data[i][4] + data[i][5] + ' ' + data[i][6],
          homeroom: data[i][7]
        }
      };
    }
  }
  return { success: false, message: "ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง" };
}

// 4. ดึงข้อมูลนักเรียนทั้งหมด
function getStudents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  if(!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < data.length; i++) {
    students.push({
      row: i + 1,
      studentId: data[i][0],
      title: data[i][1],
      firstName: data[i][2],
      lastName: data[i][3],
      name: data[i][1] + data[i][2] + ' ' + data[i][3],
      gender: data[i][4],
      class: data[i][5],
      room: data[i][6],
      preScore: data[i][7],
      postScore: data[i][8]
    });
  }
  return students;
}

// 5. บันทึกคะแนน
function saveScores(scoreData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  
  // scoreData เป็น Array ของ object {row, pre, post}
  scoreData.forEach(item => {
    // อัปเดตเฉพาะช่องที่มีการเปลี่ยนแปลง
    if(item.pre !== undefined) sheet.getRange(item.row, 8).setValue(item.pre === '' ? '' : Number(item.pre));
    if(item.post !== undefined) sheet.getRange(item.row, 9).setValue(item.post === '' ? '' : Number(item.post));
  });
  return { success: true };
}

// 6. จัดการข้อมูลนักเรียน (Admin)
function saveStudent(data, isEdit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  
  if (isEdit) {
    // แก้ไขข้อมูลเดิม
    sheet.getRange(data.row, 1, 1, 7).setValues([[data.studentId, data.title, data.firstName, data.lastName, data.gender, data.class, data.room]]);
  } else {
    // เพิ่มนักเรียนใหม่ (คะแนนเว้นว่าง)
    sheet.appendRow([data.studentId, data.title, data.firstName, data.lastName, data.gender, data.class, data.room, '', '']);
  }
  return { success: true };
}

function deleteStudent(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  sheet.deleteRow(row);
  return { success: true };
}

// 7. จัดการข้อมูลผู้ใช้ (Admin)
function getUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    users.push({
      row: i + 1, id: data[i][0], username: data[i][1], role: data[i][3],
      title: data[i][4], firstName: data[i][5], lastName: data[i][6], homeroom: data[i][7]
    });
  }
  return users;
}

function saveUser(data, isEdit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (isEdit) {
    // ไม่อัปเดตรหัสผ่านถ้าแก้ไขเฉพาะข้อมูล
    sheet.getRange(data.row, 4, 1, 5).setValues([[data.role, data.title, data.firstName, data.lastName, data.homeroom]]);
  } else {
    // สร้างผู้ใช้ใหม่
    const newId = 'U' + new Date().getTime();
    sheet.appendRow([newId, data.username, 'password', data.role, data.title, data.firstName, data.lastName, data.homeroom]);
  }
  return { success: true };
}

function deleteUser(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  sheet.deleteRow(row);
  return { success: true };
}

// 8. อัปเดตข้อมูลส่วนตัว (สำหรับผู้ใช้ทุกคน)
function updateProfile(row, title, firstName, lastName, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  
  // อัปเดตชื่อสกุล
  sheet.getRange(row, 5, 1, 3).setValues([[title, firstName, lastName]]);
  // ถ้ามีการส่งรหัสผ่านใหม่มา ให้อัปเดตรหัสผ่านด้วย
  if(password) sheet.getRange(row, 3).setValue(password);
  
  return { success: true };
}
