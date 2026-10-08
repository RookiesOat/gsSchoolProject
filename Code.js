// ID ของ Google Sheet ที่ใช้เป็นแหล่งข้อมูล
const SHEET_ID = '1TM8cyeHLiM6ybALA-8BoCbVTvtP79g8s0H0683xlftw';

// รับ HTTP GET และเลือกการทำงานตามพารามิเตอร์ action
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : undefined;

    if (action === 'getStudents') {
      return createJsonResponse(getStudentsData());
    }

    // แจ้งข้อผิดพลาดเมื่อไม่ได้ระบุ action หรือระบุค่าที่ไม่รองรับ
    throw new Error('กรุณาระบุ action=getStudents');
  } catch (error) {
    // ส่งข้อผิดพลาดกลับเป็น JSON เช่นเดียวกับผลลัพธ์ปกติ
    return createJsonResponse({ error: error.message });
  }
}

// รับ HTTP POST และเลือกการทำงานตามพารามิเตอร์ action
function doPost(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : undefined;

    if (action !== 'addStudent') {
      throw new Error('กรุณาระบุ action=addStudent');
    }

    // ตรวจสอบและแปลงเนื้อหาในคำขอจาก JSON string เป็น Object
    if (!e.postData || !e.postData.contents) {
      throw new Error('กรุณาส่งข้อมูลนักเรียนในรูปแบบ JSON');
    }
    const data = JSON.parse(e.postData.contents);
    if (!data || typeof data !== 'object' || Array.isArray(data)) {
      throw new Error('ข้อมูลนักเรียนต้องเป็น JSON object');
    }

    // เปิดไฟล์ Google Sheets และเลือกแผ่นงาน Students
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName('Students');
    if (!sheet) {
      throw new Error('ไม่พบแผ่นงานชื่อ Students');
    }

    // เรียงข้อมูลตามหัวตาราง และใช้สถานะเริ่มต้นเมื่อไม่ได้ส่ง status มา
    const newRow = [
      data.studentId,
      data.name,
      data.className,
      data.number,
      data.status === undefined ? 'กำลังเรียน' : data.status
    ];

    // เพิ่มข้อมูลนักเรียนในแถวใหม่ แล้วแจ้งผลสำเร็จเป็น JSON
    sheet.appendRow(newRow);
    return createJsonResponse({
      success: true,
      message: 'เพิ่มข้อมูลนักเรียนสำเร็จ'
    });
  } catch (error) {
    // ส่งข้อผิดพลาดจากการอ่าน JSON หรือบันทึกข้อมูลกลับเป็น JSON
    return createJsonResponse({ success: false, error: error.message });
  }
}

// อ่านข้อมูลจากชีต Students และแปลงแต่ละแถวเป็น Object
function getStudentsData() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = spreadsheet.getSheetByName('Students');

  if (!sheet) {
    throw new Error('ไม่พบแผ่นงานชื่อ Students');
  }

  const rows = sheet.getDataRange().getValues();

  // ถ้าไม่มีข้อมูล หรือมีเพียงหัวตาราง ให้ส่งอาร์เรย์ว่างกลับไป
  if (rows.length <= 1) {
    return [];
  }

  // ใช้แถวแรกเป็นชื่อ Key ของ Object ในแต่ละแถวข้อมูล
  const headers = rows[0].map(function (header) {
    return String(header).trim();
  });

  // ป้องกันข้อมูลสูญหายจากชื่อหัวตารางว่างหรือซ้ำกัน
  if (headers.some(function (header) { return header === ''; })) {
    throw new Error('หัวตารางของชีต Students ต้องไม่มีช่องว่าง');
  }
  if (new Set(headers).size !== headers.length) {
    throw new Error('หัวตารางของชีต Students ต้องไม่ซ้ำกัน');
  }

  return rows.slice(1).map(function (row) {
    const student = Object.create(null);
    headers.forEach(function (header, index) {
      student[header] = row[index];
    });
    return student;
  });
}

// แปลงข้อมูลเป็น JSON string และกำหนดชนิดเนื้อหาเป็น JSON
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------------------------------------------------
// ฟังก์ชันสำหรับสร้างโครงสร้างฐานข้อมูลทั้งหมดตาม UI ที่ออกแบบไว้
// --------------------------------------------------------
function setupDatabase() {
  const db = SpreadsheetApp.openById(SHEET_ID);
  
  // โครงสร้างที่ออกแบบไว้ทั้ง 5 ส่วน
  const schemas = {
    'Students': ['รหัสนักเรียน', 'ชื่อ-สกุล', 'ชั้นเรียน', 'เลขที่', 'สถานะ'],
    'Attendance': ['วันที่', 'ชั้นเรียน', 'รหัสนักเรียน', 'สถานะการมาเรียน', 'หมายเหตุ'],
    'Scores': ['ปีการศึกษา', 'ภาคเรียน', 'ชั้นเรียน', 'วิชา', 'รหัสนักเรียน', 'คะแนนงาน', 'คะแนนทดสอบ', 'คะแนนปลายภาค', 'สถานะส่งตรวจ'],
    'Evaluations': ['ปีการศึกษา', 'ภาคเรียน', 'รอบประเมิน', 'หมวดหมู่', 'รหัสนักเรียน', 'ผลการประเมิน', 'หมายเหตุ'],
    'Tasks': ['วันที่กำหนด', 'ชื่องาน', 'สถานะ', 'ผู้รับผิดชอบ']
  };

  // วนลูปสร้าง Sheet และหัวตาราง
  for (const [sheetName, headers] of Object.entries(schemas)) {
    let sheet = db.getSheetByName(sheetName);
    
    // ถ้ายังไม่มี Sheet ชื่อนี้ ให้สร้างใหม่
    if (!sheet) {
      sheet = db.insertSheet(sheetName);
    }
    
    // ใส่หัวตารางที่แถวแรก
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    // ปรับแต่งตัวอักษรให้เป็นตัวหนาและมีสีพื้นหลังเพื่อความสวยงาม
    headerRange.setFontWeight('bold').setBackground('#f3f4f6');
  }
}
