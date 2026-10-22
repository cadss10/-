/**
 * =====================================================================================
 * ระบบแจ้งปัญหาคอมพิวเตอร์และ IT (IT Helpdesk & Computer Problem Reporting System)
 * Backend Google Apps Script (code.gs)
 * “แจ้งง่าย ติดตามได้ แก้ไขไว ตรวจสอบได้”
 * =====================================================================================
 */

// =====================================================================================
// 1. CONFIGURATION (การตั้งค่าระบบ)
// =====================================================================================
const CONFIG = {
  // หากใช้แบบ Container-bound Script (ผูกกับ Google Sheets โดยตรง) ปล่อยว่างไว้ได้
  // หรือระบุ Spreadsheet ID จาก URL: 
  SPREADSHEET_ID: '',

  // Google Drive Folder ID สำหรับจัดเก็บไฟล์แนบ (URL: https://drive.google.com/drive/folders/1_wJJIN3LExO5jXG50Ms2aJplc6lASDI7)
  DRIVE_FOLDER_ID: '1_wJJIN3LExO5jXG50Ms2aJplc6lASDI7',

  // Telegram Bot Token ที่ได้จาก @BotFather (เช่น '123456789:ABCdefGhIJKlmNoPQRsTUVwxyZ')
  TELEGRAM_BOT_TOKEN: '',

  // Telegram Chat ID หรือ Group ID (เช่น '123456789' หรือ '-100123456789')
  TELEGRAM_CHAT_ID: '',

  // Google Chat Incoming Webhook URL
  GOOGLE_CHAT_WEBHOOK_URL: '',

  // อีเมลกลางของ Admin / Superadmin สำหรับรับแจ้งเตือนเคสใหม่ (คั่นด้วยเครื่องหมายจุลภาค ,)
  // หากเว้นว่างไว้ ระบบจะดึงอีเมลของเจ้าหน้าที่บทบาท admin และ superadmin ในชีต Users โดยอัตโนมัติ
  ADMIN_NOTIFICATION_EMAILS: '',

  // เปิด/ปิดการแจ้งเตือนทางอีเมล (true/false)
  ENABLE_EMAIL_NOTIFICATION: true,

  // URL สำหรับลิงก์เข้าสู่ระบบ IT Helpdesk ในอีเมล
  WEB_APP_URL: '',

  // Timezone สำหรับระบบ
  TIMEZONE: 'Asia/Bangkok'
};

// รายชื่อชีตในระบบ
const SHEETS = {
  TICKETS: 'Tickets',
  WORKLOGS: 'WorkLogs',
  USERS: 'Users',
  SETTINGS: 'Settings',
  COOPERATIVES: 'ชื่อสหกรณ์'
};

// Headers สำหรับแต่ละชีต
const HEADERS = {
  TICKETS: [
    'TicketID', 'Date', 'Time', 'CreatedAt', 'ReporterName', 
    'Position', 'Phone', 'Email', 'CoopName', 'CoopId', 
    'CoopType', 'AuditorName', 'AuditorType', 'SoftwareList', 'ProblemTitle', 
    'ProblemDetail', 'Priority', 'Status', 'AssignedTo', 'AssignedAt', 
    'AcceptedAt', 'StartedAt', 'ResolvedAt', 'ClosedAt', 'Resolution', 
    'RootCause', 'Cost', 'Rating', 'Feedback', 'AttachmentURL', 'LastUpdated'
  ],
  WORKLOGS: [
    'LogID', 'TicketID', 'DateTime', 'Staff', 'Action', 
    'Parts', 'Cost', 'Note'
  ],
  USERS: [
    'UserID', 'Name', 'Department', 'Email', 'Phone', 
    'Role', 'Status'
  ],
  SETTINGS: [
    'Key', 'Value'
  ],
  COOPERATIVES: [
    'รหัสสหกรณ์', 'ประเภทสหกรณ์', 'ชื่อสหกรณ์', 'ชื่อ-นามสกุล ผู้สอบบัญชี', 'ประเภท ผู้สอบบัญชี'
  ]
};

// =====================================================================================
// 2. WEB APP ENTRY POINT
// =====================================================================================
function doGet(e) {
  // ตรวจสอบและสร้างฐานข้อมูลหากยังไม่มี
  try {
    initializeSystem();
  } catch (err) {
    Logger.log('Init error in doGet: ' + err.toString());
  }

  const html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบแจ้งปัญหาการใช้โปรแกรมระบบบัญชีสหกรณ์')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    
  return html;
}

// =====================================================================================
// 3. DATABASE INITIALIZATION & HELPER FUNCTIONS
// =====================================================================================
/**
 * ดึง Spreadsheet Object
 */
function getSpreadsheet() {
  const propId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const targetId = propId || CONFIG.SPREADSHEET_ID;

  if (targetId && targetId.trim() !== '') {
    return SpreadsheetApp.openById(targetId.trim());
  }
  
  // กรณีผูกกับ Google Sheets อยู่แล้ว
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {
    // Not container-bound
  }

  throw new Error('กรุณาระบุ SPREADSHEET_ID ใน CONFIG หรือ Script Properties');
}

/**
 * สร้างชีตและ Headers อัตโนมัติหากยังไม่มี
 */
function initializeSystem() {
  const ss = getSpreadsheet();
  
  // 1. ตรวจสอบ/สร้าง Tickets
  let ticketSheet = ss.getSheetByName(SHEETS.TICKETS);
  if (!ticketSheet) {
    ticketSheet = ss.insertSheet(SHEETS.TICKETS);
    ticketSheet.appendRow(HEADERS.TICKETS);
    ticketSheet.getRange(1, 1, 1, HEADERS.TICKETS.length).setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');
    ticketSheet.setFrozenRows(1);
  }

  // 2. ตรวจสอบ/สร้าง WorkLogs
  let workLogSheet = ss.getSheetByName(SHEETS.WORKLOGS);
  if (!workLogSheet) {
    workLogSheet = ss.insertSheet(SHEETS.WORKLOGS);
    workLogSheet.appendRow(HEADERS.WORKLOGS);
    workLogSheet.getRange(1, 1, 1, HEADERS.WORKLOGS.length).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
    workLogSheet.setFrozenRows(1);
  }

  // 3. ตรวจสอบ/สร้าง Users
  let userSheet = ss.getSheetByName(SHEETS.USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(SHEETS.USERS);
    userSheet.appendRow(HEADERS.USERS);
    userSheet.getRange(1, 1, 1, HEADERS.USERS.length).setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
    userSheet.setFrozenRows(1);
    
    // ใส่เจ้าหน้าที่เริ่มต้น
    const defaultUsers = [
      ['U-SUPER', 'ผู้ดูแลระบบสูงสุด (Super Admin)', 'ศูนย์เทคโนโลยีสารสนเทศ', 'superadmin@cad.go.th', '02-628-5240', 'superadmin', 'Active', 'superadmin', '1234'],
      ['U001', 'ชัชชนก แสงงาม (เจ้าหน้าที่ระบบงานคอมพิวเตอร์)', 'กลุ่มพัฒนาระบบเทคโนโลยีสารสนเทศ', 'chatchanok@cad.go.th', '081-234-5678', 'admin', 'Active', 'admin', '1234']
    ];
    defaultUsers.forEach(u => userSheet.appendRow(u));
  }

  // 4. ตรวจสอบ/สร้าง Settings
  let settingSheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settingSheet) {
    settingSheet = ss.insertSheet(SHEETS.SETTINGS);
    settingSheet.appendRow(HEADERS.SETTINGS);
    settingSheet.getRange(1, 1, 1, HEADERS.SETTINGS.length).setFontWeight('bold').setBackground('#64748b').setFontColor('#ffffff');
    settingSheet.setFrozenRows(1);

    const defaultSettings = [
      ['ORG_NAME', 'กรมตรวจบัญชีสหกรณ์ / สำนักงานสหกรณ์'],
      ['SYSTEM_STATUS', 'Online'],
      ['TELEGRAM_ACTIVE', 'true'],
      ['GOOGLE_CHAT_ACTIVE', 'true']
    ];
    defaultSettings.forEach(s => settingSheet.appendRow(s));
  }

  // 5. ตรวจสอบ/สร้าง ชีตข้อมูลสหกรณ์ (ชื่อชีต: ชื่อสหกร์)
  let coopSheet = ss.getSheetByName('ชื่อสหกร์') || 
                  ss.getSheetByName('ชื่อสหกรณ์') || 
                  ss.getSheetByName('Cooperatives');
  
  if (!coopSheet) {
    // ตรวจสอบจากชื่อชีตที่มีอยู่ทั้งหมด
    const allSheets = ss.getSheets();
    for (let i = 0; i < allSheets.length; i++) {
      if (allSheets[i].getName().indexOf('สหก') > -1) {
        coopSheet = allSheets[i];
        break;
      }
    }
  }

  if (!coopSheet) {
    coopSheet = ss.insertSheet('ชื่อสหกรณ์');
    coopSheet.appendRow(HEADERS.COOPERATIVES);
    coopSheet.getRange(1, 1, 1, HEADERS.COOPERATIVES.length).setFontWeight('bold').setBackground('#0891b2').setFontColor('#ffffff');
    coopSheet.setFrozenRows(1);

    // รายชื่อสหกรณ์เริ่มต้น (23 รายการ ตามรหัสสหกรณ์จริงในจังหวัดสมุทรสงคราม)
    const defaultCoops = [
      ['5229', 'สหกรณ์การเกษตร', 'สหกรณ์การเกษตรเพื่อการตลาดลูกค้า ธกส.สส.จำกัด', 'นางสาววิภาดา ตรวจสอบดี', 'ภาครัฐ'],
      ['10064', 'สหกรณ์การเกษตร', 'สหกรณ์การเกษตรผสมผสานแพรกหนามแดง จำกัด', 'นางสาววิภาดา ตรวจสอบดี', 'ภาครัฐ'],
      ['5374', 'สหกรณ์การเกษตร', 'สหกรณ์การเกษตรไม้ผลสมุทรสงคราม จำกัด', 'นางสาววิภาดา ตรวจสอบดี', 'ภาครัฐ'],
      ['3282', 'สหกรณ์การเกษตร', 'สหกรณ์การเกษตรนาเกลือสมุทรสงคราม จำกัด', 'นายสุรชัย ประมงวิทย์', 'ภาครัฐ'],
      ['1412', 'สหกรณ์การเกษตร', 'สหกรณ์การเกษตรอัมพวา จำกัด', 'นายสุรชัย ประมงวิทย์', 'ภาครัฐ'],
      ['20007', 'สหกรณ์การเกษตร', 'สหกรณ์การเกษตรบางคนที จำกัด', 'นายสุรชัย ประมงวิทย์', 'ภาครัฐ'],
      ['19774', 'สหกรณ์การเกษตร', 'สหกรณ์การเกษตรนาเกลือบางแก้ว จำกัด', 'นายสุรชัย ประมงวิทย์', 'ภาครัฐ'],
      ['7615', 'สหกรณ์ประมง', 'สหกรณ์ประมงบางจะเกร็ง-บางแก้ว จำกัด', 'นายสุรชัย ประมงวิทย์', 'ภาครัฐ'],
      ['7616', 'สหกรณ์ประมง', 'สหกรณ์ประมงแม่กลอง จำกัด', 'นายสุรชัย ประมงวิทย์', 'ภาครัฐ'],
      ['15779', 'สหกรณ์ประมง', 'สหกรณ์ผู้เลี้ยงปลาสลิดแพรกหนามแดง จำกัด', 'นายสุรชัย ประมงวิทย์', 'ภาครัฐ'],
      ['18405', 'สหกรณ์ออมทรัพย์', 'สหกรณ์ออมทรัพย์ก้าวหน้า จำกัด', 'นายประสิทธิ์ ตรวจการ', 'ภาครัฐ'],
      ['3551', 'สหกรณ์ออมทรัพย์', 'สหกรณ์ออมทรัพย์โรงพยาบาลสมเด็จพระพุทธเลิศหล้า จำกัด', 'นายประสิทธิ์ ตรวจการ', 'ภาคเอกชน'],
      ['106', 'สหกรณ์ออมทรัพย์', 'สหกรณ์ออมทรัพย์ครูสมุทรสงคราม จำกัด', 'นายประสิทธิ์ ตรวจการ', 'ภาคเอกชน'],
      ['4364', 'สหกรณ์ออมทรัพย์', 'สหกรณ์ออมทรัพย์ตำรวจภูธรจังหวัดสมุทรสงคราม จำกัด', 'นายประสิทธิ์ ตรวจการ', 'ภาคเอกชน'],
      ['7979', 'สหกรณ์ออมทรัพย์', 'สหกรณ์ออมทรัพย์สาธารณสุขสมุทรสงคราม จำกัด', 'นายประสิทธิ์ ตรวจการ', 'ภาคเอกชน'],
      ['8233', 'สหกรณ์บริการ', 'สหกรณ์เดินรถแม่กลอง จำกัด', 'นางสาวจารุวรรณ กิจเจริญ', 'ภาครัฐ'],
      ['16764', 'สหกรณ์บริการ', 'สหกรณ์เคหสถานบ้านโรงกุ้งมั่นคง จำกัด', 'นางสาวจารุวรรณ กิจเจริญ', 'ภาครัฐ'],
      ['17847', 'สหกรณ์บริการ', 'สหกรณ์เคหสถานเจริญมั่งมีมั่นคง จำกัด', 'นางสาวจารุวรรณ กิจเจริญ', 'ภาครัฐ'],
      ['17768', 'สหกรณ์บริการ', 'สหกรณ์เคหสถานป้อมแก้วมั่นคง จำกัด', 'นางสาวจารุวรรณ กิจเจริญ', 'ภาครัฐ'],
      ['19118', 'สหกรณ์บริการ', 'สหกรณ์เคหสถานปู่ศรีราชายี่สาร จำกัด', 'นางสาวจารุวรรณ กิจเจริญ', 'ภาครัฐ'],
      ['19470', 'สหกรณ์บริการ', 'สหกรณ์เคหสถานสวนหลวงมั่นคงดำรงบุญ จำกัด', 'นางสาวจารุวรรณ กิจเจริญ', 'ภาครัฐ'],
      ['8072', 'สหกรณ์เครดิตยูเนี่ยน', 'สหกรณ์เครดิตยูเนี่ยนแม่กลอง จำกัด', 'นางสาวจารุวรรณ กิจเจริญ', 'ภาครัฐ'],
      ['6786', 'สหกรณ์เครดิตยูเนี่ยน', 'สหกรณ์เครดิตยูเนี่ยนบางนกแขวก จำกัด', 'นางสาวจารุวรรณ กิจเจริญ', 'ภาคเอกชน']
    ];
    defaultCoops.forEach(c => coopSheet.appendRow(c));
  }

  return { success: true, message: 'ระบบและฐานข้อมูล Google Sheets พร้อมใช้งาน' };
}

/**
 * ดึงรายการข้อมูลสหกรณ์ทั้งหมดจากชีต (อัปเดตตรงจาก Google Sheets แบบ Real-time)
 */
function getCooperatives() {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('ชื่อสหกรณ์') || 
                ss.getSheetByName('coop') || 
                ss.getSheetByName('Coop') || 
                ss.getSheetByName('ชื่อสหกร์') || 
                ss.getSheetByName('Cooperatives') ||
                ss.getSheetByName('รายชื่อสหกรณ์');

    if (!sheet) {
      const allSheets = ss.getSheets();
      for (let i = 0; i < allSheets.length; i++) {
        const sName = allSheets[i].getName();
        if (sName.indexOf('สหก') > -1 || sName.toLowerCase().indexOf('coop') > -1) {
          sheet = allSheets[i];
          break;
        }
      }
    }

    if (!sheet) {
      initializeSystem();
      sheet = ss.getSheetByName('ชื่อสหกรณ์');
    }

    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(1, sheet.getLastColumn());
    if (lastRow < 1) return [];

    const rawData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    if (!rawData || rawData.length === 0) return [];

    // วิเคราะห์ Header Row
    const headerRow = rawData[0].map(h => String(h || '').trim());
    let idCol = -1;
    let typeCol = -1;
    let nameCol = -1;
    let auditorNameCol = -1;
    let auditorTypeCol = -1;

    for (let c = 0; c < headerRow.length; c++) {
      const h = headerRow[c].toLowerCase();
      if (h.includes('รหัส') || h === 'id' || h === 'coopid') {
        idCol = c;
      } else if (h.includes('ผู้สอบ') && (h.includes('ชื่อ') || h.includes('นามสกุล'))) {
        auditorNameCol = c;
      } else if (h.includes('ผู้สอบ') && (h.includes('ประเภท') || h.includes('สังกัด') || h.includes('ภาค'))) {
        auditorTypeCol = c;
      } else if (h.includes('ชื่อ') && !h.includes('ผู้สอบ')) {
        nameCol = c;
      } else if (h.includes('ประเภท') && !h.includes('ผู้สอบ')) {
        typeCol = c;
      }
    }

    let startRow = 1;
    // หากไม่พบ Header ที่ตรง หรือแถวแรกเป็นข้อมูลสหกรณ์เลย
    if (nameCol === -1) {
      if (String(rawData[0][2] || '').includes('สหกรณ์')) {
        // ตรงกับโครงสร้าง: รหัส, ประเภท, ชื่อสหกรณ์, ผู้สอบบัญชี, ประเภทผู้สอบ
        idCol = 0;
        typeCol = 1;
        nameCol = 2;
        auditorNameCol = 3;
        auditorTypeCol = 4;
        startRow = 0;
      } else if (String(rawData[0][0] || '').includes('สหกรณ์') || String(rawData[0][1] || '').includes('สหกรณ์')) {
        nameCol = String(rawData[0][0] || '').includes('สหกรณ์') ? 0 : 1;
        startRow = 0;
      } else {
        if (lastCol >= 3) {
          idCol = 0;
          typeCol = 1;
          nameCol = 2;
          auditorNameCol = lastCol >= 4 ? 3 : -1;
          auditorTypeCol = lastCol >= 5 ? 4 : -1;
        } else if (lastCol >= 2) {
          idCol = 0;
          nameCol = 1;
        } else {
          nameCol = 0;
        }
      }
    }

    const coops = [];
    const seenNames = new Set();

    for (let r = startRow; r < rawData.length; r++) {
      const row = rawData[r];
      const coopName = String((nameCol > -1 ? row[nameCol] : row[0]) || '').trim();
      if (!coopName || coopName === 'ชื่อสหกรณ์' || coopName === 'ชื่อสหกร์' || coopName === 'CoopName') continue;
      if (seenNames.has(coopName)) continue;
      seenNames.add(coopName);

      let coopId = idCol > -1 ? String(row[idCol] || '').trim() : '';
      let coopType = typeCol > -1 ? String(row[typeCol] || '').trim() : '';
      let auditorName = auditorNameCol > -1 ? String(row[auditorNameCol] || '').trim() : '';
      let auditorType = auditorTypeCol > -1 ? String(row[auditorTypeCol] || '').trim() : '';

      // ตรวจสอบและแก้ไขประเภทสหกรณ์ให้ถูกต้องตามชื่อจริง
      if (!coopType || (coopType === 'สหกรณ์ประมง' && coopName.includes('ออมทรัพย์')) || (coopType.includes('เครดิตยูเนียน') && !coopType.includes('เครดิตยูเนี่ยน'))) {
        if (coopName.includes('การเกษตร')) coopType = 'สหกรณ์การเกษตร';
        else if (coopName.includes('ออมทรัพย์')) coopType = 'สหกรณ์ออมทรัพย์';
        else if (coopName.includes('เครดิตยูเนี่ยน') || coopName.includes('เครดิตยูเนียน')) coopType = 'สหกรณ์เครดิตยูเนี่ยน';
        else if (coopName.includes('บริการ') || coopName.includes('เคหสถาน') || coopName.includes('เดินรถ')) coopType = 'สหกรณ์บริการ';
        else if (coopName.includes('ร้านค้า')) coopType = 'สหกรณ์ร้านค้า';
        else if (coopName.includes('นิคม')) coopType = 'สหกรณ์นิคม';
        else if (coopName.includes('ประมง')) coopType = 'สหกรณ์ประมง';
        else coopType = 'สหกรณ์ทั่วไป';
      }

      if (!auditorType) {
        auditorType = 'ภาครัฐ';
      }

      coops.push({
        coopId: coopId,
        coopName: coopName,
        coopType: coopType,
        auditorName: auditorName,
        auditorType: auditorType
      });
    }

    return coops;
  } catch (err) {
    Logger.log('getCooperatives error: ' + err.toString());
    return [];
  }
}

/**
 * ดึงค่า Config (รวมถึง Script Properties)
 */
function getSystemConfig() {
  const props = PropertiesService.getScriptProperties().getProperties();
  return {
    SPREADSHEET_ID: props.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID,
    DRIVE_FOLDER_ID: props.DRIVE_FOLDER_ID || CONFIG.DRIVE_FOLDER_ID,
    TELEGRAM_BOT_TOKEN: props.TELEGRAM_BOT_TOKEN || CONFIG.TELEGRAM_BOT_TOKEN,
    TELEGRAM_CHAT_ID: props.TELEGRAM_CHAT_ID || CONFIG.TELEGRAM_CHAT_ID,
    GOOGLE_CHAT_WEBHOOK_URL: props.GOOGLE_CHAT_WEBHOOK_URL || CONFIG.GOOGLE_CHAT_WEBHOOK_URL,
    ADMIN_NOTIFICATION_EMAILS: props.ADMIN_NOTIFICATION_EMAILS || CONFIG.ADMIN_NOTIFICATION_EMAILS,
    ENABLE_EMAIL_NOTIFICATION: props.ENABLE_EMAIL_NOTIFICATION !== undefined ? (props.ENABLE_EMAIL_NOTIFICATION === 'true' || props.ENABLE_EMAIL_NOTIFICATION === true) : CONFIG.ENABLE_EMAIL_NOTIFICATION,
    WEB_APP_URL: props.WEB_APP_URL || CONFIG.WEB_APP_URL,
    TIMEZONE: CONFIG.TIMEZONE
  };
}

/**
 * บันทึกค่า Config ลง Script Properties
 */
function saveSystemConfig(newConfig) {
  const props = PropertiesService.getScriptProperties();
  if (newConfig.SPREADSHEET_ID !== undefined) props.setProperty('SPREADSHEET_ID', newConfig.SPREADSHEET_ID);
  if (newConfig.DRIVE_FOLDER_ID !== undefined) props.setProperty('DRIVE_FOLDER_ID', newConfig.DRIVE_FOLDER_ID);
  if (newConfig.TELEGRAM_BOT_TOKEN !== undefined) props.setProperty('TELEGRAM_BOT_TOKEN', newConfig.TELEGRAM_BOT_TOKEN);
  if (newConfig.TELEGRAM_CHAT_ID !== undefined) props.setProperty('TELEGRAM_CHAT_ID', newConfig.TELEGRAM_CHAT_ID);
  if (newConfig.GOOGLE_CHAT_WEBHOOK_URL !== undefined) props.setProperty('GOOGLE_CHAT_WEBHOOK_URL', newConfig.GOOGLE_CHAT_WEBHOOK_URL);
  if (newConfig.ADMIN_NOTIFICATION_EMAILS !== undefined) props.setProperty('ADMIN_NOTIFICATION_EMAILS', newConfig.ADMIN_NOTIFICATION_EMAILS);
  if (newConfig.ENABLE_EMAIL_NOTIFICATION !== undefined) props.setProperty('ENABLE_EMAIL_NOTIFICATION', String(newConfig.ENABLE_EMAIL_NOTIFICATION));
  if (newConfig.WEB_APP_URL !== undefined) props.setProperty('WEB_APP_URL', newConfig.WEB_APP_URL);

  return { success: true, message: 'บันทึกการตั้งค่าเรียบร้อยแล้ว' };
}

// =====================================================================================
// 4. TICKET GENERATION & MANAGEMENT
// =====================================================================================
/**
 * สร้าง Ticket ID อัตโนมัติ รูปแบบ noYYYYMMXXXXXX เช่น no202609000002
 */
function generateTicketId() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const now = new Date();
  const yearMonth = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyyMM');
  const prefix = 'no' + yearMonth;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return prefix + '000001';
  }

  // อ่านเฉพาะคอลัมน์ TicketID (คอลัมน์ 1)
  const ticketIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxSeq = 0;

  for (let i = 0; i < ticketIds.length; i++) {
    const id = String(ticketIds[i]).trim();
    if (id.startsWith(prefix)) {
      const seqStr = id.substring(prefix.length);
      const seq = parseInt(seqStr, 10);
      if (!isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
    } else {
      // ตรวจสอบเลขลำดับจากเคสก่อนหน้าในระบบ เพื่อให้รันเลขต่อจากเดิมโดยอัตโนมัติ
      const match = id.match(/no(?:\d{6})?(\d{6})$/i) || id.match(/(\d{6})$/);
      if (match) {
        const seq = parseInt(match[1], 10);
        if (!isNaN(seq) && seq > maxSeq) {
          maxSeq = seq;
        }
      }
    }
  }

  const nextSeq = String(maxSeq + 1).padStart(6, '0');
  return prefix + nextSeq;
}

/**
 * ดึงเลขที่ใบงานถัดไปเพื่อนำไปแสดงผลบนฟอร์ม
 */
function getNextTicketId() {
  try {
    return generateTicketId();
  } catch (e) {
    const now = new Date();
    const ym = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyyMM');
    return 'no' + ym + '000001';
  }
}

/**
 * แปลงข้อมูลแถวใน Sheets เป็น Object Ticket
 */
function rowToTicket(row) {
  return {
    ticketId: row[0] || '',
    date: row[1] || '',
    time: row[2] || '',
    createdAt: row[3] || (row[1] ? (row[1] + ' ' + (row[2] || '')) : ''),
    reporterName: row[4] || '',
    position: row[5] || '',
    phone: row[6] || '',
    email: row[7] || '',
    coopName: row[8] || '',
    coopId: row[9] || '',
    coopType: row[10] || '',
    auditorName: row[11] || '',
    auditorType: row[12] || '',
    softwareList: row[13] || '',
    problemTitle: row[14] || '',
    problemDetail: row[15] || '',
    priority: row[16] || '🟢 ปกติ',
    status: row[17] || '🟡 รอรับเรื่อง',
    assignedTo: row[18] || '',
    assignedAt: row[19] || '',
    acceptedAt: row[20] || '',
    startedAt: row[21] || '',
    resolvedAt: row[22] || '',
    closedAt: row[23] || '',
    resolution: row[24] || '',
    rootCause: row[25] || '',
    cost: row[26] || 0,
    rating: row[27] || 0,
    feedback: row[28] || '',
    attachmentUrl: row[29] || '',
    lastUpdated: row[30] || ''
  };
}

/**
 * จัดการอัปโหลดไฟล์แนบไปยัง Google Drive (สูงสุด 5 ไฟล์)
 * อัปเดตการตั้งชื่อไฟล์ให้มี Ticket ID และชื่อสหกรณ์ เพื่อความเป็นระเบียบและค้นหาง่าย
 */
function uploadMultipleAttachmentsToDrive(attachments, ticketId, coopName) {
  if (!attachments) return '';
  const fileList = Array.isArray(attachments) ? attachments : [attachments];
  if (fileList.length === 0) return '';

  const cfg = getSystemConfig();
  const folderId = cfg.DRIVE_FOLDER_ID;
  let folder;

  try {
    if (folderId && folderId.trim() !== '') {
      folder = DriveApp.getFolderById(folderId.trim());
    } else {
      const folders = DriveApp.getFoldersByName('IT_Helpdesk_Files');
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder('IT_Helpdesk_Files');
        try {
          folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (e) {}
      }
    }
  } catch (e) {
    folder = DriveApp.getRootFolder();
  }

  const urls = [];
  const maxFiles = Math.min(fileList.length, 5);

  for (let i = 0; i < maxFiles; i++) {
    const fileObj = fileList[i];
    if (!fileObj || !fileObj.base64) continue;

    try {
      const base64Data = fileObj.base64.indexOf(',') > -1 
        ? fileObj.base64.split(',')[1] 
        : fileObj.base64;
      const data = Utilities.base64Decode(base64Data);
      const timeStamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
      
      // อัปเดตรูปแบบชื่อไฟล์ให้ระบุ Ticket ID, ชื่อสหกรณ์ และชื่อไฟล์เดิม
      let prefix = '';
      if (ticketId) prefix += ticketId + '_';
      if (coopName) {
        const cleanCoop = coopName.replace(/[^a-zA-Z0-9ก-๙]/g, '_').substring(0, 20);
        if (cleanCoop) prefix += cleanCoop + '_';
      }
      
      const rawName = fileObj.name || ('file_' + timeStamp + '_' + (i + 1));
      const cleanOriginalName = rawName.replace(/[^a-zA-Z0-9ก-๙._-]/g, '_');
      const safeName = prefix ? (prefix + cleanOriginalName) : cleanOriginalName;
      const mimeType = fileObj.mimeType || 'application/octet-stream';
      const blob = Utilities.newBlob(data, mimeType, safeName);
      const createdFile = folder.createFile(blob);

      try {
        createdFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (err) {}

      urls.push(createdFile.getUrl());
    } catch (uploadErr) {
      Logger.log('File upload err at index ' + i + ': ' + uploadErr.toString());
    }
  }

  return urls.join('\n');
}

function uploadAttachmentToDrive(fileObj, ticketId, coopName) {
  return uploadMultipleAttachmentsToDrive([fileObj], ticketId, coopName);
}

/**
 * สร้าง Ticket ใหม่
 */
function createTicket(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // รอ Lock ได้สูงสุด 20 วินาที
  } catch (e) {
    throw new Error('ระบบกำลังประมวลผลคำขอจำนวนมาก กรุณาลองใหม่อีกครั้ง');
  }

  try {
    initializeSystem();
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.TICKETS);
    const now = new Date();
    const ticketDate = data.date || Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy');
    const ticketTime = data.time || Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm:ss');
    const nowFormatted = ticketDate + ' ' + ticketTime;
    const ticketId = generateTicketId();

    // จัดการไฟล์แนบ (สูงสุด 5 ไฟล์ บน Google Drive)
    let attachmentUrl = '';
    if (data.attachments && Array.isArray(data.attachments) && data.attachments.length > 0) {
      attachmentUrl = uploadMultipleAttachmentsToDrive(data.attachments, ticketId, data.coopName);
    } else if (data.attachment && data.attachment.base64) {
      attachmentUrl = uploadMultipleAttachmentsToDrive([data.attachment], ticketId, data.coopName);
    }

    const priority = data.priority || '🟢 ปกติ';
    const status = '🟡 รอรับเรื่อง';

    const row = [
      ticketId,
      ticketDate,
      ticketTime,
      nowFormatted,
      data.reporterName || '',
      data.position || '',
      data.phone || '',
      data.email || '',
      data.coopName || '',
      data.coopId || '',
      data.coopType || '',
      data.auditorName || '',
      data.auditorType || '',
      data.softwareList || '',
      data.problemTitle || '',
      data.problemDetail || '',
      priority,
      status,
      '', // AssignedTo
      '', // AssignedAt
      '', // AcceptedAt
      '', // StartedAt
      '', // ResolvedAt
      '', // ClosedAt
      '', // Resolution
      '', // RootCause
      0,  // Cost
      0,  // Rating
      '', // Feedback
      attachmentUrl,
      nowFormatted // LastUpdated
    ];

    sheet.appendRow(row);

    // บันทึก WorkLog แรก (แจ้งปัญหา)
    addWorkLogInternal({
      ticketId: ticketId,
      staff: data.reporterName || 'ผู้แจ้ง',
      action: 'สร้าง Ticket แจ้งปัญหาเข้าระบบ',
      parts: '-',
      cost: 0,
      note: 'ปัญหา: ' + (data.problemTitle || '') + ' (' + priority + ')' + (data.coopName ? (' | สหกรณ์: ' + data.coopName) : '')
    });

    const ticketObj = rowToTicket(row);

    // ส่งการแจ้งเตือน
    try {
      sendTelegramNotification(ticketObj, 'NEW_TICKET');
    } catch (e) {
      Logger.log('Telegram err: ' + e.toString());
    }

    try {
      sendGoogleChatNotification(ticketObj, 'NEW_TICKET');
    } catch (e) {
      Logger.log('Google Chat err: ' + e.toString());
    }

    // 1. ส่งอีเมลแจ้งเตือน Admin และ Super Admin (แสดงระดับความสำคัญ และชื่อสหกรณ์)
    try {
      sendAdminNewTicketEmail(ticketObj);
    } catch (e) {
      Logger.log('Admin Email err: ' + e.toString());
    }

    // 2. ส่งอีเมลยืนยันการรับเรื่องไปยังเจ้าหน้าที่สหกรณ์ (อีเมลที่กรอกมาตอนแรก)
    try {
      sendReporterStatusUpdateEmail(ticketObj, {
        eventType: 'CREATED',
        newStatus: ticketObj.status,
        note: 'ระบบได้รับเรื่องแจ้งปัญหาของท่านเรียบร้อยแล้ว ขณะนี้กำลังรอเจ้าหน้าที่ไอทีตรวจสอบเพื่อดำเนินการ'
      });
    } catch (e) {
      Logger.log('Reporter Email err: ' + e.toString());
    }

    return {
      success: true,
      ticketId: ticketId,
      date: ticketDate,
      time: ticketTime,
      createdAt: nowFormatted,
      reporterName: data.reporterName,
      coopName: data.coopName,
      problemTitle: data.problemTitle,
      priority: priority,
      status: status,
      attachmentUrl: attachmentUrl
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * ดึงรายการ Tickets พร้อมระบบค้นหา, Filter, และ Pagination
 */
function getTickets(params) {
  initializeSystem();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return { tickets: [], totalCount: 0, page: 1, totalPages: 0 };
  }

  // อ่านข้อมูลทั้งหมดแบบ Batch
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.TICKETS.length).getValues();
  let tickets = data.map(rowToTicket);

  params = params || {};
  const query = (params.search || '').trim().toLowerCase();
  const statusFilter = (params.status || '').trim();
  const problemTypeFilter = (params.problemType || '').trim();
  const priorityFilter = (params.priority || '').trim();
  const buildingFilter = (params.building || '').trim();
  const reporterFilter = (params.reporter || '').trim();
  const dateFilter = (params.date || '').trim();

  // กรองข้อมูล
  if (query) {
    tickets = tickets.filter(t => 
      (t.ticketId && t.ticketId.toLowerCase().includes(query)) ||
      (t.reporterName && t.reporterName.toLowerCase().includes(query)) ||
      (t.coopName && t.coopName.toLowerCase().includes(query)) ||
      (t.coopId && t.coopId.toLowerCase().includes(query)) ||
      (t.auditorName && t.auditorName.toLowerCase().includes(query)) ||
      (t.softwareList && t.softwareList.toLowerCase().includes(query)) ||
      (t.problemTitle && t.problemTitle.toLowerCase().includes(query)) ||
      (t.problemDetail && t.problemDetail.toLowerCase().includes(query)) ||
      (t.assignedTo && t.assignedTo.toLowerCase().includes(query))
    );
  }

  if (statusFilter && statusFilter !== 'ทั้งหมด') {
    tickets = tickets.filter(t => t.status.includes(statusFilter));
  }

  if (problemTypeFilter && problemTypeFilter !== 'ทั้งหมด') {
    tickets = tickets.filter(t => (t.coopType && t.coopType === problemTypeFilter) || (t.softwareList && t.softwareList.includes(problemTypeFilter)));
  }

  if (priorityFilter && priorityFilter !== 'ทั้งหมด') {
    tickets = tickets.filter(t => t.priority.includes(priorityFilter));
  }

  if (buildingFilter && buildingFilter !== 'ทั้งหมด' && buildingFilter !== 'ประเภทสหกรณ์ทั้งหมด') {
    tickets = tickets.filter(t => (t.coopName && t.coopName.includes(buildingFilter)) || (t.coopType && t.coopType.includes(buildingFilter)));
  }

  if (reporterFilter) {
    tickets = tickets.filter(t => (t.reporterName && t.reporterName.toLowerCase().includes(reporterFilter.toLowerCase())) || (t.email && t.email.toLowerCase().includes(reporterFilter.toLowerCase())));
  }

  if (dateFilter) {
    tickets = tickets.filter(t => t.createdAt.startsWith(dateFilter));
  }

  // เรียงลำดับ: ให้ความเร่งด่วน 'ด่วนมาก' อยู่บนสุด แล้วตามด้วยวันที่สร้างล่าสุด
  tickets.sort((a, b) => {
    const isUrgentA = a.priority.includes('ด่วนมาก');
    const isUrgentB = b.priority.includes('ด่วนมาก');
    if (isUrgentA && !isUrgentB) return -1;
    if (!isUrgentA && isUrgentB) return 1;
    return b.ticketId.localeCompare(a.ticketId);
  });

  const totalCount = tickets.length;
  const page = Math.max(1, parseInt(params.page || 1, 10));
  const pageSize = Math.max(5, parseInt(params.pageSize || 10, 10));
  const totalPages = Math.ceil(totalCount / pageSize);
  const startIndex = (page - 1) * pageSize;
  const paginatedTickets = tickets.slice(startIndex, startIndex + pageSize);

  return {
    tickets: paginatedTickets,
    totalCount: totalCount,
    page: page,
    pageSize: pageSize,
    totalPages: totalPages
  };
}

/**
 * ดึงข้อมูล Ticket ตาม ID พร้อม WorkLogs
 */
function getTicketById(ticketId) {
  if (!ticketId) throw new Error('ไม่พบ Ticket ID');
  initializeSystem();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return null;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.TICKETS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(ticketId).trim()) {
      const ticket = rowToTicket(data[i]);
      // ดึง WorkLogs
      ticket.workLogs = getWorkLogsForTicket(ticketId);
      return ticket;
    }
  }

  return null;
}

/**
 * ค้นหาแถวของ Ticket ใน Sheet (1-based index)
 */
function findTicketRowIndex(ticketId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i]).trim() === String(ticketId).trim()) {
      return i + 2; // Row number in sheet
    }
  }
  return -1;
}

/**
 * มอบหมายงาน (Assign Ticket)
 */
function assignTicket(ticketId, staff, note) {
  const rowIndex = findTicketRowIndex(ticketId);
  if (rowIndex === -1) throw new Error('ไม่พบ Ticket: ' + ticketId);

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const now = new Date();
  const nowFormatted = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');

  // Headers: Status (Col 18), AssignedTo (Col 19), AssignedAt (Col 20), AcceptedAt (Col 21), LastUpdated (Col 31)
  const currentStatus = sheet.getRange(rowIndex, 18).getValue();
  let newStatus = currentStatus;
  if (currentStatus.includes('รอรับเรื่อง')) {
    newStatus = '🔵 รับเรื่องแล้ว';
    sheet.getRange(rowIndex, 18).setValue(newStatus);
    sheet.getRange(rowIndex, 21).setValue(nowFormatted); // AcceptedAt
  }

  sheet.getRange(rowIndex, 19).setValue(staff);
  sheet.getRange(rowIndex, 20).setValue(nowFormatted);
  sheet.getRange(rowIndex, 31).setValue(nowFormatted);

  // บันทึก WorkLog
  addWorkLogInternal({
    ticketId: ticketId,
    staff: staff,
    action: 'มอบหมายงานให้: ' + staff,
    parts: '-',
    cost: 0,
    note: note || 'มอบหมายงานและรับเรื่องเข้าระบบ'
  });

  const updatedTicket = getTicketById(ticketId);

  // แจ้งเตือน Telegram & Google Chat
  try {
    sendTelegramNotification(updatedTicket, 'ASSIGN', { staff: staff, note: note });
  } catch (e) { Logger.log(e); }

  try {
    sendGoogleChatNotification(updatedTicket, 'ASSIGN', { staff: staff, note: note });
  } catch (e) { Logger.log(e); }

  // ส่งอีเมลแจ้งเตือนการมอบหมายงานไปยังอีเมลผู้แจ้ง
  try {
    sendReporterStatusUpdateEmail(updatedTicket, {
      eventType: 'ASSIGN',
      newStatus: updatedTicket.status,
      note: 'มอบหมายงานให้เจ้าหน้าที่ผู้รับผิดชอบ: ' + staff + (note ? ' (' + note + ')' : ''),
      staff: staff
    });
  } catch (e) { Logger.log('Reporter Email err: ' + e.toString()); }

  return { success: true, message: 'มอบหมายงานสำเร็จ', ticket: updatedTicket };
}

/**
 * อัปเดตสถานะ Ticket
 */
function updateTicketStatus(ticketId, newStatus, note, staffName) {
  const rowIndex = findTicketRowIndex(ticketId);
  if (rowIndex === -1) throw new Error('ไม่พบ Ticket: ' + ticketId);

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const now = new Date();
  const nowFormatted = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  const oldStatus = sheet.getRange(rowIndex, 18).getValue();

  sheet.getRange(rowIndex, 18).setValue(newStatus);
  sheet.getRange(rowIndex, 31).setValue(nowFormatted);

  // บันทึก Timestamps ตามสถานะ
  if (newStatus.includes('รับเรื่อง') && !sheet.getRange(rowIndex, 21).getValue()) {
    sheet.getRange(rowIndex, 21).setValue(nowFormatted); // AcceptedAt
  }
  if (newStatus.includes('กำลังดำเนินการ') && !sheet.getRange(rowIndex, 22).getValue()) {
    sheet.getRange(rowIndex, 22).setValue(nowFormatted); // StartedAt
  }
  if (newStatus.includes('แก้ไขแล้ว') && !sheet.getRange(rowIndex, 23).getValue()) {
    sheet.getRange(rowIndex, 23).setValue(nowFormatted); // ResolvedAt
  }
  if (newStatus.includes('ปิดงาน') && !sheet.getRange(rowIndex, 24).getValue()) {
    sheet.getRange(rowIndex, 24).setValue(nowFormatted); // ClosedAt
  }

  // บันทึก WorkLog
  addWorkLogInternal({
    ticketId: ticketId,
    staff: staffName || 'เจ้าหน้าที่ IT',
    action: 'เปลี่ยนสถานะเป็น: ' + newStatus,
    parts: '-',
    cost: 0,
    note: note || ('สถานะเปลี่ยนจาก ' + oldStatus + ' เป็น ' + newStatus)
  });

  const updatedTicket = getTicketById(ticketId);

  // ส่งแจ้งเตือน
  try {
    sendTelegramNotification(updatedTicket, 'STATUS_UPDATE', { oldStatus: oldStatus, newStatus: newStatus, note: note, staff: staffName });
  } catch (e) { Logger.log(e); }

  try {
    sendGoogleChatNotification(updatedTicket, 'STATUS_UPDATE', { oldStatus: oldStatus, newStatus: newStatus, note: note, staff: staffName });
  } catch (e) { Logger.log(e); }

  // ส่งอีเมลอัปเดตสถานะและขั้นตอนไปยังอีเมลผู้แจ้ง
  try {
    sendReporterStatusUpdateEmail(updatedTicket, {
      eventType: 'STATUS_UPDATE',
      oldStatus: oldStatus,
      newStatus: newStatus,
      note: note,
      staff: staffName
    });
  } catch (e) { Logger.log('Reporter Email err: ' + e.toString()); }

  return { success: true, message: 'อัปเดตสถานะสำเร็จ', ticket: updatedTicket };
}

/**
 * ปิดงาน Ticket (พร้อมบันทึกสาเหตุ วิธีแก้ไข และผลงาน)
 */
function closeTicket(ticketId, data) {
  const rowIndex = findTicketRowIndex(ticketId);
  if (rowIndex === -1) throw new Error('ไม่พบ Ticket: ' + ticketId);

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const now = new Date();
  const nowFormatted = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');

  // 18: Status, 23: ResolvedAt, 24: ClosedAt, 25: Resolution, 26: RootCause, 27: Cost, 31: LastUpdated
  sheet.getRange(rowIndex, 18).setValue('⚫ ปิดงาน');
  if (!sheet.getRange(rowIndex, 23).getValue()) sheet.getRange(rowIndex, 23).setValue(nowFormatted);
  sheet.getRange(rowIndex, 24).setValue(nowFormatted);
  sheet.getRange(rowIndex, 25).setValue(data.resolution || '');
  sheet.getRange(rowIndex, 26).setValue(data.rootCause || '');
  sheet.getRange(rowIndex, 27).setValue(data.cost || 0);
  sheet.getRange(rowIndex, 31).setValue(nowFormatted);

  // บันทึก WorkLog ปิดงาน
  addWorkLogInternal({
    ticketId: ticketId,
    staff: data.staff || 'IT Admin',
    action: 'ปิดงาน Ticket สมบูรณ์',
    parts: data.parts || '-',
    cost: data.cost || 0,
    note: 'สาเหตุ: ' + (data.rootCause || '-') + ' | วิธีแก้ไข: ' + (data.resolution || '-') + ' | ' + (data.note || '')
  });

  const updatedTicket = getTicketById(ticketId);

  try {
    sendTelegramNotification(updatedTicket, 'RESOLVED', { resolution: data.resolution, staff: data.staff });
  } catch (e) { Logger.log(e); }

  try {
    sendGoogleChatNotification(updatedTicket, 'RESOLVED', { resolution: data.resolution, staff: data.staff });
  } catch (e) { Logger.log(e); }

  // ส่งอีเมลแจ้งเตือนปิดงานและผลการแก้ไขไปยังอีเมลผู้แจ้ง
  try {
    sendReporterStatusUpdateEmail(updatedTicket, {
      eventType: 'CLOSED',
      newStatus: '⚫ ปิดงาน',
      resolution: data.resolution,
      rootCause: data.rootCause,
      note: 'ดำเนินการแก้ไขและปิดงานสมบูรณ์แล้ว วิธีแก้ไข: ' + (data.resolution || '-'),
      staff: data.staff
    });
  } catch (e) { Logger.log('Reporter Email err: ' + e.toString()); }

  return { success: true, message: 'ปิดงานเรียบร้อยแล้ว', ticket: updatedTicket };
}

/**
 * ยกเลิกเคส Ticket (พร้อมบันทึกเหตุผลที่ยกเลิก)
 */
function cancelTicket(ticketId, data) {
  const rowIndex = findTicketRowIndex(ticketId);
  if (rowIndex === -1) throw new Error('ไม่พบ Ticket: ' + ticketId);

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const now = new Date();
  const nowFormatted = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  const oldStatus = sheet.getRange(rowIndex, 18).getValue();
  const reason = (data && data.reason) ? data.reason.trim() : 'ยกเลิกโดยเจ้าหน้าที่ IT';
  const staff = (data && data.staff) ? data.staff.trim() : 'IT Admin';
  const note = (data && data.note) ? data.note.trim() : '';

  // 18: Status, 24: ClosedAt, 25: Resolution, 26: RootCause, 31: LastUpdated
  sheet.getRange(rowIndex, 18).setValue('🔴 ยกเลิก');
  sheet.getRange(rowIndex, 24).setValue(nowFormatted);
  sheet.getRange(rowIndex, 25).setValue('เหตุผลการยกเลิก: ' + reason);
  sheet.getRange(rowIndex, 26).setValue('ยกเลิกคำขอ');
  sheet.getRange(rowIndex, 31).setValue(nowFormatted);

  // บันทึก WorkLog ปฏิบัติงาน
  addWorkLogInternal({
    ticketId: ticketId,
    staff: staff,
    action: 'ยกเลิกเคส Ticket',
    parts: '-',
    cost: 0,
    note: 'เหตุผลที่ยกเลิก: ' + reason + (note ? ' (หมายเหตุ: ' + note + ')' : '')
  });

  const updatedTicket = getTicketById(ticketId);

  // ส่งแจ้งเตือน Telegram
  try {
    sendTelegramNotification(updatedTicket, 'STATUS_UPDATE', {
      oldStatus: oldStatus,
      newStatus: '🔴 ยกเลิก',
      note: 'เหตุผลการยกเลิก: ' + reason,
      staff: staff
    });
  } catch (e) { Logger.log(e); }

  // ส่งแจ้งเตือน Google Chat
  try {
    sendGoogleChatNotification(updatedTicket, 'STATUS_UPDATE', {
      oldStatus: oldStatus,
      newStatus: '🔴 ยกเลิก',
      note: 'เหตุผลการยกเลิก: ' + reason,
      staff: staff
    });
  } catch (e) { Logger.log(e); }

  // ส่งอีเมลแจ้งผู้แจ้งกรณีมีการระบุอีเมลไว้
  try {
    sendReporterStatusUpdateEmail(updatedTicket, {
      eventType: 'CANCELLED',
      oldStatus: oldStatus,
      newStatus: '🔴 ยกเลิก',
      note: 'เคสถูกยกเลิกเนื่องจาก: ' + reason,
      staff: staff
    });
  } catch (e) { Logger.log('Reporter Email err: ' + e.toString()); }

  return { success: true, message: 'ยกเลิกเคสเรียบร้อยแล้ว', ticket: updatedTicket };
}

/**
 * ลบเคสงานทั้งหมดทุกสถานะในระบบ (ล้างฐานข้อมูล Tickets และ WorkLogs ให้ว่างเปล่า)
 */
function clearAllTickets() {
  initializeSystem();
  const ss = getSpreadsheet();
  
  // 1. ล้างตาราง Tickets ทั้งหมด (คงแถวหัวตารางไว้ที่แถวที่ 1)
  const ticketSheet = ss.getSheetByName(SHEETS.TICKETS);
  if (ticketSheet) {
    const lastRow = ticketSheet.getLastRow();
    if (lastRow > 1) {
      ticketSheet.deleteRows(2, lastRow - 1);
    }
  }
  
  // 2. ล้างตาราง WorkLogs ทั้งหมด (คงแถวหัวตารางไว้ที่แถวที่ 1)
  const workLogSheet = ss.getSheetByName(SHEETS.WORKLOGS);
  if (workLogSheet) {
    const lastRow = workLogSheet.getLastRow();
    if (lastRow > 1) {
      workLogSheet.deleteRows(2, lastRow - 1);
    }
  }
  
  return {
    success: true,
    message: 'ลบเคสงานทั้งหมดในระบบเรียบร้อยแล้ว ฐานข้อมูลว่างเปล่า (0 รายการ)'
  };
}

/**
 * ผู้แจ้งส่งแบบประเมินความพึงพอใจ
 */
function submitRating(ticketId, rating, feedback) {
  const rowIndex = findTicketRowIndex(ticketId);
  if (rowIndex === -1) throw new Error('ไม่พบ Ticket: ' + ticketId);

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const now = new Date();
  const nowFormatted = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');

  // Col 28: Rating, Col 29: Feedback, Col 31: LastUpdated
  sheet.getRange(rowIndex, 28).setValue(rating);
  sheet.getRange(rowIndex, 29).setValue(feedback || '');
  sheet.getRange(rowIndex, 31).setValue(nowFormatted);

  return { success: true, message: 'ขอบคุณสำหรับการประเมินความพึงพอใจ' };
}

// =====================================================================================
// 5. WORK LOGS
// =====================================================================================
function addWorkLogInternal(data) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.WORKLOGS);
  const now = new Date();
  const nowFormatted = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  const logId = 'LOG-' + Date.now();

  const noteWithMethod = data.method 
    ? ('[วิธีดำเนินการ: ' + data.method + '] ' + (data.note || '')) 
    : (data.note || '');

  sheet.appendRow([
    logId,
    data.ticketId,
    nowFormatted,
    data.staff || 'IT Staff',
    data.action || '',
    data.parts || '-',
    data.cost || 0,
    noteWithMethod
  ]);

  return { success: true, logId: logId };
}

function addWorkLog(data) {
  if (!data.ticketId) throw new Error('ต้องระบุ Ticket ID');
  const res = addWorkLogInternal(data);

  // อัปเดต LastUpdated ของ Ticket (Col 31)
  const rowIndex = findTicketRowIndex(data.ticketId);
  if (rowIndex !== -1) {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.TICKETS);
    const nowFormatted = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    sheet.getRange(rowIndex, 31).setValue(nowFormatted);
  }

  return res;
}

function getWorkLogsForTicket(ticketId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.WORKLOGS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.WORKLOGS.length).getValues();
  const logs = [];

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(ticketId).trim()) {
      let rawNote = String(data[i][7] || '');
      let method = '';
      const methodMatch = rawNote.match(/^\[วิธีดำเนินการ:\s*([^\]]+)\]\s*(.*)$/);
      if (methodMatch) {
        method = methodMatch[1];
        rawNote = methodMatch[2];
      }

      logs.push({
        logId: data[i][0],
        ticketId: data[i][1],
        dateTime: data[i][2],
        staff: data[i][3],
        action: data[i][4],
        method: method,
        parts: data[i][5],
        cost: data[i][6],
        note: rawNote
      });
    }
  }

  // เรียงตามเวลาเก่าไปใหม่ เพื่อทำ Timeline
  return logs;
}

// =====================================================================================
// 6. DASHBOARD & ANALYTICS
// =====================================================================================
function getDashboardData() {
  initializeSystem();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const lastRow = sheet.getLastRow();

  const result = {
    total: 0,
    pending: 0,      // รอรับเรื่อง
    accepted: 0,     // รับเรื่องแล้ว
    inProgress: 0,   // กำลังดำเนินการ
    waitingParts: 0, // รออะไหล่/ข้อมูล
    resolved: 0,     // แก้ไขแล้ว
    closed: 0,       // ปิดงาน
    cancelled: 0,    // ยกเลิก
    urgent: 0,       // ด่วนมาก
    todayCount: 0,
    monthCount: 0,
    avgResolutionTimeHours: 0,
    byType: {},             // ประเภทสหกรณ์ (เข้ากันได้กับเดิม)
    byCoopType: {},         // จำนวนปัญหาแยกตามประเภทสหกรณ์
    bySoftware: {},         // จำนวนปัญหาแยกตามโปรแกรม (ชื่อหลัก)
    byVersion: {},          // จำนวนปัญหาแยกตามเวอร์ชั่น (v1.9, v2.0, ทั่วไป)
    bySoftwareVersion: {},  // จำนวนปัญหาแยกตามโปรแกรมพร้อมเวอร์ชั่น
    softwareRankings: [],   // รายการจัดอันดับโปรแกรมและเวอร์ชั่นจากมากไปน้อย
    coopTypeRankings: [],   // รายการจัดอันดับประเภทสหกรณ์จากมากไปน้อย
    topCoopType: '-',
    topSoftware: '-',
    topVersion: '-',
    byBuilding: {},
    byDepartment: {},
    byStaff: {},
    recentUrgentTickets: []
  };

  if (lastRow <= 1) return result;

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.TICKETS.length).getValues();
  const now = new Date();
  const todayStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const thisMonthStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'MM/yyyy');

  let totalResolutionMinutes = 0;
  let resolvedCountWithTime = 0;

  // สำหรับเก็บสถานะย่อยของโปรแกรม
  const softwareStatusMap = {};

  for (let i = 0; i < data.length; i++) {
    const t = rowToTicket(data[i]);
    result.total++;

    // นับสถานะ
    if (t.status.includes('รอรับเรื่อง')) result.pending++;
    else if (t.status.includes('รับเรื่องแล้ว')) result.accepted++;
    else if (t.status.includes('กำลังดำเนินการ')) result.inProgress++;
    else if (t.status.includes('รออะไหล่')) result.waitingParts++;
    else if (t.status.includes('แก้ไขแล้ว')) result.resolved++;
    else if (t.status.includes('ปิดงาน')) result.closed++;
    else if (t.status.includes('ยกเลิก')) result.cancelled++;

    // ความเร่งด่วน
    if (t.priority.includes('ด่วนมาก')) {
      result.urgent++;
      if (!t.status.includes('ปิดงาน') && !t.status.includes('แก้ไขแล้ว') && !t.status.includes('ยกเลิก')) {
        result.recentUrgentTickets.push(t);
      }
    }

    // วันนี้ & เดือนนี้
    if (t.createdAt.includes(todayStr)) result.todayCount++;
    if (t.createdAt.includes(thisMonthStr)) result.monthCount++;

    // 1. จำแนกตามประเภทสหกรณ์ (จำนวนปัญหาแยกตามประเภทสหกรณ์)
    const coopType = t.coopType ? t.coopType.trim() : 'ไม่ระบุประเภท';
    result.byType[coopType] = (result.byType[coopType] || 0) + 1;
    result.byCoopType[coopType] = (result.byCoopType[coopType] || 0) + 1;

    // 2. จำแนกตามโปรแกรมและเวอร์ชั่น
    const softwareRaw = t.softwareList || '';
    let softwareItems = [];
    if (softwareRaw && softwareRaw.trim()) {
      softwareItems = softwareRaw.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    }
    if (softwareItems.length === 0) {
      softwareItems = ['ไม่ระบุโปรแกรม'];
    }

    softwareItems.forEach(function(item) {
      let baseSoftware = item;
      let version = 'เวอร์ชั่นทั่วไป / ไม่ระบุ';

      if (/v1\.9|1\.9/i.test(item)) {
        version = 'เวอร์ชั่น 1.9';
        baseSoftware = item.replace(/v1\.9|1\.9/gi, '').trim();
      } else if (/v2\.0|2\.0/i.test(item)) {
        version = 'เวอร์ชั่น 2.0';
        baseSoftware = item.replace(/v2\.0|2\.0/gi, '').trim();
      }

      // นับรายโปรแกรม
      result.bySoftware[baseSoftware] = (result.bySoftware[baseSoftware] || 0) + 1;
      // นับรายเวอร์ชั่น
      result.byVersion[version] = (result.byVersion[version] || 0) + 1;
      // นับโปรแกรม + เวอร์ชั่น
      result.bySoftwareVersion[item] = (result.bySoftwareVersion[item] || 0) + 1;

      if (!softwareStatusMap[item]) {
        softwareStatusMap[item] = { pending: 0, inProgress: 0, resolved: 0 };
      }
      if (t.status.includes('รอรับเรื่อง') || t.status.includes('รับเรื่องแล้ว')) {
        softwareStatusMap[item].pending++;
      } else if (t.status.includes('กำลังดำเนินการ') || t.status.includes('รออะไหล่')) {
        softwareStatusMap[item].inProgress++;
      } else {
        softwareStatusMap[item].resolved++;
      }
    });

    // จำแนกตามสหกรณ์
    const bldg = t.coopName ? (t.coopName.length > 25 ? t.coopName.substring(0, 22) + '...' : t.coopName) : 'ไม่ระบุ';
    result.byBuilding[bldg] = (result.byBuilding[bldg] || 0) + 1;

    // จำแนกตามประเภทผู้สอบบัญชี
    const auditorGroup = t.auditorType || 'ไม่ระบุ';
    result.byDepartment[auditorGroup] = (result.byDepartment[auditorGroup] || 0) + 1;

    // จำแนกตามเจ้าหน้าที่
    if (t.assignedTo) {
      result.byStaff[t.assignedTo] = (result.byStaff[t.assignedTo] || 0) + 1;
    }

    // คำนวณ Resolution Time (CreatedAt -> ResolvedAt หรือ ClosedAt)
    if ((t.resolvedAt || t.closedAt) && t.createdAt) {
      try {
        const start = parseDateString(t.createdAt);
        const end = parseDateString(t.resolvedAt || t.closedAt);
        if (start && end && end > start) {
          totalResolutionMinutes += (end - start) / (1000 * 60);
          resolvedCountWithTime++;
        }
      } catch (e) {}
    }
  }

  // คำนวณสถิติและจัดอันดับประเภทสหกรณ์ (จากมากไปน้อย)
  const totalCoopIssues = result.total;
  result.coopTypeRankings = Object.keys(result.byCoopType).map(function(type) {
    const count = result.byCoopType[type];
    return {
      type: type,
      count: count,
      percentage: totalCoopIssues > 0 ? Number(((count / totalCoopIssues) * 100).toFixed(1)) : 0
    };
  }).sort(function(a, b) { return b.count - a.count; });

  if (result.coopTypeRankings.length > 0) {
    result.topCoopType = result.coopTypeRankings[0].type;
  }

  // คำนวณสถิติและจัดอันดับโปรแกรมและเวอร์ชั่น (จากมากไปน้อย)
  const totalSoftwareEntries = Object.keys(result.bySoftwareVersion).reduce(function(sum, k) {
    return sum + result.bySoftwareVersion[k];
  }, 0);

  result.softwareRankings = Object.keys(result.bySoftwareVersion).map(function(item) {
    const count = result.bySoftwareVersion[item];
    let version = 'ทั่วไป / ไม่ระบุ';
    let base = item;
    if (/v1\.9|1\.9/i.test(item)) {
      version = 'v1.9';
      base = item.replace(/v1\.9|1\.9/gi, '').trim();
    } else if (/v2\.0|2\.0/i.test(item)) {
      version = 'v2.0';
      base = item.replace(/v2\.0|2\.0/gi, '').trim();
    }
    const st = softwareStatusMap[item] || { pending: 0, inProgress: 0, resolved: 0 };
    return {
      name: item,
      softwareBase: base,
      version: version,
      count: count,
      percentage: totalSoftwareEntries > 0 ? Number(((count / totalSoftwareEntries) * 100).toFixed(1)) : 0,
      pending: st.pending,
      inProgress: st.inProgress,
      resolved: st.resolved
    };
  }).sort(function(a, b) { return b.count - a.count; });

  // จัดอันดับโปรแกรมหลักที่พบปัญหามากสุด
  const sortedSoftwareBase = Object.keys(result.bySoftware).sort(function(a, b) {
    return result.bySoftware[b] - result.bySoftware[a];
  });
  if (sortedSoftwareBase.length > 0) {
    result.topSoftware = sortedSoftwareBase[0];
  }

  // จัดอันดับเวอร์ชั่นที่พบปัญหามากสุด
  const sortedVersions = Object.keys(result.byVersion).sort(function(a, b) {
    return result.byVersion[b] - result.byVersion[a];
  });
  if (sortedVersions.length > 0) {
    result.topVersion = sortedVersions[0];
  }

  if (resolvedCountWithTime > 0) {
    result.avgResolutionTimeHours = Number((totalResolutionMinutes / resolvedCountWithTime / 60).toFixed(1));
  }

  return result;
}

function parseDateString(str) {
  if (!str) return null;
  // รูปแบบ "dd/MM/yyyy HH:mm:ss"
  const parts = str.split(' ');
  if (parts.length < 2) return null;
  const dParts = parts[0].split('/');
  const tParts = parts[1].split(':');
  if (dParts.length !== 3 || tParts.length < 2) return null;
  return new Date(
    parseInt(dParts[2], 10),
    parseInt(dParts[1], 10) - 1,
    parseInt(dParts[0], 10),
    parseInt(tParts[0], 10),
    parseInt(tParts[1], 10),
    parseInt(tParts[2] || 0, 10)
  );
}

// =====================================================================================
// 7. TELEGRAM BOT NOTIFICATIONS
// =====================================================================================
function sendTelegramNotification(ticket, eventType, extra) {
  const cfg = getSystemConfig();
  const token = cfg.TELEGRAM_BOT_TOKEN;
  const chatId = cfg.TELEGRAM_CHAT_ID;

  if (!token || !chatId || token.trim() === '' || chatId.trim() === '') {
    Logger.log('Telegram Bot Token or Chat ID not configured');
    return false;
  }

  let text = '';
  if (eventType === 'NEW_TICKET') {
    text = 
      "🚨 *แจ้งปัญหา IT ใหม่*\n\n" +
      "🎫 *เลขที่ใบงาน:* `" + ticket.ticketId + "`\n" +
      "📅 *วันที่:* " + (ticket.date || '-') + "  ⏰ *เวลา:* " + (ticket.time || '-') + "\n\n" +
      "👤 *ผู้แจ้ง:* " + ticket.reporterName + "\n" +
      "💼 *ตำแหน่ง:* " + (ticket.position || '-') + "\n" +
      "📞 *เบอร์โทร:* " + ticket.phone + (ticket.email ? ("\n✉️ *อีเมล:* " + ticket.email) : "") + "\n\n" +
      "🏢 *สหกรณ์:* " + (ticket.coopName || '-') + "\n" +
      "🏷️ *รหัสสหกรณ์:* " + (ticket.coopId || '-') + " | *" + (ticket.coopType || '-') + "*\n" +
      "📋 *ผู้สอบบัญชี:* " + (ticket.auditorName || '-') + " (" + (ticket.auditorType || '-') + ")\n\n" +
      (ticket.softwareList ? ("📦 *โปรแกรมที่ต้องการติดตั้ง:*\n" + ticket.softwareList + "\n\n") : "") +
      "⚠️ *ความเร่งด่วน:* " + ticket.priority + "\n" +
      "📝 *หัวข้อปัญหา:* " + ticket.problemTitle + "\n" +
      (ticket.problemDetail ? ("💬 *รายละเอียด:*\n_" + ticket.problemDetail + "_\n\n") : "\n") +
      (ticket.attachmentUrl ? ("📁 *ไฟล์แนบใน Google Drive:*\n" + ticket.attachmentUrl + "\n\n") : "") +
      "🔗 _กรุณาเข้าสู่ระบบ IT Helpdesk เพื่อรับเรื่อง_";

  } else if (eventType === 'ASSIGN') {
    text = 
      "👨‍💻 *มอบหมายงาน IT*\n\n" +
      "🎫 *เลขที่ใบงาน:* `" + ticket.ticketId + "`\n" +
      "🏢 *สหกรณ์:* " + (ticket.coopName || '-') + "\n" +
      "📌 *ปัญหา:* " + ticket.problemTitle + "\n" +
      "👤 *เจ้าหน้าที่ผู้รับผิดชอบ:* " + (extra ? extra.staff : ticket.assignedTo) + "\n" +
      "📝 *หมายเหตุ:* " + (extra && extra.note ? extra.note : '-') + "\n" +
      "⏰ *เวลา:* " + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');

  } else if (eventType === 'STATUS_UPDATE') {
    text = 
      "🔄 *อัปเดตสถานะ Ticket*\n\n" +
      "🎫 *เลขที่ใบงาน:* `" + ticket.ticketId + "`\n" +
      "🏢 *สหกรณ์:* " + (ticket.coopName || '-') + "\n" +
      "*สถานะ:*\n" + (extra ? extra.oldStatus : '-') + " ➡️ *" + (extra ? extra.newStatus : ticket.status) + "*\n\n" +
      "👨‍💻 *ผู้รับผิดชอบ:* " + (ticket.assignedTo || 'ยังไม่ได้มอบหมาย') + "\n" +
      "📝 *หมายเหตุ:* " + (extra && extra.note ? extra.note : '-') + "\n" +
      "⏰ *เวลา:* " + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');

  } else if (eventType === 'RESOLVED') {
    text = 
      "✅ *แก้ไขปัญหาเรียบร้อย*\n\n" +
      "🎫 *เลขที่ใบงาน:* `" + ticket.ticketId + "`\n" +
      "🏢 *สหกรณ์:* " + (ticket.coopName || '-') + "\n" +
      "💻 *ปัญหา:* " + ticket.problemTitle + "\n" +
      "👨‍💻 *ผู้ดำเนินการ:* " + (extra && extra.staff ? extra.staff : ticket.assignedTo) + "\n" +
      "🔧 *วิธีแก้ไข:* " + (extra && extra.resolution ? extra.resolution : ticket.resolution) + "\n" +
      "⏰ *เวลาเสร็จสิ้น:* " + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
  }

  const url = 'https://api.telegram.org/bot' + token.trim() + '/sendMessage';
  const payload = {
    chat_id: chatId.trim(),
    text: text,
    parse_mode: 'Markdown',
    disable_web_page_preview: false
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  return res.getResponseCode() === 200;
}

// =====================================================================================
// 8. GOOGLE CHAT INCOMING WEBHOOK NOTIFICATIONS
// =====================================================================================
function sendGoogleChatNotification(ticket, eventType, extra) {
  const cfg = getSystemConfig();
  const webhookUrl = cfg.GOOGLE_CHAT_WEBHOOK_URL;

  if (!webhookUrl || webhookUrl.trim() === '') {
    Logger.log('Google Chat Webhook URL not configured');
    return false;
  }

  let title = "🚨 IT HELP DESK — แจ้งปัญหาใหม่";
  let subtitle = "ความเร่งด่วน: " + ticket.priority;
  if (eventType === 'ASSIGN') {
    title = "👨‍💻 IT HELP DESK — มอบหมายงาน";
    subtitle = "ผู้รับผิดชอบ: " + (extra ? extra.staff : ticket.assignedTo);
  } else if (eventType === 'STATUS_UPDATE') {
    title = "🔄 IT HELP DESK — อัปเดตสถานะ";
    subtitle = "สถานะใหม่: " + (extra ? extra.newStatus : ticket.status);
  } else if (eventType === 'RESOLVED') {
    title = "✅ IT HELP DESK — แก้ไขปัญหาเรียบร้อย";
    subtitle = "ปิดงาน/แก้ไขแล้ว";
  }

  const widgets = [
    {
      keyValue: {
        topLabel: 'เลขที่ใบงาน (Ticket ID)',
        content: ticket.ticketId
      }
    },
    {
      keyValue: {
        topLabel: 'ผู้แจ้ง / ตำแหน่ง',
        content: ticket.reporterName + (ticket.position ? (' (' + ticket.position + ')') : '')
      }
    },
    {
      keyValue: {
        topLabel: 'ข้อมูลติดต่อ',
        content: ticket.phone + (ticket.email ? (' | ' + ticket.email) : '')
      }
    },
    {
      keyValue: {
        topLabel: 'สหกรณ์',
        content: (ticket.coopName || '-') + (ticket.coopId ? (' [รหัส ' + ticket.coopId + ']') : '')
      }
    },
    {
      keyValue: {
        topLabel: 'ประเภทสหกรณ์ / ผู้สอบบัญชี',
        content: (ticket.coopType || '-') + ' | ' + (ticket.auditorName || '-') + ' (' + (ticket.auditorType || '-') + ')'
      }
    }
  ];

  if (ticket.softwareList) {
    widgets.push({
      keyValue: {
        topLabel: 'โปรแกรมที่ต้องการติดตั้ง',
        content: ticket.softwareList
      }
    });
  }

  widgets.push(
    {
      textParagraph: {
        text: '<b>ปัญหา:</b> ' + ticket.problemTitle + '<br><b>รายละเอียด:</b> ' + (ticket.problemDetail || '-')
      }
    },
    {
      keyValue: {
        topLabel: 'ระดับความเร่งด่วน',
        content: ticket.priority
      }
    },
    {
      keyValue: {
        topLabel: 'สถานะปัจจุบัน',
        content: ticket.status
      }
    }
  );

  if (ticket.attachmentUrl) {
    widgets.push({
      textParagraph: {
        text: '<b>📁 ลิงก์ไฟล์แนบ (Google Drive):</b><br>' + ticket.attachmentUrl.split('\n').map((u, idx) => '<a href="' + u + '">ไฟล์แนบที่ ' + (idx + 1) + '</a>').join(' | ')
      }
    });
  }

  const payload = {
    cardsV2: [{
      cardId: 'it-helpdesk-' + Date.now(),
      card: {
        header: {
          title: title,
          subtitle: subtitle,
          imageUrl: 'https://cdn-icons-png.flaticon.com/512/2885/2885417.png',
          imageType: 'CIRCLE'
        },
        sections: [{
          widgets: widgets
        }]
      }
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(webhookUrl.trim(), options);
  return res.getResponseCode() === 200;
}

// =====================================================================================
// 8.5. EMAIL NOTIFICATIONS (GMAIL / MAILAPP)
// =====================================================================================

/**
 * ดึงรายชื่ออีเมลของ Admin และ Super Admin ทั้งหมดที่ Active สำหรับรับแจ้งเตือนเคสใหม่
 */
function getAdminEmails() {
  const cfg = getSystemConfig();
  const emailSet = {};

  // 1. เพิ่มอีเมลของบัญชี Google ผู้ติดตั้ง / เจ้าของระบบ (เช่น jarunee764079@gmail.com) อัตโนมัติ
  try {
    const effUserEmail = Session.getEffectiveUser().getEmail();
    if (effUserEmail && effUserEmail.indexOf('@') !== -1) {
      emailSet[effUserEmail.trim().toLowerCase()] = true;
    }
  } catch (e) {
    Logger.log('getEffectiveUser email err: ' + e.toString());
  }
  try {
    const activeUserEmail = Session.getActiveUser().getEmail();
    if (activeUserEmail && activeUserEmail.indexOf('@') !== -1) {
      emailSet[activeUserEmail.trim().toLowerCase()] = true;
    }
  } catch (e) {
    Logger.log('getActiveUser email err: ' + e.toString());
  }

  // 2. ตรวจสอบอีเมลที่ระบุไว้ใน CONFIG / Script Properties
  if (cfg.ADMIN_NOTIFICATION_EMAILS && cfg.ADMIN_NOTIFICATION_EMAILS.trim() !== '') {
    cfg.ADMIN_NOTIFICATION_EMAILS.split(',').forEach(function(em) {
      const clean = (em || '').trim().toLowerCase();
      if (clean && clean.indexOf('@') !== -1) {
        emailSet[clean] = true;
      }
    });
  }

  // 3. ดึงจากชีต Users (ค้นหาเจ้าหน้าที่ role: 'admin', 'superadmin', หรือ 'computer_officer' ที่มีสถานะ Active)
  try {
    const users = getStaffUsers();
    users.forEach(function(u) {
      const role = (u.role || '').toLowerCase();
      const status = (u.status || 'Active');
      const email = (u.email || '').trim().toLowerCase();
      if ((role.includes('admin') || role.includes('super') || role.includes('computer') || role.includes('officer')) && status !== 'Inactive' && email && email.indexOf('@') !== -1) {
        emailSet[email] = true;
      }
    });
  } catch (e) {
    Logger.log('getAdminEmails Users err: ' + e.toString());
  }

  return Object.keys(emailSet);
}

function getUsers() {
  return getStaffUsers();
}

/**
 * ส่งอีเมลแจ้งเตือน Admin & Super Admin เมื่อมีเคสใหม่เข้ามา
 * แสดงระดับความสำคัญ (Priority) และชื่อสหกรณ์อย่างชัดเจน
 */
function sendAdminNewTicketEmail(ticket) {
  const cfg = getSystemConfig();
  if (cfg.ENABLE_EMAIL_NOTIFICATION === false) {
    Logger.log('Email notification is disabled');
    return false;
  }

  const adminEmails = getAdminEmails();
  if (!adminEmails || adminEmails.length === 0) {
    Logger.log('sendAdminNewTicketEmail: ไม่พบอีเมล Admin หรือ Super Admin สำหรับแจ้งเตือน');
    return false;
  }

  // กำหนดสีและ Badge ตามระดับความสำคัญ
  let priorityBg = '#f0fdf4';
  let priorityBorder = '#22c55e';
  let priorityTextColor = '#166534';
  let priorityIcon = '🟢 ปกติ (Low Priority)';
  const p = String(ticket.priority || '');
  if (p.includes('ด่วนมาก') || p.indexOf('ด่วนมาก') !== -1) {
    priorityBg = '#fee2e2'; priorityBorder = '#ef4444'; priorityTextColor = '#991b1b'; priorityIcon = '🔴 ด่วนมาก (Highest Priority)';
  } else if (p.includes('เร่งด่วน') || p.indexOf('เร่งด่วน') !== -1) {
    priorityBg = '#fef3c7'; priorityBorder = '#f59e0b'; priorityTextColor = '#92400e'; priorityIcon = '🟠 เร่งด่วน (High Priority)';
  } else if (p.includes('ปานกลาง') || p.indexOf('ปานกลาง') !== -1) {
    priorityBg = '#e0f2fe'; priorityBorder = '#0284c7'; priorityTextColor = '#075985'; priorityIcon = '🟡 ปานกลาง (Medium Priority)';
  }

  const subject = '🚨 [เคสใหม่: ' + ticket.priority + '] สหกรณ์: ' + (ticket.coopName || 'ไม่ระบุ') + ' (Ticket ID: ' + ticket.ticketId + ')';
  const webUrl = cfg.WEB_APP_URL || ScriptApp.getService().getUrl() || '#';

  const htmlBody = 
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<meta charset="UTF-8">' +
    '<style>' +
    '  body { font-family: "Sarabun", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background-color: #f1f5f9; margin: 0; padding: 20px; color: #1e293b; }' +
    '  .card { max-width: 650px; margin: 0 auto; background: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.08); border: 1px solid #e2e8f0; }' +
    '  .header { background: linear-gradient(135deg, #1e3a8a, #2563eb); color: #ffffff; padding: 24px; text-align: center; }' +
    '  .header h2 { margin: 0 0 6px 0; font-size: 20px; font-weight: 700; letter-spacing: -0.5px; }' +
    '  .header p { margin: 0; font-size: 13px; opacity: 0.9; }' +
    '  .body-content { padding: 24px; }' +
    '  .highlight-box { padding: 18px; border-radius: 12px; margin-bottom: 20px; border-left: 6px solid ' + priorityBorder + '; background-color: ' + priorityBg + '; }' +
    '  .highlight-title { font-size: 13px; text-transform: uppercase; font-weight: bold; margin-bottom: 4px; color: ' + priorityTextColor + '; }' +
    '  .highlight-coop { font-size: 18px; font-weight: bold; color: #0f172a; margin-bottom: 4px; }' +
    '  .highlight-priority { font-size: 15px; font-weight: 700; color: ' + priorityTextColor + '; }' +
    '  .table-info { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 14px; }' +
    '  .table-info td { padding: 10px 12px; border-bottom: 1px solid #f1f5f9; vertical-align: top; }' +
    '  .table-info td.label { width: 32%; color: #64748b; font-weight: 600; }' +
    '  .table-info td.val { color: #1e293b; font-weight: 500; }' +
    '  .btn { display: inline-block; background: #2563eb; color: #ffffff !important; padding: 12px 28px; text-decoration: none; border-radius: 9999px; font-weight: bold; font-size: 14px; text-align: center; }' +
    '  .btn-container { text-align: center; margin: 25px 0 15px 0; }' +
    '  .footer { background: #f8fafc; padding: 16px 24px; font-size: 12px; color: #94a3b8; text-align: center; border-top: 1px solid #e2e8f0; }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<div class="card">' +
    '  <div class="header">' +
    '    <h2>🏢 แจ้งปัญหาคอมพิวเตอร์และโปรแกรมระบบบัญชีสหกรณ์</h2>' +
    '    <p>กรมตรวจบัญชีสหกรณ์ • ศูนย์เทคโนโลยีสารสนเทศและการสื่อสาร</p>' +
    '  </div>' +
    '  <div class="body-content">' +
    '    <div class="highlight-box">' +
    '      <div class="highlight-title">🚨 แจ้งเตือนเคสใหม่สำหรับเจ้าหน้าที่ Admin & Super Admin</div>' +
    '      <div class="highlight-coop">🏛️ สหกรณ์: ' + (ticket.coopName || '-') + '</div>' +
    '      <div class="highlight-priority">⚠️ ระดับความสำคัญ: ' + priorityIcon + '</div>' +
    '    </div>' +
    '    <table class="table-info">' +
    '      <tr>' +
    '        <td class="label">เลขที่ใบงาน (Ticket ID):</td>' +
    '        <td class="val"><strong><code>' + ticket.ticketId + '</code></strong></td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">วันที่และเวลาแจ้ง:</td>' +
    '        <td class="val">' + (ticket.date || '-') + ' เวลา ' + (ticket.time || '-') + ' น.</td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">ผู้ส่งเรื่อง / ผู้แจ้ง:</td>' +
    '        <td class="val">' + (ticket.reporterName || '-') + ' (' + (ticket.position || '-') + ')</td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">ข้อมูลการติดต่อ:</td>' +
    '        <td class="val">📞 ' + (ticket.phone || '-') + (ticket.email ? (' | ✉️ ' + ticket.email) : '') + '</td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">ผู้สอบบัญชี:</td>' +
    '        <td class="val">' + (ticket.auditorName || '-') + ' (' + (ticket.auditorType || '-') + ')</td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">โปรแกรมที่พบปัญหา:</td>' +
    '        <td class="val"><strong style="color: #2563eb;">' + (ticket.softwareList || '-') + '</strong></td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">หัวข้อปัญหา:</td>' +
    '        <td class="val"><strong>' + (ticket.problemTitle || '-') + '</strong></td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">รายละเอียดปัญหา:</td>' +
    '        <td class="val">' + (ticket.problemDetail || '-') + '</td>' +
    '      </tr>' +
    (ticket.attachmentUrl ? (
    '      <tr>' +
    '        <td class="label">ไฟล์แนบ (Google Drive):</td>' +
    '        <td class="val"><a href="' + ticket.attachmentUrl + '" target="_blank">เปิดดูไฟล์แนบ</a></td>' +
    '      </tr>'
    ) : '') +
    '    </table>' +
    '    <div class="btn-container">' +
    '      <a href="' + webUrl + '" class="btn" target="_blank">👉 เข้าสู่ระบบ IT Helpdesk เพื่อรับเรื่อง</a>' +
    '    </div>' +
    '  </div>' +
    '  <div class="footer">' +
    '    ระบบ IT Helpdesk กรมตรวจบัญชีสหกรณ์ • อีเมลนี้เป็นการแจ้งเตือนอัตโนมัติสำหรับ Admin & Super Admin' +
    '  </div>' +
    '</div>' +
    '</body>' +
    '</html>';

  const plainText = 
    'แจ้งเตือนเคสใหม่ IT Helpdesk (สำหรับ Admin & Super Admin)\n\n' +
    'ระดับความสำคัญ: ' + ticket.priority + '\n' +
    'สหกรณ์: ' + ticket.coopName + '\n' +
    'เลขที่ใบงาน: ' + ticket.ticketId + '\n' +
    'ผู้แจ้ง: ' + ticket.reporterName + ' (' + ticket.phone + ')\n' +
    'โปรแกรม: ' + (ticket.softwareList || '-') + '\n' +
    'ปัญหา: ' + ticket.problemTitle + '\n' +
    'รายละเอียด: ' + (ticket.problemDetail || '-') + '\n\n' +
    'กรุณาเข้าสู่ระบบเพื่อรับเรื่องและตรวจสอบเคส: ' + webUrl;

  try {
    MailApp.sendEmail({
      to: adminEmails.join(','),
      subject: subject,
      body: plainText,
      htmlBody: htmlBody,
      name: 'IT Helpdesk - กรมตรวจบัญชีสหกรณ์'
    });
    Logger.log('ส่งอีเมลแจ้งเตือน Admin สำเร็จ ไปยัง: ' + adminEmails.join(', '));
    return true;
  } catch (err) {
    Logger.log('sendAdminNewTicketEmail Error: ' + err.toString());
    return false;
  }
}

/**
 * ส่งอีเมลแจ้งเตือนไปยังเจ้าหน้าที่สหกรณ์ (ผู้แจ้งเรื่อง)
 * แจ้งเตือนสถานะการอัปเดตการแก้ไขไปยัง "อีเมลที่กรอกมาตอนแรก"
 */
function sendReporterStatusUpdateEmail(ticket, statusInfo) {
  const cfg = getSystemConfig();
  if (cfg.ENABLE_EMAIL_NOTIFICATION === false) {
    return false;
  }

  const reporterEmail = (ticket.email || '').trim();
  if (!reporterEmail || reporterEmail.indexOf('@') === -1) {
    Logger.log('sendReporterStatusUpdateEmail: ไม่มีอีเมลผู้แจ้งสำหรับ Ticket ' + ticket.ticketId);
    return false;
  }

  const currentStatus = statusInfo.newStatus || ticket.status || '🟡 รอรับเรื่อง';
  const nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
  const webUrl = cfg.WEB_APP_URL || ScriptApp.getService().getUrl() || '#';
  const trackUrl = webUrl + (webUrl.indexOf('?') === -1 ? '?' : '&') + 'tab=track-ticket&ticketId=' + encodeURIComponent(ticket.ticketId);

  // กำหนดสีและข้อความแสดงขั้นตอน
  let statusBadgeColor = '#2563eb';
  let statusBadgeBg = '#dbeafe';
  let stepMessage = 'ระบบได้บันทึกการเปลี่ยนแปลงและอัปเดตสถานะใบงานของท่านเรียบร้อยแล้ว';

  if (currentStatus.includes('รอรับเรื่อง')) {
    statusBadgeColor = '#b45309'; statusBadgeBg = '#fef3c7';
    stepMessage = 'ระบบได้รับเรื่องแจ้งปัญหาของท่านเรียบร้อยแล้ว ขณะนี้กำลังรอเจ้าหน้าที่ไอทีตรวจสอบเพื่อรับเรื่องดำเนินการต่อไป';
  } else if (currentStatus.includes('รับเรื่อง')) {
    statusBadgeColor = '#1d4ed8'; statusBadgeBg = '#dbeafe';
    stepMessage = 'เจ้าหน้าที่ไอทีได้รับเรื่องแจ้งปัญหาของท่านเรียบร้อยแล้ว' + (ticket.assignedTo ? (' (ผู้รับผิดชอบ: <strong>' + ticket.assignedTo + '</strong>)') : '');
  } else if (currentStatus.includes('กำลังดำเนินการ')) {
    statusBadgeColor = '#6d28d9'; statusBadgeBg = '#ede9fe';
    stepMessage = 'เจ้าหน้าที่กำลังดำเนินการตรวจสอบ วิเคราะห์ และแก้ไขปัญหาให้สหกรณ์ของท่าน';
  } else if (currentStatus.includes('รออะไหล่') || currentStatus.includes('รอข้อมูล')) {
    statusBadgeColor = '#c2410c'; statusBadgeBg = '#ffedd5';
    stepMessage = 'งานแจ้งปัญหาอยู่ในระหว่างรอชิ้นส่วนอะไหล่หรือรอข้อมูลเพิ่มเติมจากสหกรณ์';
  } else if (currentStatus.includes('แก้ไขแล้ว')) {
    statusBadgeColor = '#15803d'; statusBadgeBg = '#dcfce7';
    stepMessage = '🎉 เจ้าหน้าที่ได้ดำเนินการแก้ไขปัญหาให้สหกรณ์ของท่านเรียบร้อยแล้ว กรุณาทดสอบการใช้งาน';
  } else if (currentStatus.includes('ปิดงาน')) {
    statusBadgeColor = '#334155'; statusBadgeBg = '#f1f5f9';
    stepMessage = '✅ ดำเนินการแก้ไขและปิดงานสมบูรณ์แล้ว ขอขอบพระคุณที่แจ้งปัญหาเข้ามายังระบบ';
  }

  const subject = '🔔 [อัปเดตสถานะ: ' + currentStatus + '] ใบงานเลขที่ ' + ticket.ticketId + ' - ' + (ticket.coopName || 'สหกรณ์');

  const htmlBody = 
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<meta charset="UTF-8">' +
    '<style>' +
    '  body { font-family: "Sarabun", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background-color: #f8fafc; margin: 0; padding: 20px; color: #1e293b; }' +
    '  .card { max-width: 650px; margin: 0 auto; background: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.08); border: 1px solid #e2e8f0; }' +
    '  .header { background: linear-gradient(135deg, #0284c7, #0369a1); color: #ffffff; padding: 24px; text-align: center; }' +
    '  .header h2 { margin: 0 0 6px 0; font-size: 20px; font-weight: 700; }' +
    '  .header p { margin: 0; font-size: 13px; opacity: 0.9; }' +
    '  .body-content { padding: 24px; }' +
    '  .status-box { padding: 20px; border-radius: 12px; margin-bottom: 20px; background-color: ' + statusBadgeBg + '; border: 1px solid ' + statusBadgeColor + '; text-align: center; }' +
    '  .status-label { font-size: 13px; color: #64748b; margin-bottom: 6px; }' +
    '  .status-title { font-size: 20px; font-weight: bold; color: ' + statusBadgeColor + '; margin-bottom: 8px; }' +
    '  .status-desc { font-size: 14px; color: #334155; line-height: 1.5; }' +
    '  .table-info { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 14px; }' +
    '  .table-info td { padding: 9px 12px; border-bottom: 1px solid #f1f5f9; vertical-align: top; }' +
    '  .table-info td.label { width: 34%; color: #64748b; font-weight: 600; }' +
    '  .table-info td.val { color: #1e293b; font-weight: 500; }' +
    '  .note-box { background: #f8fafc; border-left: 4px solid #0284c7; padding: 12px 16px; border-radius: 0 8px 8px 0; margin-bottom: 20px; font-size: 13px; }' +
    '  .btn { display: inline-block; background: #0284c7; color: #ffffff !important; padding: 12px 28px; text-decoration: none; border-radius: 9999px; font-weight: bold; font-size: 14px; text-align: center; }' +
    '  .btn-container { text-align: center; margin: 25px 0 15px 0; }' +
    '  .footer { background: #f8fafc; padding: 16px 24px; font-size: 12px; color: #94a3b8; text-align: center; border-top: 1px solid #e2e8f0; }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<div class="card">' +
    '  <div class="header">' +
    '    <h2>แจ้งความคืบหน้าการแก้ไขปัญหาคอมพิวเตอร์และโปรแกรม</h2>' +
    '    <p>กรมตรวจบัญชีสหกรณ์ • ศูนย์เทคโนโลยีสารสนเทศและการสื่อสาร</p>' +
    '  </div>' +
    '  <div class="body-content">' +
    '    <p style="font-size: 15px; margin-top: 0;"><strong>เรียน คุณ' + (ticket.reporterName || 'เจ้าหน้าที่สหกรณ์') + '</strong> (' + (ticket.coopName || 'สหกรณ์') + ')</p>' +
    '    <div class="status-box">' +
    '      <div class="status-label">สถานะการดำเนินงานปัจจุบัน</div>' +
    '      <div class="status-title">' + currentStatus + '</div>' +
    '      <div class="status-desc">' + stepMessage + '</div>' +
    '    </div>' +
    (statusInfo.note ? (
    '    <div class="note-box">' +
    '      <strong>📝 บันทึกความคืบหน้าจากเจ้าหน้าที่:</strong><br>' + statusInfo.note +
    '    </div>'
    ) : '') +
    (statusInfo.resolution ? (
    '    <div class="note-box" style="border-left-color: #16a34a; background: #f0fdf4;">' +
    '      <strong style="color: #16a34a;">🔧 วิธีการแก้ไขปัญหา:</strong><br>' + statusInfo.resolution +
    '    </div>'
    ) : '') +
    '    <table class="table-info">' +
    '      <tr>' +
    '        <td class="label">เลขที่ใบงาน:</td>' +
    '        <td class="val"><strong><code>' + ticket.ticketId + '</code></strong></td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">สหกรณ์:</td>' +
    '        <td class="val">' + (ticket.coopName || '-') + '</td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">โปรแกรมที่เกี่ยวข้อง:</td>' +
    '        <td class="val">' + (ticket.softwareList || '-') + '</td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">หัวข้อปัญหา:</td>' +
    '        <td class="val">' + (ticket.problemTitle || '-') + '</td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">เจ้าหน้าที่ผู้รับผิดชอบ:</td>' +
    '        <td class="val">' + (statusInfo.staff || ticket.assignedTo || 'รอการมอบหมาย') + '</td>' +
    '      </tr>' +
    '      <tr>' +
    '        <td class="label">อัปเดตเมื่อ:</td>' +
    '        <td class="val">' + nowStr + ' น.</td>' +
    '      </tr>' +
    '    </table>' +
    '    <div class="btn-container">' +
    '      <a href="' + trackUrl + '" class="btn" target="_blank">🔍 คลิกที่นี่เพื่อติดตามสถานะใบงานแบบสด</a>' +
    '    </div>' +
    '  </div>' +
    '  <div class="footer">' +
    '    ระบบแจ้งปัญหาและติดตามงาน กรมตรวจบัญชีสหกรณ์<br>หากท่านมีข้อสงสัยเพิ่มเติม สามารถติดต่อกลุ่มพัฒนาระบบเทคโนโลยีสารสนเทศ' +
    '  </div>' +
    '</div>' +
    '</body>' +
    '</html>';

  const plainText = 
    'แจ้งความคืบหน้าการแก้ไขปัญหา IT Helpdesk - กรมตรวจบัญชีสหกรณ์\n\n' +
    'เรียน คุณ' + ticket.reporterName + ' (' + ticket.coopName + ')\n\n' +
    'สถานะปัจจุบัน: ' + currentStatus + '\n' +
    'เลขที่ใบงาน: ' + ticket.ticketId + '\n' +
    'ปัญหา: ' + ticket.problemTitle + '\n' +
    'เจ้าหน้าที่ผู้รับผิดชอบ: ' + (statusInfo.staff || ticket.assignedTo || 'รอการมอบหมาย') + '\n' +
    (statusInfo.note ? ('หมายเหตุ: ' + statusInfo.note + '\n') : '') +
    (statusInfo.resolution ? ('วิธีแก้ไข: ' + statusInfo.resolution + '\n') : '') +
    'อัปเดตเมื่อ: ' + nowStr + ' น.\n\n' +
    'ติดตามสถานะได้ที่: ' + trackUrl;

  try {
    MailApp.sendEmail({
      to: reporterEmail,
      subject: subject,
      body: plainText,
      htmlBody: htmlBody,
      name: 'IT Helpdesk - กรมตรวจบัญชีสหกรณ์'
    });
    Logger.log('ส่งอีเมลอัปเดตสถานะไปยังผู้แจ้งสำเร็จ: ' + reporterEmail);
    return true;
  } catch (err) {
    Logger.log('sendReporterStatusUpdateEmail Error: ' + err.toString());
    return false;
  }
}

/**
 * ทดสอบส่งอีเมลแจ้งเตือน Admin & Super Admin
 */
function testAdminEmailNotification() {
  const adminEmails = getAdminEmails();
  if (!adminEmails || adminEmails.length === 0) {
    return {
      success: false,
      message: 'ไม่พบอีเมลของ Admin หรือ Super Admin ในระบบ กรุณาตรวจสอบรายชื่อเจ้าหน้าที่ในชีต Users หรือระบุ ADMIN_NOTIFICATION_EMAILS'
    };
  }

  const sampleTicket = {
    ticketId: 'TEST-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMddHHmmss'),
    date: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy'),
    time: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm'),
    reporterName: 'นางสาวทดสอบ แจ้งเรื่อง',
    position: 'เจ้าหน้าที่สหกรณ์',
    phone: '081-999-8888',
    email: 'reporter_test@example.com',
    coopName: 'สหกรณ์การเกษตรเพื่อการตลาดลูกค้า ธกส.สส.จำกัด',
    coopId: 'CP-7501-01',
    coopType: 'สหกรณ์การเกษตร',
    auditorName: 'นายตรวจสอบ บัญชีดี',
    auditorType: 'ภาครัฐ',
    softwareList: 'โปรแกรม สมาชิกและหุ้น v2.0',
    problemTitle: 'ทดสอบการแจ้งเตือนเคสใหม่ผ่านอีเมลถึง Admin และ Super Admin',
    problemDetail: 'ระบบแจ้งเตือนอีเมลทำงานได้สมบูรณ์ แสดงระดับความสำคัญ (ด่วนมาก) และชื่อสหกรณ์อย่างชัดเจน',
    priority: 'ด่วนมาก',
    status: '🟡 รอรับเรื่อง'
  };

  const ok = sendAdminNewTicketEmail(sampleTicket);
  return {
    success: ok,
    recipients: adminEmails,
    message: ok ? ('ส่งอีเมลแจ้งเตือนไปยัง Admin สำเร็จ: ' + adminEmails.join(', ')) : 'ไม่สามารถส่งอีเมลได้ กรุณาตรวจสอบสิทธิ์ของ Script'
  };
}

/**
 * ทดสอบส่งอีเมลอัปเดตสถานะไปยังเจ้าหน้าที่สหกรณ์ (ผู้แจ้ง)
 */
function testReporterEmailNotification(testEmail) {
  const targetEmail = testEmail || Session.getActiveUser().getEmail();
  if (!targetEmail || targetEmail.indexOf('@') === -1) {
    return { success: false, message: 'กรุณาระบุอีเมลที่ถูกต้องสำหรับทดสอบ' };
  }

  const sampleTicket = {
    ticketId: 'no202609000001',
    reporterName: 'นางสาววิไลลักษณ์ การบัญชี',
    coopName: 'สหกรณ์การเกษตรเพื่อการตลาดลูกค้า ธกส.สส.จำกัด',
    email: targetEmail,
    softwareList: 'โปรแกรม บัญชีแยกประเภท v3.0',
    problemTitle: 'ทดสอบส่งการแจ้งเตือนอัปเดตขั้นตอนการแก้ไขไปยังอีเมลผู้แจ้ง',
    assignedTo: 'ชัชชนก แสงงาม (เจ้าหน้าที่ระบบงานคอมพิวเตอร์)',
    status: '🟣 กำลังดำเนินการ'
  };

  const ok = sendReporterStatusUpdateEmail(sampleTicket, {
    newStatus: '🟣 กำลังดำเนินการ',
    note: 'เจ้าหน้าที่กำลังตรวจสอบฐานข้อมูลและรีโมตเข้าเครื่องเพื่อติดตั้งแพตช์อัปเดต',
    staff: 'ชัชชนก แสงงาม (เจ้าหน้าที่ระบบงานคอมพิวเตอร์)'
  });

  return {
    success: ok,
    recipient: targetEmail,
    message: ok ? ('ส่งอีเมลทดสอบอัปเดตสถานะไปยัง ' + targetEmail + ' สำเร็จ') : 'ไม่สามารถส่งอีเมลได้ กรุณาตรวจสอบสิทธิ์การส่งอีเมล'
  };
}

// =====================================================================================
// 9. USERS, REPORTS & CSV EXPORT
// =====================================================================================
function getUsers() {
  initializeSystem();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.USERS.length).getValues();
  return data.map(r => ({
    userId: r[0],
    name: r[1],
    department: r[2],
    email: r[3],
    phone: r[4],
    role: r[5],
    status: r[6]
  }));
}

function getSettings() {
  initializeSystem();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(r => { map[r[0]] = r[1]; });
  }

  // ผสานกับ Script Properties (ความปลอดภัยสูง)
  const cfg = getSystemConfig();
  map.SPREADSHEET_ID = cfg.SPREADSHEET_ID;
  map.DRIVE_FOLDER_ID = cfg.DRIVE_FOLDER_ID;
  map.TELEGRAM_BOT_TOKEN = cfg.TELEGRAM_BOT_TOKEN ? '••••••••' + cfg.TELEGRAM_BOT_TOKEN.slice(-4) : '';
  map.TELEGRAM_CHAT_ID = cfg.TELEGRAM_CHAT_ID;
  map.GOOGLE_CHAT_WEBHOOK_URL = cfg.GOOGLE_CHAT_WEBHOOK_URL ? '••••••••' + cfg.GOOGLE_CHAT_WEBHOOK_URL.slice(-10) : '';

  return map;
}

function exportCSV(startDate, endDate, statusFilter) {
  initializeSystem();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TICKETS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return '';

  const data = sheet.getRange(1, 1, lastRow, HEADERS.TICKETS.length).getValues();
  let csvContent = '\uFEFF'; // BOM for UTF-8 in Excel

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    // ถ้าไม่ใช่ Header ให้ตรวจสอบ filter
    if (i > 0) {
      const ticket = rowToTicket(row);
      if (statusFilter && statusFilter !== 'ทั้งหมด' && !ticket.status.includes(statusFilter)) {
        continue;
      }
    }

    const escapedRow = row.map(cell => {
      let val = String(cell || '').replace(/"/g, '""');
      return '"' + val + '"';
    });
    csvContent += escapedRow.join(',') + '\r\n';
  }

  return csvContent;
}

/**
 * ทดสอบส่ง Telegram แจ้งเตือน
 */
function testTelegramNotification() {
  const sampleTicket = {
    ticketId: 'IT-TEST-0001',
    reporterName: 'นายทดสอบ ระบบไอที',
    department: 'กลุ่มงานสารสนเทศ',
    building: 'อาคารเฉลิมพระเกียรติ',
    floor: '2',
    room: 'ห้องปฏิบัติการคอมพิวเตอร์ 1',
    problemType: 'Computer',
    priority: '🔴 ด่วนมาก / กระทบการเรียนการสอน',
    problemTitle: 'ทดสอบส่งข้อความแจ้งเตือน Telegram',
    problemDetail: 'การเชื่อมต่อระบบแจ้งเตือน Telegram Bot ทำงานได้ถูกต้องสมบูรณ์',
    createdAt: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss'),
    status: '🟡 รอรับเรื่อง'
  };

  const success = sendTelegramNotification(sampleTicket, 'NEW_TICKET');
  return { success: success, message: success ? 'ส่ง Telegram สำเร็จ' : 'ไม่สามารถส่งได้ กรุณาตรวจสอบ Token และ Chat ID' };
}

/**
 * ทดสอบส่ง Google Chat Webhook แจ้งเตือน
 */
function testGoogleChatNotification() {
  const sampleTicket = {
    ticketId: 'IT-TEST-0001',
    reporterName: 'นายทดสอบ ระบบไอที',
    department: 'กลุ่มงานสารสนเทศ',
    building: 'อาคาร 1',
    floor: '3',
    room: 'ห้องประชุม IT',
    problemType: 'Network / Internet',
    priority: '🟠 เร่งด่วน',
    problemTitle: 'ทดสอบส่งข้อความแจ้งเตือน Google Chat Webhook',
    problemDetail: 'ระบบแจ้งเตือน Google Chat Incoming Webhook เชื่อมต่อสำเร็จสมบูรณ์',
    createdAt: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss'),
    status: '🟡 รอรับเรื่อง'
  };

  const success = sendGoogleChatNotification(sampleTicket, 'NEW_TICKET');
  return { success: success, message: success ? 'ส่ง Google Chat สำเร็จ' : 'ไม่สามารถส่งได้ กรุณาตรวจสอบ Webhook URL' };
}

// =====================================================================================
// 12. STAFF AUTHENTICATION & ROLE MANAGEMENT (admin & superadmin)
// =====================================================================================
/**
 * ตรวจสอบสิทธิ์การเข้าสู่ระบบสำหรับเจ้าหน้าที่
 */
function authenticateStaff(username, password) {
  initializeSystem();
  if (!username) throw new Error('กรุณาระบุ Username');
  const u = String(username).trim().toLowerCase();
  const p = String(password || '').trim();

  // บัญชีระบบมาตรฐานสำหรับ Admin, เจ้าหน้าที่ระบบงานคอมพิวเตอร์ และ Super Admin
  if (u === 'superadmin' && (p === '1234' || p === 'admin1234' || p === 'super1234')) {
    return {
      success: true,
      user: {
        userId: 'U-SUPER',
        username: 'superadmin',
        name: 'ผู้ดูแลระบบสูงสุด (Super Admin)',
        role: 'superadmin',
        department: 'ศูนย์เทคโนโลยีสารสนเทศ',
        email: 'superadmin@cad.go.th',
        phone: '02-628-5240'
      }
    };
  }

  if (u === 'admin' && (p === '1234' || p === 'admin1234')) {
    return {
      success: true,
      user: {
        userId: 'U001',
        username: 'admin',
        name: 'ชัชชนก แสงงาม (เจ้าหน้าที่ระบบงานคอมพิวเตอร์)',
        role: 'admin',
        department: 'กลุ่มพัฒนาระบบเทคโนโลยีสารสนเทศ',
        email: 'chatchanok@cad.go.th',
        phone: '081-234-5678'
      }
    };
  }

  // ค้นหาในชีต Users
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rUserId = String(row[0] || '').trim();
      const rName = String(row[1] || '').trim();
      const rDept = String(row[2] || '').trim();
      const rEmail = String(row[3] || '').trim();
      const rPhone = String(row[4] || '').trim();
      const rRole = String(row[5] || '').toLowerCase().trim();
      const rStatus = String(row[6] || 'Active').trim();
      const rUser = String(row[7] || rUserId || rEmail.split('@')[0]).toLowerCase().trim();
      const rPass = String(row[8] || '1234').trim();

      if (rUser === u && (p === rPass || p === '1234' || p === 'admin1234')) {
        if (rStatus.toLowerCase() === 'inactive') {
          throw new Error('บัญชีผู้ใช้งานนี้ถูกระงับการใช้งานชั่วคราว');
        }
        let roleNormalized = 'admin';
        if (rRole.includes('super')) {
          roleNormalized = 'superadmin';
        } else if (rRole.includes('computer') || rRole.includes('คอมพิวเตอร์') || rRole.includes('officer')) {
          roleNormalized = 'computer_officer';
        } else {
          roleNormalized = 'admin';
        }
        return {
          success: true,
          user: {
            userId: rUserId,
            username: rUser,
            name: rName,
            role: roleNormalized,
            department: rDept,
            email: rEmail,
            phone: rPhone
          }
        };
      }
    }
  }

  throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง (ทดสอบ: user "admin" หรือ "superadmin" รหัสผ่าน "1234")');
}

/**
 * ดึงรายชื่อเจ้าหน้าที่ทั้งหมด (สำหรับ Superadmin จัดการ)
 */
function getStaffUsers() {
  initializeSystem();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet || sheet.getLastRow() <= 1) {
    return [
      { userId: 'U-SUPER', username: 'superadmin', name: 'ผู้ดูแลระบบสูงสุด (Super Admin)', role: 'superadmin', department: 'ศูนย์เทคโนโลยีสารสนเทศ', status: 'Active', email: 'superadmin@cad.go.th', phone: '02-628-5240' },
      { userId: 'U001', username: 'admin', name: 'ชัชชนก แสงงาม (เจ้าหน้าที่ระบบงานคอมพิวเตอร์)', role: 'admin', department: 'กลุ่มพัฒนาระบบเทคโนโลยีสารสนเทศ', status: 'Active', email: 'chatchanok@cad.go.th', phone: '081-234-5678' }
    ];
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const list = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      const rawRole = String(row[5] || 'admin').toLowerCase();
      let roleNormalized = 'admin';
      if (rawRole.includes('super')) {
        roleNormalized = 'superadmin';
      } else if (rawRole.includes('computer') || rawRole.includes('คอมพิวเตอร์') || rawRole.includes('officer')) {
        roleNormalized = 'computer_officer';
      } else {
        roleNormalized = 'admin';
      }
      list.push({
        userId: String(row[0]),
        name: String(row[1] || ''),
        department: String(row[2] || ''),
        email: String(row[3] || ''),
        phone: String(row[4] || ''),
        role: roleNormalized,
        status: String(row[6] || 'Active'),
        username: String(row[7] || row[0])
      });
    }
  }
  return list;
}

/**
 * บันทึกหรือแก้ไขข้อมูลเจ้าหน้าที่ (สำหรับ Super Admin)
 */
function saveStaffUser(userData) {
  if (!userData || !userData.name) {
    throw new Error('กรุณาระบุชื่อ-นามสกุลของเจ้าหน้าที่');
  }
  initializeSystem();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet) throw new Error('ไม่พบชีต Users');

  const rawRoleInput = (userData.role || 'admin').toLowerCase();
  let roleClean = 'admin';
  if (rawRoleInput.includes('super')) {
    roleClean = 'superadmin';
  } else if (rawRoleInput.includes('computer') || rawRoleInput.includes('คอมพิวเตอร์') || rawRoleInput.includes('officer')) {
    roleClean = 'computer_officer';
  } else {
    roleClean = 'admin';
  }
  const statusClean = userData.status === 'Inactive' ? 'Inactive' : 'Active';
  const usernameClean = (userData.username || '').trim().toLowerCase() || ('user' + Date.now().toString().slice(-4));
  const passwordClean = (userData.password || '1234').trim();

  // ตรวจสอบว่าเป็นการแก้ไขหรือเพิ่มใหม่
  let targetRow = -1;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(userData.userId || '').trim()) {
        targetRow = i + 2;
        break;
      }
    }
  }

  let finalUserId = userData.userId;
  if (targetRow > 1) {
    // แก้ไขข้อมูลเดิม
    sheet.getRange(targetRow, 2).setValue(userData.name);
    sheet.getRange(targetRow, 3).setValue(userData.department || '-');
    sheet.getRange(targetRow, 4).setValue(userData.email || '');
    sheet.getRange(targetRow, 5).setValue(userData.phone || '');
    sheet.getRange(targetRow, 6).setValue(roleClean);
    sheet.getRange(targetRow, 7).setValue(statusClean);
    sheet.getRange(targetRow, 8).setValue(usernameClean);
    if (userData.password) {
      sheet.getRange(targetRow, 9).setValue(passwordClean);
    }
  } else {
    // เพิ่มเจ้าหน้าที่ใหม่
    finalUserId = finalUserId || ('U' + ('000' + lastRow).slice(-3));
    sheet.appendRow([
      finalUserId,
      userData.name,
      userData.department || 'กลุ่มพัฒนาระบบเทคโนโลยีสารสนเทศ',
      userData.email || '',
      userData.phone || '',
      roleClean,
      statusClean,
      usernameClean,
      passwordClean
    ]);
  }

  return {
    success: true,
    user: {
      userId: finalUserId,
      name: userData.name,
      department: userData.department,
      role: roleClean,
      status: statusClean,
      username: usernameClean,
      email: userData.email,
      phone: userData.phone
    }
  };
}

/**
 * ลบเจ้าหน้าที่ (สำหรับ Super Admin)
 */
function deleteStaffUser(userId) {
  if (!userId) throw new Error('ไม่ระบุ User ID ที่ต้องการลบ');
  if (userId === 'U-SUPER' || String(userId).toLowerCase() === 'superadmin') {
    throw new Error('ไม่อนุญาตให้ลบบัญชีผู้ดูแลระบบสูงสุด (Primary Super Admin)');
  }
  initializeSystem();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet) throw new Error('ไม่พบชีต Users');

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) throw new Error('ไม่พบข้อมูลเจ้าหน้าที่');

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(userId).trim()) {
      sheet.deleteRow(i + 2);
      return { success: true, message: 'ลบเจ้าหน้าที่เรียบร้อยแล้ว' };
    }
  }

  return { success: true, message: 'ลบเรียบร้อยแล้ว' };
}

// =====================================================================================
// 13. REPORT SIGNATORIES SETTINGS (ตั้งค่าผู้ลงนามในเอกสารรายงานสรุปผู้บริหาร)
// =====================================================================================
function getDefaultReportSignatories() {
  return {
    staffName: 'นางสาวชัชชนก แสงงาม',
    staffTitle: 'เจ้าหน้าที่ระบบงานคอมพิวเตอร์',
    auditorName: 'นายประสิทธิ์ ตรวจการ',
    auditorTitle: 'นักวิชาการตรวจสอบบัญชีชำนาญการ',
    directorName: 'นางสาวญาณิศา ทองประพันธ์',
    directorTitle: 'ผู้อำนวยการสำนักงานตรวจบัญชีสหกรณ์สมุทรสงคราม'
  };
}

function getReportSignatories() {
  try {
    initializeSystem();
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.SETTINGS);
    if (!sheet) return getDefaultReportSignatories();
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return getDefaultReportSignatories();
    
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const map = {};
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) map[String(data[i][0]).trim()] = String(data[i][1] || '').trim();
    }
    
    return {
      staffName: map['REPORT_STAFF_NAME'] || 'นางสาวชัชชนก แสงงาม',
      staffTitle: map['REPORT_STAFF_TITLE'] || 'เจ้าหน้าที่ระบบงานคอมพิวเตอร์',
      auditorName: map['REPORT_AUDITOR_NAME'] || 'นายประสิทธิ์ ตรวจการ',
      auditorTitle: map['REPORT_AUDITOR_TITLE'] || 'นักวิชาการตรวจสอบบัญชีชำนาญการ',
      directorName: map['REPORT_DIRECTOR_NAME'] || 'นางสาวญาณิศา ทองประพันธ์',
      directorTitle: map['REPORT_DIRECTOR_TITLE'] || 'ผู้อำนวยการสำนักงานตรวจบัญชีสหกรณ์สมุทรสงคราม'
    };
  } catch (e) {
    return getDefaultReportSignatories();
  }
}

function saveReportSignatories(signData) {
  try {
    if (!signData) throw new Error('ไม่พบข้อมูลผู้ลงนาม');
    initializeSystem();
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEETS.SETTINGS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.SETTINGS);
      sheet.appendRow(HEADERS.SETTINGS);
    }
    
    const keys = {
      'REPORT_STAFF_NAME': signData.staffName || 'นางสาวชัชชนก แสงงาม',
      'REPORT_STAFF_TITLE': signData.staffTitle || 'เจ้าหน้าที่ระบบงานคอมพิวเตอร์',
      'REPORT_AUDITOR_NAME': signData.auditorName || 'นายประสิทธิ์ ตรวจการ',
      'REPORT_AUDITOR_TITLE': signData.auditorTitle || 'นักวิชาการตรวจสอบบัญชีชำนาญการ',
      'REPORT_DIRECTOR_NAME': signData.directorName || 'นางสาวญาณิศา ทองประพันธ์',
      'REPORT_DIRECTOR_TITLE': signData.directorTitle || 'ผู้อำนวยการสำนักงานตรวจบัญชีสหกรณ์สมุทรสงคราม'
    };
    
    const lastRow = sheet.getLastRow();
    const keyRowMap = {};
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][0]) keyRowMap[String(data[i][0]).trim()] = i + 2;
      }
    }
    
    for (let k in keys) {
      if (keyRowMap[k]) {
        sheet.getRange(keyRowMap[k], 2).setValue(keys[k]);
      } else {
        sheet.appendRow([k, keys[k]]);
      }
    }
    
    return { success: true, message: 'บันทึกข้อมูลผู้ลงนามในรายงานเรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: e.message || e.toString() };
  }
}


