// 1. นำโค้ดนี้ไปวางใน Google Apps Script ของคุณ
// 2. กดเรียกใช้ "setupSheet" เพื่อสร้างหัวคอลัมน์ใหม่ให้ครบทุกตัวแปร
// 3. กด Deploy เป็น Web App (สิทธิ์เข้าถึง: ทุกคน) และนำ URL ไปใส่ใน HTML

const SHEET_NAME = "SurveyData_Full";

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    // sheet.clear(); // ปิดไว้เพื่อป้องกันข้อมูลเก่าหาย หากต้องการล้างตารางให้เอา // ออก
  }

  // สร้างชุดหัวคอลัมน์ตามแบบสอบถามจริงทั้งหมด
  const headers = [
    "Timestamp", "Consent", 
    // ส่วนที่ 1
    "Gender", "Age", "Age_Group", "Education", "Occupation", "Occupation_Other", "Income", "Marital",
    "Weight", "Height", "BMI", "Waist", "BMI_Group", "Duration", "Comorbidity", "Insurance", "Drugs",
    // ส่วนที่ 2 ด้านที่ 1 อาหาร
    "Diet_1","Diet_2","Diet_3","Diet_4","Diet_5","Diet_6","Diet_7","Diet_8","Diet_9","Diet_10","Diet_11","Diet_12",
    // ส่วนที่ 2 ด้านที่ 2 ออกกำลังกาย
    "Ex_13","Ex_14","Ex_15","Ex_16","Ex_17","Ex_18","Ex_19","Ex_20","Ex_21","Ex_22",
    // ส่วนที่ 2 ด้านที่ 3 บุหรี่
    "Smoke_Status", "Smoke_Quit_Years", "Smoke_Quit_Months", "Smoke_23","Smoke_24","Smoke_25","Smoke_26","Smoke_27","Smoke_28","Smoke_29","Smoke_30",
    // ส่วนที่ 2 ด้านที่ 4 แอลกอฮอล์
    "Alc_Status", "Alc_Quit_Years", "Alc_Quit_Months", "Alc_31","Alc_32","Alc_33","Alc_34","Alc_35","Alc_36","Alc_37","Alc_38","Alc_39","Alc_40",
    // ส่วนที่ 3 การปฏิบัติตามคำแนะนำ (ยา, นัด, ปรับวิถีชีวิต)
    "Med_1","Med_2","Med_3","Med_4","Med_5","Med_6","Med_7","Med_8","Med_9","Med_10","Med_11","Med_12",
    "Appt_13","Appt_14","Appt_15","Appt_16","Appt_17","Appt_18","Appt_19","Appt_20",
    "Life_21","Life_22","Life_23","Life_24","Life_25","Life_26","Life_27","Life_28","Life_29","Life_30",
    // ส่วนที่ 4.1 ความรู้
    "Know_1","Know_2","Know_3","Know_4","Know_5","Know_6","Know_7","Know_8","Know_9","Know_10",
    "Know_11","Know_12","Know_13","Know_14","Know_15","Know_16","Know_17","Know_18","Know_19","Know_20",
    // ส่วนที่ 4.2 การรับรู้
    "Perc_Susc_1","Perc_Susc_2","Perc_Susc_3","Perc_Susc_4","Perc_Susc_5",
    "Perc_Sev_6","Perc_Sev_7","Perc_Sev_8","Perc_Sev_9","Perc_Sev_10",
    "Perc_Ben_11","Perc_Ben_12","Perc_Ben_13","Perc_Ben_14","Perc_Ben_15",
    "Perc_Bar_16","Perc_Bar_17","Perc_Bar_18","Perc_Bar_19","Perc_Bar_20",
    "Perc_Eff_21","Perc_Eff_22","Perc_Eff_23","Perc_Eff_24","Perc_Eff_25",
    // ส่วนที่ 5 แรงสนับสนุน
    "Supp_Emo_1","Supp_Emo_2","Supp_Emo_3","Supp_Emo_4","Supp_Emo_5",
    "Supp_App_6","Supp_App_7","Supp_App_8","Supp_App_9","Supp_App_10",
    "Supp_Info_11","Supp_Info_12","Supp_Info_13","Supp_Info_14","Supp_Info_15",
    "Supp_Inst_16","Supp_Inst_17","Supp_Inst_18","Supp_Inst_19","Supp_Inst_20",
    // ส่วนที่ 6 ปัญหา
    "Prob_Control", "Prob_Compliance", "Suggestions",
    // ส่วนเวชระเบียน
    "Lab_TC", "Lab_LDL", "Lab_HDL", "Lab_TG", "CV_Risk", "Target_LDL", "Control_Status"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#1e3a8a");
  headerRange.setFontColor("white");
  sheet.setFrozenRows(1);
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = JSON.parse(e.postData.contents);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // อัปเดตข้อมูลในแถวที่มีอยู่แล้ว (Lab + ข้อมูลที่แก้ไข)
    if (data.action === "update") {
      const rowIndex = parseInt(data.rowIndex) + 2; // +2: 0-based index + header row
      const skip = new Set(["action", "rowIndex"]);
      Object.keys(data).forEach(field => {
        if (skip.has(field)) return;
        const col = headers.indexOf(field) + 1;
        if (col > 0) sheet.getRange(rowIndex, col).setValue(data[field]);
      });
      return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // เพิ่มแถวใหม่ (submit แบบสอบถาม)
    const rowData = headers.map(header => {
      if (header === "Timestamp") return new Date();
      if (Array.isArray(data[header])) return data[header].join(", ");
      return data[header] || "";
    });
    sheet.appendRow(rowData);

    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);

  const headers = data[0];
  const rows = data.slice(1);
  const jsonArray = rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => { obj[header] = row[index]; });
    return obj;
  });

  return ContentService.createTextOutput(JSON.stringify(jsonArray))
    .setMimeType(ContentService.MimeType.JSON);
}