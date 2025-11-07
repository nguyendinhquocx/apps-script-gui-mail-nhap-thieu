/**
 * GOOGLE SHEETS EMAIL REMINDER - PERSONALIZED & FILTERED
 * Tự động gửi email nhắc nhở nhân viên điền thiếu data
 * FILTER: Chỉ tháng <= hiện tại, PERSONALIZED với tên nhân viên
 * FIX: 3 fields có delay time - ngày hóa đơn, doanh thu thực hiện, tháng GNDT
 */

// =============== CONFIGURATION ===============
// const EMPLOYEE_EMAIL = "quoc.nguyen3@hoanmy.com"; // THAY EMAIL NÀY CHO TỪNG SHEET
// const EMPLOYEE_EMAIL = "quynh.bui@hoanmy.com, luan.tran@hoanmy.com, quoc.nguyen3@hoanmy.com";
const EMPLOYEE_EMAIL = "quoc.nguyen3@hoanmy.com";
const SHEET_NAME = "file nhap chc"; // Tên sheet chứa data
const EMAIL_SUBJECT = "Cập nhật thông tin hợp đồng";

// Map columns to readable names với priority order
const REQUIRED_FIELDS = {
  'D': 'ngày ký hợp đồng',     // HIGH PRIORITY
  'F': 'doanh thu',            // HIGH PRIORITY - chỉ tính trống, không tính 0
  'G': 'số người khám',        // HIGH PRIORITY - chỉ tính trống, không tính 0
  'C': 'mã hợp đồng',
  'E': 'trạng thái ký',
  'H': 'ngày hóa đơn',         // Chỉ tính từ tháng hiện tại - 2 trở xuống
  'I': 'doanh thu thực hiện',  // Chỉ tính từ tháng hiện tại - 2 trở xuống
  'J': 'ngày bắt đầu khám',
  'K': 'ngày kết thúc khám',
  'L': 'tháng GNDT'            // Chỉ tính từ tháng hiện tại - 2 trở xuống
};

// Priority fields lên đầu trong mỗi tháng
const PRIORITY_FIELDS = ['ngày ký hợp đồng', 'doanh thu', 'số người khám'];
// Skip fields cho tháng hiện tại (vì sẽ có ở tháng sau)
const SKIP_IN_CURRENT_MONTH = ['ngày hóa đơn', 'doanh thu thực hiện'];
// Fields chỉ tính trống, không tính 0 - TRIỆT ĐỂ
const NUMERIC_FIELDS = ['doanh thu', 'số người khám', 'doanh thu thực hiện'];
// Fields có delay time - chỉ tính từ tháng hiện tại - 2 trở xuống
const DELAY_FIELDS = ['ngày hóa đơn', 'doanh thu thực hiện', 'tháng GNDT'];

// =============== MAIN FUNCTIONS ===============

function dailyEmailCheck() {
  try {
    console.log("=== Daily Check ULTIMATE FIX ===");
    const result = scanMissingDataUltimateFix();
    
    if (Object.keys(result.groupedData).length > 0) {
      sendPersonalizedEmailUltimate(result.groupedData, result.employeeName);
      console.log(`Email sent to ${result.employeeName} - ${Object.keys(result.groupedData).length} companies`);
    } else {
      console.log("No missing data - no email sent");
    }
    
    PropertiesService.getScriptProperties().setProperty('lastRun', new Date().toString());
    
  } catch (error) {
    console.error("Error in dailyEmailCheck:", error);
    GmailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      "Error - Apps Script Email Reminder",
      `Error occurred: ${error.toString()}`
    );
  }
}

function manualCheck() {
  console.log("=== Manual Check ULTIMATE FIX ===");
  dailyEmailCheck();
}

/**
 * ULTIMATE FIX: Group by Company - Matrix Design
 */
function scanMissingDataUltimateFix() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getRange("A:BZ").getValues();

  const groupedByCompany = {};
  const currentMonth = new Date().getMonth() + 1;
  let employeeName = "";

  console.log(`=== SCAN START - Matrix Design ===`);
  console.log(`Current month: ${currentMonth} - Processing months <= ${currentMonth}`);
  console.log(`Numeric fields (0 = OK): ${NUMERIC_FIELDS.join(', ')}`);
  console.log(`Delay fields (only check months <= ${currentMonth - 2}): ${DELAY_FIELDS.join(', ')}`);

  // Loop through rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Check: employee code (A), company code (B), year 2025 (Q)
    if (row[0] && row[1] && row[16] && row[16].toString().includes('2025')) {

      // *** CHECK COLUMN BX - SKIP nếu có 'x' hoặc 'X' ***
      const bxIndex = multiColumnToIndex('BX'); // Tính chính xác index của BX
      const skipFlag = row[bxIndex];

      // DEBUG LOGGING for row 356
      if (i + 1 === 356) {
        console.log(`=== DEBUG ROW 356 ===`);
        console.log(`BX Index: ${bxIndex}`);
        console.log(`BX Value: "${skipFlag}"`);
        console.log(`BX Type: ${typeof skipFlag}`);
        console.log(`Row length: ${row.length}`);
        console.log(`Will skip: ${skipFlag && (skipFlag.toString().trim().toLowerCase() === 'x')}`);
      }

      if (skipFlag && (skipFlag.toString().trim().toLowerCase() === 'x')) {
        console.log(`Row ${i+1}: SKIPPED - marked as complete in column BX (value: "${skipFlag}")`);
        continue; // Bỏ qua hàng này hoàn toàn
      }
      
      const month = row[17] || 'Unknown'; // Column R (index 17)
      const monthNumber = parseInt(month) || 0;
      
      // *** FILTER: CHỈ XỬ LÝ THÁNG <= HIỆN TẠI ***
      if (monthNumber > currentMonth) {
        continue; // Skip tháng lớn hơn hiện tại
      }
      
      const companyName = row[20] || extractCompanyName(row[1]); // Column U or extract from B
      const rowNumber = i + 1;

      // Lấy tên nhân viên từ cột O (index 14)
      if (!employeeName && row[14]) {
        employeeName = extractFirstName(row[14].toString());
      }

      // Check each required field với ULTIMATE LOGIC
      const missingFields = [];
      Object.keys(REQUIRED_FIELDS).forEach(col => {
        const colIndex = columnLetterToIndex(col);
        const fieldName = REQUIRED_FIELDS[col];

        // SKIP certain fields for current month
        if (monthNumber === currentMonth && SKIP_IN_CURRENT_MONTH.includes(fieldName)) {
          return;
        }

        // SPECIAL LOGIC cho các fields có delay time
        if (DELAY_FIELDS.includes(fieldName) && monthNumber > (currentMonth - 2)) {
          return;
        }

        const cellValue = row[colIndex];
        let isMissing = false;

        // ULTIMATE LOGIC: check numeric vs text fields
        if (NUMERIC_FIELDS.includes(fieldName)) {
          isMissing = isEmptyValue(cellValue);
        } else {
          isMissing = isEmptyValue(cellValue);
        }

        if (isMissing) {
          missingFields.push(fieldName);
          console.log(`✓ MISSING: Row ${rowNumber}, month ${monthNumber}, "${fieldName}" for ${companyName}`);
        }
      });

      // Nếu có fields thiếu, add vào groupedByCompany
      if (missingFields.length > 0) {
        const companyKey = `${companyName}|${rowNumber}|${monthNumber}`;

        if (!groupedByCompany[companyKey]) {
          groupedByCompany[companyKey] = {
            companyName: companyName,
            rowNumber: rowNumber,
            month: monthNumber,
            missingFields: new Set()
          };
        }

        // Collect missing fields for this month
        missingFields.forEach(field => {
          groupedByCompany[companyKey].missingFields.add(field);
        });
      }
    }
  }

  console.log(`Found ${Object.keys(groupedByCompany).length} company-month entries with missing data`);

  return {
    groupedData: groupedByCompany,
    employeeName: employeeName || "bạn"
  };
}

/**
 * ULTIMATE HELPER: Check if value is truly empty (but not 0)
 */
function isEmptyValue(value) {
  // Null or undefined = empty
  if (value === null || value === undefined) {
    return true;
  }
  
  // Empty string or whitespace only = empty  
  if (typeof value === 'string' && value.trim() === '') {
    return true;
  }
  
  // Number 0 = NOT empty (this is the key fix)
  if (typeof value === 'number') {
    return false; // ANY number (including 0) is NOT empty
  }
  
  // String that represents a number (including "0") = NOT empty
  if (typeof value === 'string') {
    const numValue = parseFloat(value.trim());
    if (!isNaN(numValue)) {
      return false; // It's a valid number string, NOT empty
    }
  }
  
  // Everything else: check if truthy
  return !value;
}

/**
 * SEND EMAIL - Matrix Table Design
 */
function sendPersonalizedEmailUltimate(groupedData, employeeName) {
  if (Object.keys(groupedData).length === 0) return;

  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  // Sort: Month descending (recent first), then by missing count descending
  const sortedCompanyKeys = Object.keys(groupedData).sort((a, b) => {
    const compA = groupedData[a];
    const compB = groupedData[b];

    // Primary sort: Month (descending - tháng gần nhất trước)
    if (compB.month !== compA.month) {
      return compB.month - compA.month;
    }

    // Secondary sort: Missing fields count (descending - thiếu nhiều trước)
    const countA = compA.missingFields.size;
    const countB = compB.missingFields.size;
    if (countB !== countA) {
      return countB - countA;
    }

    // Tertiary sort: Company name (alphabetically)
    return compA.companyName.localeCompare(compB.companyName);
  });

  const totalEntries = sortedCompanyKeys.length;
  const fieldNames = Object.values(REQUIRED_FIELDS);

  // === HTML EMAIL - MATRIX TABLE ===
  let htmlContent = `
<div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; color: #000; max-width: 1100px; line-height: 1.5;">
<p>Kính gửi chị <strong>${employeeName}</strong>,</p>
<p>Trong sheet 'file nhap chc' có <strong>${totalEntries}</strong> mục đang thiếu thông tin, chị cập nhật vào em với nha.</p>
<p><a href="${sheetUrl}" style="color: #000; text-decoration: underline;">Google Sheet</a></p>

<table style="border-collapse: collapse; width: 100%; margin-top: 25px; font-size: 12px;">
  <thead>
    <tr style="border-bottom: 1px solid #000;">
      <th style="padding: 10px 8px; text-align: center; font-weight: bold; width: 60px;">Tháng</th>
      <th style="padding: 10px 8px; text-align: left; font-weight: bold;">Công ty</th>
`;

  // Header columns for each field
  fieldNames.forEach(fieldName => {
    const shortName = fieldName
      .replace('ngày ký hợp đồng', 'Ngày KýHĐ')
      .replace('doanh thu', 'Doanh Thu')
      .replace('số người khám', 'Số NK')
      .replace('mã hợp đồng', 'Mã HĐ')
      .replace('trạng thái ký', 'TT Ký')
      .replace('ngày hóa đơn', 'Ngày HĐơn')
      .replace('doanh thu thực hiện', 'DT TH')
      .replace('ngày bắt đầu khám', 'Ngày BĐ')
      .replace('ngày kết thúc khám', 'Ngày KT')
      .replace('tháng GNDT', 'GNDT');

    htmlContent += `      <th style="padding: 10px 6px; text-align: center; font-weight: bold; width: 70px;">${shortName}</th>\n`;
  });

  htmlContent += `    </tr>
  </thead>
  <tbody>
`;

  // Table rows
  sortedCompanyKeys.forEach(companyKey => {
    const company = groupedData[companyKey];

    htmlContent += `    <tr style="border-bottom: 1px solid #e0e0e0;">
      <td style="padding: 10px 8px; text-align: center; font-weight: bold;">${company.month}</td>
      <td style="padding: 10px 8px;">${company.companyName} <span style="color: #999; font-size: 11px;">(Hàng ${company.rowNumber})</span></td>
`;

    // Check mark for each field - BLACK COLOR
    fieldNames.forEach(fieldName => {
      const isMissing = company.missingFields.has(fieldName);
      htmlContent += `      <td style="padding: 10px 6px; text-align: center; color: #000; font-weight: bold;">${isMissing ? 'x' : ''}</td>\n`;
    });

    htmlContent += `    </tr>
`;
  });

  htmlContent += `  </tbody>
</table>

<p style="margin-top: 25px; color: #666;">Trân trọng</p>

</div>`;

  // === PLAIN TEXT VERSION ===
  let textContent = `Kính gửi chị ${employeeName},\n\n`;
  textContent += `Trong sheet 'file nhap chc' có ${totalEntries} mục đang thiếu thông tin, chị cập nhật vào em với nha.\n\n`;
  textContent += `Link: ${sheetUrl}\n\n`;

  sortedCompanyKeys.forEach(companyKey => {
    const company = groupedData[companyKey];
    textContent += `Tháng ${company.month} - ${company.companyName} (Hàng ${company.rowNumber}): ${Array.from(company.missingFields).join(', ')}\n`;
  });

  // === SEND EMAIL ===
  try {
    GmailApp.sendEmail(
      EMPLOYEE_EMAIL,
      EMAIL_SUBJECT,
      textContent,
      {
        htmlBody: htmlContent,
        name: "Data System"
      }
    );

    console.log(`Email sent successfully to ${EMPLOYEE_EMAIL}`);
    console.log(`Recipient: ${employeeName}`);
    console.log(`Total entries with missing data: ${totalEntries}`);

  } catch (error) {
    console.error("Email sending error:", error);
    throw error;
  }
}

// =============== HELPER FUNCTIONS ===============

/**
 * Extract company name from "CODE - COMPANY NAME" format
 */
function extractCompanyName(fullString) {
  if (!fullString) return "Unknown Company";
  
  const parts = fullString.split(' - ');
  return parts.length > 1 ? parts.slice(1).join(' - ') : fullString;
}

/**
 * Extract first name from full name (lấy tên cuối)
 */
function extractFirstName(fullName) {
  if (!fullName) return "bạn";
  
  const parts = fullName.trim().split(' ');
  return parts[parts.length - 1] || "bạn";
}

/**
 * Convert column letter to index (A=0, B=1, ...)
 */
function columnLetterToIndex(letter) {
  return letter.charCodeAt(0) - 65;
}

/**
 * Convert multi-letter column to index (A=0, B=1, ..., AA=26, AB=27, ..., BX=75)
 */
function multiColumnToIndex(columnName) {
  let result = 0;
  for (let i = 0; i < columnName.length; i++) {
    result = result * 26 + (columnName.charCodeAt(i) - 64);
  }
  return result - 1; // Convert to 0-based index
}

/**
 * Convert index back to column name (0=A, 1=B, ..., 75=BX)
 */
function indexToColumn(index) {
  let result = '';
  let num = index + 1; // Convert to 1-based
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}

// =============== SETUP FUNCTIONS ===============

function setupDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyEmailCheck') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('dailyEmailCheck')
    .timeBased()
    .everyDays(1)
    .atHour(14)
    .create();
  
  console.log("Daily trigger setup at 2 PM");
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  console.log("All triggers deleted");
}

function testConfiguration() {
  console.log("=== Test Configuration ULTIMATE ===");
  console.log(`Email: ${EMPLOYEE_EMAIL}`);
  console.log(`Sheet: ${SHEET_NAME}`);
  console.log(`Priority fields: ${PRIORITY_FIELDS.join(', ')}`);
  console.log(`Numeric fields (0 = NOT missing): ${NUMERIC_FIELDS.join(', ')}`);
  console.log(`Delay fields: ${DELAY_FIELDS.join(', ')}`);

  const currentMonth = new Date().getMonth() + 1;
  console.log(`Current month: ${currentMonth}`);
  console.log(`Delay fields only check months <= ${currentMonth - 2}`);

  // Test column index calculation
  console.log("=== Testing Column Index Calculation ===");
  console.log(`BX column index: ${multiColumnToIndex('BX')}`); // Should be correct
  console.log(`BU column index: ${multiColumnToIndex('BU')}`);
  console.log(`BV column index: ${multiColumnToIndex('BV')}`);
  console.log(`BW column index: ${multiColumnToIndex('BW')}`);

  // Test isEmptyValue function
  console.log("=== Testing isEmptyValue function ===");
  console.log(`isEmptyValue(0): ${isEmptyValue(0)}`); // Should be false
  console.log(`isEmptyValue("0"): ${isEmptyValue("0")}`); // Should be false
  console.log(`isEmptyValue(""): ${isEmptyValue("")}`); // Should be true
  console.log(`isEmptyValue(" "): ${isEmptyValue(" ")}`); // Should be true
  console.log(`isEmptyValue(null): ${isEmptyValue(null)}`); // Should be true
  console.log(`isEmptyValue(undefined): ${isEmptyValue(undefined)}`); // Should be true
  console.log(`isEmptyValue(123): ${isEmptyValue(123)}`); // Should be false
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const testData = sheet.getRange("A1:V5").getValues();
    
    console.log("Sheet readable");
    if (testData[1]) {
      console.log("Sample row:", testData[1]);
      console.log("Column G (số người khám):", testData[1][6], "Type:", typeof testData[1][6]);
    }
    
    // Test ultimate scan logic
    console.log("=== Testing ULTIMATE scan logic ===");
    const result = scanMissingDataUltimateFix();
    console.log("Employee name extracted:", result.employeeName);
    console.log("Grouped result keys:", Object.keys(result.groupedData));
    
  } catch (error) {
    console.error("Test error:", error);
  }
}

function checkLastRun() {
  const lastRun = PropertiesService.getScriptProperties().getProperty('lastRun');
  console.log(`Last run: ${lastRun || 'Never run'}`);
}


// =============== TUESDAY & FRIDAY TRIGGERS ===============

function setupTuesdayFridayTriggers() {
  // Xóa hết trigger cũ trước
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyEmailCheck') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Trigger 1: Thứ 3 (TUESDAY)
  ScriptApp.newTrigger('dailyEmailCheck')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(14)
    .create();

  // Trigger 2: Thứ 6 (FRIDAY)
  ScriptApp.newTrigger('dailyEmailCheck')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(14)
    .create();

  console.log("Triggers setup: Tuesday & Friday at 2 PM (14:00)");
  console.log("Next runs will be on Tuesdays and Fridays at 14:00");
}

function checkActiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  console.log(`Total active triggers: ${triggers.length}`);
  
  triggers.forEach((trigger, index) => {
    console.log(`Trigger ${index + 1}:`);
    console.log(`- Function: ${trigger.getHandlerFunction()}`);
    console.log(`- Type: ${trigger.getEventType()}`);
    
    if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      const source = trigger.getTriggerSource();
      console.log(`- Source: ${source}`);
    }
  });
}