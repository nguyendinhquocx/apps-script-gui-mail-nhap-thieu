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
      console.log(`Email sent to ${result.employeeName} - ${Object.keys(result.groupedData).length} months`);
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
 * ULTIMATE FIX: Triệt để check numeric vs text fields + delay fields
 */
function scanMissingDataUltimateFix() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  // Expand range to BZ to ensure BX column is included
  const data = sheet.getRange("A:BZ").getValues();
  
  const groupedByMonth = {};
  const currentMonth = new Date().getMonth() + 1; // Current month (1-12)
  let employeeName = ""; // Lấy tên từ cột O
  
  console.log(`=== ULTIMATE FIX SCAN ===`);
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
      
      // Lấy tên nhân viên từ cột O (index 14) - chỉ lấy lần đầu
      if (!employeeName && row[14]) {
        employeeName = extractFirstName(row[14].toString());
      }
      
      // Check each required field với ULTIMATE LOGIC
      Object.keys(REQUIRED_FIELDS).forEach(col => {
        const colIndex = columnLetterToIndex(col);
        const fieldName = REQUIRED_FIELDS[col];
        
        // SKIP certain fields for current month
        if (monthNumber === currentMonth && SKIP_IN_CURRENT_MONTH.includes(fieldName)) {
          return; // Skip this field for current month
        }
        
        // *** SPECIAL LOGIC cho các fields có delay time ***
        if (DELAY_FIELDS.includes(fieldName) && monthNumber > (currentMonth - 2)) {
          console.log(`Skipping "${fieldName}" for month ${monthNumber} (> ${currentMonth - 2}) - delay field`);
          return; // Skip delay fields cho tháng hiện tại và tháng trước
        }
        
        const cellValue = row[colIndex];
        let isMissing = false;
        
        // *** ULTIMATE LOGIC: TRIỆT ĐỂ check numeric fields ***
        if (NUMERIC_FIELDS.includes(fieldName)) {
          // TRIỆT ĐỂ: Với numeric fields, chỉ missing nếu null/undefined/empty string
          isMissing = isEmptyValue(cellValue);
          
          // DEBUG LOG
          console.log(`Row ${i+1}, Field "${fieldName}", Value: "${cellValue}", Type: ${typeof cellValue}, Missing: ${isMissing}`);
          
        } else {
          // Với TEXT fields: tính cả trống và whitespace
          isMissing = isEmptyValue(cellValue);
        }
        
        // If field is missing theo logic trên
        if (isMissing) {
          
          // Initialize structure if not exists
          if (!groupedByMonth[month]) {
            groupedByMonth[month] = {};
          }
          if (!groupedByMonth[month][fieldName]) {
            groupedByMonth[month][fieldName] = [];
          }
          
          // Add company to this group
          groupedByMonth[month][fieldName].push({
            rowNumber: i + 1,
            companyName: companyName
          });
          
          console.log(`✓ MISSING: Row ${i+1}, "${fieldName}" = "${cellValue}"`);
        }
      });
    }
  }
  
  console.log(`Found missing data in ${Object.keys(groupedByMonth).length} months`);
  console.log("Months with data:", Object.keys(groupedByMonth));
  console.log("Employee name:", employeeName);
  
  return {
    groupedData: groupedByMonth,
    employeeName: employeeName || "bạn" // fallback
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
 * SEND PERSONALIZED EMAIL - ULTIMATE VERSION
 */
function sendPersonalizedEmailUltimate(groupedData, employeeName) {
  if (Object.keys(groupedData).length === 0) return;
  
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  // === HTML EMAIL CONTENT ===
  let htmlContent = `
<div style="font-family: Arial, sans-serif; color: #000; max-width: 800px; line-height: 1.4;">
<p style="margin-bottom: 5px;">Kính gửi chị <strong>${employeeName}</strong>,</p>
<p style="margin-bottom: 15px;">Các trường thông tin trong file nhập đang bị thiếu một số hạng mục ở các tháng, có gì chị cập nhật vào em với nha.</p>
<p style="margin-bottom: 20px;"><strong><a href="${sheetUrl}" style="color: #000; text-decoration: underline;">Mở Google Sheet</a></strong></p>
`;

  // Sort months DESCENDING (larger months first) 
  const sortedMonths = Object.keys(groupedData).sort((a, b) => {
    const monthA = parseInt(a) || 0;
    const monthB = parseInt(b) || 0;
    return monthB - monthA; // Descending order
  });
  
  console.log(`Month order (high to low): ${sortedMonths.join(', ')}`);
  
  sortedMonths.forEach(month => {
    htmlContent += `
<div style="margin-bottom: 25px; border: 2px solid #000; padding: 15px;">
  <h4 style="margin: 0 0 15px 0; background: #e0e0e0; padding: 10px; border-bottom: 2px solid #000; font-weight: bold;">
    THÁNG ${month}
  </h4>
`;

    // Sort fields by priority (priority fields first)
    const monthFields = Object.keys(groupedData[month]);
    const priorityFieldsInMonth = monthFields.filter(field => PRIORITY_FIELDS.includes(field)).sort();
    const otherFieldsInMonth = monthFields.filter(field => !PRIORITY_FIELDS.includes(field)).sort();
    const sortedFields = [...priorityFieldsInMonth, ...otherFieldsInMonth];

    sortedFields.forEach(fieldType => {
      const companies = groupedData[month][fieldType];
      const isPriority = PRIORITY_FIELDS.includes(fieldType);
      
      htmlContent += `
  <div style="margin-bottom: 15px;">
    <h5 style="margin: 8px 0 5px 0; color: ${isPriority ? '#c00' : '#d00'}; text-transform: uppercase; font-weight: ${isPriority ? 'bold' : 'normal'};">
      ${isPriority ? '⚠️ ' : ''}${fieldType} (${companies.length} công ty)
    </h5>
    <table border="1" style="border-collapse: collapse; width: 100%; margin-bottom: 10px;">
      <tr style="background: #f5f5f5;">
        <th style="padding: 6px; text-align: left; width: 80px; border: 1px solid #999;">Dòng</th>
        <th style="padding: 6px; text-align: left; border: 1px solid #999;">Công ty</th>
      </tr>
`;

      companies.forEach(company => {
        htmlContent += `
      <tr>
        <td style="padding: 6px; border: 1px solid #ccc;">${company.rowNumber}</td>
        <td style="padding: 6px; border: 1px solid #ccc;">${company.companyName}</td>
      </tr>
`;
      });
      
      htmlContent += `    </table>
  </div>
`;
    });
    
    htmlContent += `</div>
`;
  });
  
  htmlContent += `
<p style="margin-top: 20px; text-align: left;"><em>Trân trọng</em></p>
</div>`;

  // === PLAIN TEXT VERSION ===
  let textContent = `Kính gửi chị ${employeeName},\n`;
  textContent += `Các trường thông tin đang bị thiếu một số hạng mục, có gì chị cập nhật giúp em nha.\n\n`;
  textContent += `Link: ${sheetUrl}\n\n`;
  
  sortedMonths.forEach(month => {
    textContent += `=== THÁNG ${month} ===\n`;
    
    // Same sorting logic
    const monthFields = Object.keys(groupedData[month]);
    const priorityFieldsInMonth = monthFields.filter(field => PRIORITY_FIELDS.includes(field)).sort();
    const otherFieldsInMonth = monthFields.filter(field => !PRIORITY_FIELDS.includes(field)).sort();
    const sortedFields = [...priorityFieldsInMonth, ...otherFieldsInMonth];
    
    sortedFields.forEach(fieldType => {
      const companies = groupedData[month][fieldType];
      const isPriority = PRIORITY_FIELDS.includes(fieldType);
      
      textContent += `\n${isPriority ? '[PRIORITY] ' : ''}${fieldType.toUpperCase()} (${companies.length} công ty):\n`;
      
      companies.forEach(company => {
        textContent += `  • Dòng ${company.rowNumber}: ${company.companyName}\n`;
      });
    });
    
    textContent += "\n";
  });
  
  textContent += "\nTrân trọng";

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
    
    // Count total missing items
    let totalMissing = 0;
    Object.values(groupedData).forEach(monthData => {
      Object.values(monthData).forEach(fieldArray => {
        totalMissing += fieldArray.length;
      });
    });
    
    console.log(`Email sent successfully to ${EMPLOYEE_EMAIL}`);
    console.log(`Recipient: ${employeeName}`);
    console.log(`Total: ${totalMissing} missing items across ${sortedMonths.length} months`);
    
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
  
  console.log("Triggers setup: Tuesday & Friday at 2 PM");
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