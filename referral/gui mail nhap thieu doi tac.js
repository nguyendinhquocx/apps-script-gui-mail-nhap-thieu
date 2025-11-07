/**
 * GOOGLE SHEETS EMAIL REMINDER - TEAM REFERRAL
 * Tự động gửi email nhắc nhở nhân viên bổ sung thông tin đối tác
 * FILTER: Chỉ tháng <= hiện tại, PERSONALIZED
 */

// =============== CONFIGURATION ===============
const EMPLOYEE_EMAIL = "quoc.nguyen3@hoanmy.com"; // THAY EMAIL NÀY CHO TỪNG NHÂN VIÊN
const SHEET_NAME = "thong tin doi tac";
const EMAIL_SUBJECT = "Bổ sung thông tin đối tác";

// Map columns - 6 fields cần check
const REQUIRED_FIELDS = {
  'C': 'nơi công tác',
  'D': 'mã chuyên khoa',
  'F': 'loại hình hợp tác',
  'G': 'ngày hợp tác',
  'I': 'ID đối tác',
  'S': 'hiệu lực HĐ'
};

// =============== MAIN FUNCTIONS ===============

function dailyEmailCheck() {
  try {
    console.log("=== Daily Check - Team Referral ===");
    const result = scanMissingData();

    if (Object.keys(result.groupedData).length > 0) {
      sendMinimalEmail(result.groupedData, result.employeeName);
      console.log(`Email sent to ${result.employeeName} - ${Object.keys(result.groupedData).length} partners`);
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
  console.log("=== Manual Check - Team Referral ===");
  dailyEmailCheck();
}

/**
 * Scan missing data in sheet - GROUP BY PARTNER
 */
function scanMissingData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getRange("A:Z").getValues();

  const groupedByPartner = {};
  const currentMonth = new Date().getMonth() + 1;
  const currentYear = new Date().getFullYear();
  let employeeName = "";

  console.log(`=== SCAN START ===`);
  console.log(`Current: ${currentMonth}/${currentYear}`);

  // Loop through rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Check: có tên đối tác (A) và năm hợp tác (N)
    if (row[0] && row[13]) {
      const yearValue = row[13];
      const monthValue = row[14];
      const year = parseInt(yearValue) || 0;
      const month = parseInt(monthValue) || 0;

      // FILTER: Chỉ xử lý năm hiện tại và tháng <= hiện tại
      if (year !== currentYear || month > currentMonth) {
        continue;
      }

      const partnerName = row[0]; // Column A
      const rowNumber = i + 1;

      // Lấy tên nhân viên từ cột J (index 9)
      if (!employeeName && row[9]) {
        employeeName = extractFirstName(row[9].toString());
      }

      // Check each required field
      const missingFields = [];
      Object.keys(REQUIRED_FIELDS).forEach(col => {
        const colIndex = columnLetterToIndex(col);
        const fieldName = REQUIRED_FIELDS[col];
        const cellValue = row[colIndex];

        if (isEmptyValue(cellValue)) {
          missingFields.push({
            fieldName: fieldName,
            month: month
          });
          console.log(`✓ MISSING: Row ${rowNumber}, "${fieldName}" (month ${month}) for ${partnerName}`);
        }
      });

      // Nếu có fields thiếu, add vào groupedByPartner
      if (missingFields.length > 0) {
        const partnerKey = `${partnerName}|${rowNumber}`;

        if (!groupedByPartner[partnerKey]) {
          groupedByPartner[partnerKey] = {
            partnerName: partnerName,
            rowNumber: rowNumber,
            missingByMonth: {}
          };
        }

        // Group missing fields by month
        missingFields.forEach(item => {
          if (!groupedByPartner[partnerKey].missingByMonth[item.month]) {
            groupedByPartner[partnerKey].missingByMonth[item.month] = [];
          }
          groupedByPartner[partnerKey].missingByMonth[item.month].push(item.fieldName);
        });
      }
    }
  }

  console.log(`Found ${Object.keys(groupedByPartner).length} partners with missing data`);

  return {
    groupedData: groupedByPartner,
    employeeName: employeeName || "bạn"
  };
}

/**
 * Check if value is empty
 */
function isEmptyValue(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string' && value.trim() === '') return true;
  if (typeof value === 'number') return false; // Any number is NOT empty
  return !value;
}

/**
 * SEND MINIMAL EMAIL - Matrix Table Design
 */
function sendMinimalEmail(groupedData, employeeName) {
  if (Object.keys(groupedData).length === 0) return;

  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  // Sort partners by number of missing fields (descending), then alphabetically
  const sortedPartnerKeys = Object.keys(groupedData).sort((a, b) => {
    const partnerA = groupedData[a];
    const partnerB = groupedData[b];

    // Count missing fields for each partner
    const countA = new Set();
    Object.values(partnerA.missingByMonth).forEach(monthFields => {
      monthFields.forEach(field => countA.add(field));
    });

    const countB = new Set();
    Object.values(partnerB.missingByMonth).forEach(monthFields => {
      monthFields.forEach(field => countB.add(field));
    });

    // Sort by count descending, then by name
    if (countB.size !== countA.size) {
      return countB.size - countA.size; // More missing = higher priority
    }
    return partnerA.partnerName.localeCompare(partnerB.partnerName);
  });

  // Field names in order
  const fieldNames = Object.values(REQUIRED_FIELDS);

  // Count total partners
  const totalPartners = sortedPartnerKeys.length;

  // === HTML EMAIL - MATRIX TABLE ===
  let htmlContent = `
<div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; color: #000; max-width: 900px; line-height: 1.5;">
<p>Kính gửi chị <strong>${employeeName}</strong>,</p>
<p>Trong sheet 'thong tin doi tac' có <strong>${totalPartners}</strong> đối tác đang thiếu thông tin, chị cập nhật vào em với nha.</p>
<p><a href="${sheetUrl}" style="color: #000; text-decoration: underline;">Google Sheet</a></p>

<table style="border-collapse: collapse; width: 100%; margin-top: 25px; font-size: 13px;">
  <thead>
    <tr style="border-bottom: 1px solid #000;">
      <th style="padding: 10px 8px; text-align: left; font-weight: light;">Đối tác</th>
`;

  // Header columns for each field
  fieldNames.forEach(fieldName => {
    // Shorten field names for header
    const shortName = fieldName
      .replace('nơi công tác', 'Nơi CT')
      .replace('mã chuyên khoa', 'Chuyên Khoa')
      .replace('loại hình hợp tác', 'Loại HT')
      .replace('ngày hợp tác', 'Ngày HT')
      .replace('ID đối tác', 'ID')
      .replace('hiệu lực HĐ', 'Hiệu Lực');

    htmlContent += `      <th style="padding: 10px 8px; text-align: center; font-weight: bold; width: 80px;">${shortName}</th>\n`;
  });

  htmlContent += `    </tr>
  </thead>
  <tbody>
`;

  // Table rows - each partner
  sortedPartnerKeys.forEach(partnerKey => {
    const partner = groupedData[partnerKey];

    // Collect all missing fields (across all months)
    const allMissingFields = new Set();
    Object.values(partner.missingByMonth).forEach(monthFields => {
      monthFields.forEach(field => allMissingFields.add(field));
    });

    htmlContent += `    <tr style="border-bottom: 1px solid #e0e0e0;">
      <td style="padding: 10px 8px;">${partner.partnerName} <span style="color: #999; font-size: 12px; ">(Hàng ${partner.rowNumber})</span></td>
`;

    // Check mark for each field - BLACK COLOR
    fieldNames.forEach(fieldName => {
      const isMissing = allMissingFields.has(fieldName);
      htmlContent += `      <td style="padding: 10px 8px; text-align: center; color: #000; font-weight: light;">${isMissing ? 'x' : ''}</td>\n`;
    });

    htmlContent += `    </tr>
`;
  });

  htmlContent += `  </tbody>
</table>

<p style="margin-top: 25px;  color: #666;">Trân trọng</p>

</div>`;

  // === PLAIN TEXT VERSION ===
  let textContent = `Chị ${employeeName},\n\n`;
  textContent += `Một số thông tin đối tác đang thiếu, chị bổ sung giúp em.\n\n`;
  textContent += `Link: ${sheetUrl}\n\n`;

  // Simple list for plain text
  sortedPartnerKeys.forEach(partnerKey => {
    const partner = groupedData[partnerKey];

    const allMissingFields = new Set();
    Object.values(partner.missingByMonth).forEach(monthFields => {
      monthFields.forEach(field => allMissingFields.add(field));
    });

    textContent += `${partner.partnerName} (Dòng ${partner.rowNumber}): ${Array.from(allMissingFields).join(', ')}\n`;
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
    console.log(`Total partners with missing data: ${sortedPartnerKeys.length}`);

  } catch (error) {
    console.error("Email sending error:", error);
    throw error;
  }
}

// =============== HELPER FUNCTIONS ===============

/**
 * Extract first name from full name
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
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result - 1;
}

// =============== SETUP FUNCTIONS ===============

function setupTuesdayFridayTriggers() {
  // Xóa trigger cũ
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyEmailCheck') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Trigger: Thứ 3
  ScriptApp.newTrigger('dailyEmailCheck')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(14)
    .create();

  // Trigger: Thứ 6
  ScriptApp.newTrigger('dailyEmailCheck')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(14)
    .create();

  console.log("Triggers setup: Tuesday & Friday at 2 PM");
}

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
  console.log("=== Test Configuration ===");
  console.log(`Email: ${EMPLOYEE_EMAIL}`);
  console.log(`Sheet: ${SHEET_NAME}`);
  console.log(`Fields to check: ${Object.values(REQUIRED_FIELDS).join(', ')}`);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const testData = sheet.getRange("A1:Z5").getValues();

    console.log("Sheet readable");
    if (testData[1]) {
      console.log("Sample row columns N,O (year,month):", testData[1][13], testData[1][14]);
    }

    const result = scanMissingData();
    console.log("Employee name:", result.employeeName);
    console.log("Partners with missing data:", Object.keys(result.groupedData).length);

  } catch (error) {
    console.error("Test error:", error);
  }
}

function checkLastRun() {
  const lastRun = PropertiesService.getScriptProperties().getProperty('lastRun');
  console.log(`Last run: ${lastRun || 'Never run'}`);
}

function checkActiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  console.log(`Total active triggers: ${triggers.length}`);

  triggers.forEach((trigger, index) => {
    console.log(`Trigger ${index + 1}:`);
    console.log(`- Function: ${trigger.getHandlerFunction()}`);
    console.log(`- Type: ${trigger.getEventType()}`);
  });
}
