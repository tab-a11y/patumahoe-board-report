// ============================================================
// [SCHOOL NAME] — BOARD REPORT GENERATOR
// Google Apps Script — Gemini Version (outputs .docx via service)
//
// SETUP INSTRUCTIONS:
// 1. In your Google Sheet go to Extensions → Apps Script
// 2. Paste this entire script replacing any existing code
// 3. Go to Project Settings (gear icon) → Script Properties
// 4. Add: GEMINI_API_KEY   = your key from aistudio.google.com
// 5. Add: DOCX_SERVICE_URL = https://patumahoe-board-report.onrender.com
// 6. Save, run onOpen once to grant permissions
// 7. Add a button in your Sheet and assign generateBoardReport
// ============================================================

// ── CONFIG ───────────────────────────────────────────────────────────────
const SOURCE_FOLDER_ID = 'YOUR_FOLDER_ID_HERE';   // ← CHANGE THIS
const SCHOOL_NAME      = '[School Name]';          // ← CHANGE THIS
const SCHOOL_VISION    = '[School vision here]';   // ← CHANGE THIS
const GEMINI_API_URL   = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent';
const MAX_TOKENS       = 4000;

// ── SYSTEM PROMPT ────────────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are the Board Report Curator for [School Name].  // ← CHANGE THIS

Your role is to review all source images provided and produce a clear, professional
Principal Report to the Board of Trustees.

The report must be accurate, concise, board-appropriate, and written in a calm,
factual principal voice. Do not over-explain. Do not invent information. Do not
produce unsupported commentary.

SCHOOL VISION: [School vision here]  // ← CHANGE THIS

REPORT STRUCTURE — produce sections in this order:
1. Title / Header (School name, month/year, board meeting)
2. Enrolments & Roll (table + commentary)
3. Staffing Notes (numbered)
4. Health & Safety (table + Comments on Injuries)
5. Attendance (table + commentary)
6. Property (numbered updates)
7. Assurances to Board (table)
8. Policy Review (table)
9. Progress Against Annual Plan Goals (all goal areas)
10. Connection & Community
11. Report prepared by / Date sign-off

FORMAT RULES — follow exactly:
- Major section headings: use ## (e.g. ## ENROLMENTS & ROLL)
- Sub-headings: use ### (e.g. ### Staffing Notes)
- Tables: use markdown pipe tables with a header row and separator row
- Bullet points: use - (dash space)
- Numbered items: use 1. 2. 3.
- Placeholders: use [square brackets] for missing data

SECTION RULES:

ROLL TABLE: Columns — Year Level | Enrolments | Withdrawals | Total | Out of Zone
Rows: Year 0–6, TOTAL row, OOZ %.
Commentary: explain roll movement, notable year levels, OOZ trends. Do not just repeat numbers.

STAFFING NOTES: Numbered. Cover staffing pressure, banking deficit, new/start-up classes,
roll projections, vacancies, appointments, leave, risks.
Use [Staffing update to be provided] if no staffing data is present.

HEALTH & SAFETY TABLE: Columns — Incident Type | Staff | Students | Other | Total
Rows: Minor injuries (on-site only) | Injuries requiring further investigation | Serious harm (WorkSafe).
Comments: note incidents needing follow-up, patterns, WorkSafe notifications.
State clearly if no serious harm injuries occurred.

ATTENDANCE TABLE: Columns — Attendance Band | Number of Students | % of Students
Bands: 90.1–100 | 80.1–90 | 70.1–80 | 0–70
Also include GOOD / WORRYING / CONCERNING / SERIOUS CONCERN percentages if visible.
Target statement: "Target: 90% of students attending 80% of the time or more."
Commentary: whether school is on track, concern groups, actions under attendance plan.

PROPERTY: Numbered updates. Each item: status | risk or impact | next step.
Use [Property update to be provided] if no property data is present.

ASSURANCES TO BOARD — always include these THREE every-term assurances:
1. Curriculum and Student Achievement Policy
2. Risk Management
3. Emergency Management

Then add TERM-SPECIFIC assurances:
- Term 1: School Planning and Reporting; Inclusive School Culture; Maori Educational Achievement; Learning Support; Health Education; Health Safety and Welfare Policy; Worker Engagement; Digital Technology and Online Safety; Safety Checking and Police Vetting.
- Term 2: Teaching Staff registration; Performance Management; Staff Conduct; Appointment Policy; Employment Policy and EEO; Child Protection; Safety Checking and Police Vetting; Cellphones and Personal Digital Devices.
- Term 3: Student Attendance; Reporting about Student Progress and Achievement; Bullying and Harassment; Behaviour Management; Computer Security and Cybersecurity; Concerns and Complaints Policy; Finance and Asset Management.
- Term 4: School Records Retention; Food and Nutrition; Opening and Closing the School; Income; Gifts; Protected Disclosure; Bullying and Harassment; Behaviour Management; Concerns and Complaints; Visitors; Finance and Asset Management.

Write each assurance beginning with: "The principal assures the board that..."

POLICY REVIEW TABLE: Columns — # | Policy Name | Review Type | Notes / Action Required
2026 Term 1: Alcohol Drugs and Harmful Substances; Sun Protection; Digital Technology and Online Safety; Cellphones and Personal Digital Devices.
2026 Term 2: Daily School Bus; School Swimming Pool / Swimming Off Site; Education Outside the Classroom (EOTC); EOTC Governance Roles and Responsibilities; EOTC Risk Assessment and Management.
2026 Term 3: School Community Engagement Policy; Inclusive School Culture; Enrolment; Student Attendance; Student Uniform; Concerns and Complaints Policy.
2026 Term 4: Curriculum and Student Achievement Policy; Reporting about Student Progress and Achievement; Learning Support; Maori Educational Achievement; Health Education; Privacy Policy.

ANNUAL PLAN PROGRESS — use these goal areas and actions:

## [Goal Area 1]  // ← REPLACE ALL OF THIS with the new school's annual plan
- [Action] | [Term timing]

MISSING DATA RULE: If data is missing use:
[Roll data not provided] | [Health and Safety data not provided] |
[Attendance data not provided] | [Staffing update to be provided] |
[Property update to be provided] | [Annual plan progress not provided]

WRITING STYLE:
- Plain New Zealand English
- Factual, concise, calm, board-appropriate
- No emotional language, vague commentary, or jargon
- Commentary: 3-6 bullet points, interpret data not just repeat it`;

// ── USER PROMPT ───────────────────────────────────────────────────────────
function buildUserPrompt(month, year, term) {
  return `Please produce the Principal Report to the Board of Trustees for ${month} ${year} (${term}).

The source images attached contain the school data for this reporting period.
Extract all data carefully and accurately. Do not guess figures that are unclear.

Where data sections are missing, use the appropriate placeholder rather than inventing information.

After the report, add a brief MISSING DATA NOTE listing any sections with placeholders so the principal knows what to add before the board meeting.`;
}

// ── MAIN FUNCTION ─────────────────────────────────────────────────────────
function generateBoardReport() {

  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const month = sheet.getRange('B2').getValue() || getCurrentMonth();
  const year  = sheet.getRange('B3').getValue() || new Date().getFullYear();
  const term  = sheet.getRange('B4').getValue() || getCurrentTerm();

  const confirm = ui.alert(
    'Generate Board Report',
    `Generate report for ${month} ${year} (${term})?\n\nThis will read all image files from the source folder.`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  ui.alert('Running... This will take 30–60 seconds. Click OK and wait for the next message.');

  try {

    // ── Step 1: Get API keys ─────────────────────────────────────────
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      ui.alert('Error: GEMINI_API_KEY not found in Script Properties.\n\nGet your key from aistudio.google.com and add it to Script Properties.');
      return;
    }

    const docxServiceUrl = PropertiesService.getScriptProperties().getProperty('DOCX_SERVICE_URL');
    if (!docxServiceUrl) {
      ui.alert('Error: DOCX_SERVICE_URL not found in Script Properties.\n\nAdd: https://patumahoe-board-report.onrender.com');
      return;
    }

    // ── Step 2: Read images from Drive folder ────────────────────────
    const folder    = DriveApp.getFolderById(SOURCE_FOLDER_ID);
    const files     = folder.getFiles();
    const parts     = [];   // Gemini uses "parts" not "content blocks"
    const fileNames = [];

    while (files.hasNext()) {
      const file     = files.next();
      const fileName = file.getName().toLowerCase();

      let mimeType;
      if      (fileName.endsWith('.png'))  { mimeType = 'image/png'; }
      else if (fileName.endsWith('.jpg') ||
               fileName.endsWith('.jpeg')) { mimeType = 'image/jpeg'; }
      else if (fileName.endsWith('.gif'))  { mimeType = 'image/gif'; }
      else if (fileName.endsWith('.webp')) { mimeType = 'image/webp'; }
      else {
        Logger.log('Skipping unsupported file type: ' + file.getName());
        continue;
      }

      const base64 = Utilities.base64Encode(file.getBlob().getBytes());

      // Gemini image format: inlineData with mimeType and data
      parts.push({
        inlineData: {
          mimeType: mimeType,
          data: base64
        }
      });

      fileNames.push(file.getName());
      Logger.log('Loaded image: ' + file.getName() + ' as ' + mimeType);
    }

    if (parts.length === 0) {
      ui.alert('No image files found in the source folder.\n\nPlease upload PNG or JPG screenshots to the Drive folder and try again.');
      return;
    }

    // Add the text prompt as the final part
    parts.push({ text: buildUserPrompt(month, year, term) });

    Logger.log('Images loaded: ' + fileNames.join(', '));

    // ── Step 3: Call Gemini API ──────────────────────────────────────
    const payload = {
      system_instruction: {
        parts: [{ text: SYSTEM_PROMPT }]
      },
      contents: [{
        role: 'user',
        parts: parts
      }],
      generationConfig: {
        maxOutputTokens: MAX_TOKENS
      }
    };

    const geminiResponse = UrlFetchApp.fetch(GEMINI_API_URL + '?key=' + apiKey, {
      method: 'post',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (geminiResponse.getResponseCode() !== 200) {
      Logger.log('Gemini API Error: ' + geminiResponse.getContentText());
      ui.alert('Gemini API Error ' + geminiResponse.getResponseCode() + ':\n' + geminiResponse.getContentText().substring(0, 300));
      return;
    }

    // Gemini response path: candidates[0].content.parts[0].text
    const geminiResult = JSON.parse(geminiResponse.getContentText());
    const reportText   = geminiResult.candidates[0].content.parts[0].text;
    Logger.log('Gemini response received, length: ' + reportText.length);

    // ── Step 4: Send to docx service ────────────────────────────────
    Logger.log('Sending to docx service: ' + docxServiceUrl);

    const docxResponse = UrlFetchApp.fetch(docxServiceUrl + '/generate', {
      method: 'post',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({ reportText, month, year, term, schoolName: SCHOOL_NAME, schoolVision: SCHOOL_VISION }),
      muteHttpExceptions: true
    });

    if (docxResponse.getResponseCode() !== 200) {
      Logger.log('Docx service error: ' + docxResponse.getContentText());
      ui.alert('Docx service error ' + docxResponse.getResponseCode() + ':\n' + docxResponse.getContentText().substring(0, 200));
      return;
    }

    // ── Step 5: Save .docx to Drive folder ──────────────────────────
    const docxBlob = docxResponse.getBlob()
      .setName('Board Report — ' + month + ' ' + year + '.docx')
      .setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');

    const docxFile = folder.createFile(docxBlob);
    Logger.log('Docx saved: ' + docxFile.getUrl());

    // ── Step 6: Log the run ──────────────────────────────────────────
    logRun(sheet, month, year, term, fileNames.length, docxFile.getUrl());

    // ── Done ─────────────────────────────────────────────────────────
    ui.alert(
      '✅ Report Generated!',
      'Board Report for ' + month + ' ' + year + ' is ready as a Word document.\n\n' +
      'Files used: ' + fileNames.length + ' images\n\n' +
      'Find the .docx file in your Drive folder.',
      ui.ButtonSet.OK
    );

  } catch (err) {
    Logger.log('Error: ' + err.toString());
    ui.alert('Something went wrong:\n\n' + err.toString() + '\n\nCheck Apps Script Logs for details.');
  }
}

// ── HELPER: Log each run ──────────────────────────────────────────────────
function logRun(sheet, month, year, term, fileCount, docUrl) {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet   = ss.getSheetByName('Run Log');
    if (!logSheet) {
      logSheet = ss.insertSheet('Run Log');
      logSheet.appendRow(['Date', 'Month', 'Year', 'Term', 'Files Processed', 'Report Link']);
      logSheet.getRange('1:1').setFontWeight('bold');
    }
    logSheet.appendRow([new Date().toLocaleString('en-NZ'), month, year, term, fileCount, docUrl]);
  } catch(e) {
    Logger.log('Log error (non-fatal): ' + e.toString());
  }
}

// ── HELPER: Current month name ────────────────────────────────────────────
function getCurrentMonth() {
  return ['January','February','March','April','May','June',
          'July','August','September','October','November','December'][new Date().getMonth()];
}

// ── HELPER: Current NZ school term ───────────────────────────────────────
function getCurrentTerm() {
  const m = new Date().getMonth() + 1;
  if (m >= 2 && m <= 4) return 'Term 1';
  if (m >= 5 && m <= 7) return 'Term 2';
  if (m >= 8 && m <= 9) return 'Term 3';
  return 'Term 4';
}

// ── MENU ──────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Board Report')
    .addItem('Generate Report', 'generateBoardReport')
    .addItem('Clear Source Folder (images only)', 'clearSourceFolder')
    .addSeparator()
    .addItem('View Run Log', 'openRunLog')
    .addToUi();
}

// ── UTILITY: Clear images from source folder ──────────────────────────────
function clearSourceFolder() {
  const ui      = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Clear Source Folder',
    'This will permanently delete all image files from the source folder.\nOnly do this AFTER the report has been reviewed and approved.',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const folder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
  const files  = folder.getFiles();
  let count    = 0;
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName().toLowerCase();
    if (name.endsWith('.png') || name.endsWith('.jpg') || name.endsWith('.jpeg')) {
      file.setTrashed(true);
      count++;
    }
  }
  ui.alert(count + ' image file(s) moved to trash. The .docx report remains in the folder.');
}

// ── UTILITY: Jump to Run Log ──────────────────────────────────────────────
function openRunLog() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Run Log');
  if (logSheet) {
    ss.setActiveSheet(logSheet);
  } else {
    SpreadsheetApp.getUi().alert('No Run Log yet — generate a report first.');
  }
}
