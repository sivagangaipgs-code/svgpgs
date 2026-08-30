
const CONFIG = {
  SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/1RFOwRAuua1DYghIKjNijqpqOr6bFHcgAELXsZSGc4LI/edit?gid=0#gid=0',
  LOGIN_SHEET: 'Login mails',
  DATA_SHEET: 'Sheet2',
  SUBJECTS: [
    'Physics',
    'Chemistry',
    'Mathematics',
    'Computer Science',
    'Computer Applications'
  ]
};

function getSpreadsheet_() {
  try {
    const ss = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
    Logger.log('Connected to spreadsheet: ' + ss.getName());
    return ss;
  } catch (error) {
    throw new Error(
      'Unable to open SvgDB spreadsheet. Check that this Google account has access. ' +
      'Original error: ' + error.message
    );
  }
}

function doGet(e) {

  const page =
    (e && e.parameter && e.parameter.page)
      ? e.parameter.page
      : "Index";

  return HtmlService
    .createHtmlOutputFromFile(page)
    .setTitle("SVG PG Teachers' Resource Portal")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}
function testPortalSetup() {
  const ss = getSpreadsheet_();

  Logger.log('Spreadsheet name: ' + ss.getName());
  Logger.log('Spreadsheet URL: ' + ss.getUrl());

  const loginSheet = ss.getSheetByName(CONFIG.LOGIN_SHEET);
  if (!loginSheet) {
    throw new Error('Sheet "' + CONFIG.LOGIN_SHEET + '" was not found.');
  }

  const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET);
  if (!dataSheet) {
    throw new Error('Sheet "' + CONFIG.DATA_SHEET + '" was not found.');
  }

  Logger.log('Login mails: OK');
  Logger.log('Sheet2: OK');
  Logger.log('================================');
  Logger.log('SrgDB CONNECTION SUCCESSFUL');
  Logger.log('================================');
}


function updateResourceStatus(rowNumber, newStatus) {

  const sheet = getSpreadsheet_().getSheetByName(CONFIG.DATA_SHEET);

  const headers = sheet.getRange(
    1,
    1,
    1,
    sheet.getLastColumn()
  )
  .getValues()[0]
  .map(h => String(h || '').trim().toLowerCase());

  const statusCol = headers.indexOf('status') + 1;

  if (statusCol === 0) {
    throw new Error('Status column not found.');
  }

  sheet.getRange(rowNumber, statusCol).setValue(newStatus);

  return {
    success: true,
    row: rowNumber,
    status: newStatus
  };
}
function onFormSubmitSetStatus(e) {
  var sheet = e.range.getSheet();
  
  // Only run on Sheet2
  if (sheet.getName() !== "Sheet2") return;
  
  var row = e.range.getRow();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusCol = headers.indexOf("Status") + 1; // 1-based column index
  
  if (statusCol > 0) {
    sheet.getRange(row, statusCol).setValue("No");
  }
}
/**
 * Runs automatically whenever the Google Form
 * adds a new response to the response sheet.
 */
function onFormSubmit(e) {

  try {

    if (!e || !e.namedValues) {
      Logger.log("Form submission event data not found.");
      return;
    }

    // ---------------------------------------------------------
    // GET RESPONDENT EMAIL
    // ---------------------------------------------------------

    var email = "";

    // Possible email column names
    var emailHeaders = [
      "Email Address",
      "Email address",
      "Email",
      "E-mail",
      "E-mail Address",
      "Email ID",
      "Email Id",
      "EMAIL"
    ];

    for (var i = 0; i < emailHeaders.length; i++) {

      var header = emailHeaders[i];

      if (e.namedValues[header]) {

        email = e.namedValues[header][0];

        if (email) {
          break;
        }
      }
    }

    // ---------------------------------------------------------
    // IF EMAIL WAS NOT FOUND
    // ---------------------------------------------------------

    if (!email) {
      Logger.log("Email ID was not found in the form response.");
      return;
    }

    email = email.toString().trim().toLowerCase();

    // ---------------------------------------------------------
    // OPEN SPREADSHEET
    // ---------------------------------------------------------

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Your Login mails sheet
    var loginSheet = ss.getSheetByName("Login mails");

    if (!loginSheet) {
      throw new Error('Sheet "Login mails" was not found.');
    }

    // ---------------------------------------------------------
    // STORE EMAIL
    // ---------------------------------------------------------

    var lastRow = loginSheet.getLastRow();

    // Check whether this email already exists
    var existingEmails = [];

    if (lastRow > 1) {

      existingEmails = loginSheet
        .getRange(2, 1, lastRow - 1, 1)
        .getValues()
        .flat()
        .map(function(value) {
          return value.toString().trim().toLowerCase();
        });
    }

    // ---------------------------------------------------------
    // ADD ONLY IF EMAIL DOES NOT ALREADY EXIST
    // ---------------------------------------------------------

    if (existingEmails.indexOf(email) === -1) {

      loginSheet.appendRow([
        email,
        new Date()
      ]);

      Logger.log("New email added: " + email);

    } else {

      Logger.log("Email already exists: " + email);
    }

  } catch (error) {

    Logger.log("ERROR: " + error.message);

  }
}

function createOnFormSubmitTrigger() {
  var ss = SpreadsheetApp.openById('1RFOwRAuua1DYghIKjNijqpqOr6bFHcgAELXsZSGc4LI');

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log('Trigger created successfully.');
}
function getPortalData() {
  try {

    const publishedOnly = readResources_().filter(function(r) {
      return String(r.status || '').trim().toLowerCase() === 'yes';
    });

    return {
      success: true,
      authorized: true,
      email: '',
      subjects: CONFIG.SUBJECTS,
      resources: publishedOnly
    };

  } catch (error) {

    return {
      success: false,
      authorized: false,
      email: '',
      subjects: CONFIG.SUBJECTS,
      resources: [],
      message: error.message
    };

  }
}
function createFormSubmitTrigger() {
  const ss = getSpreadsheet_();

  // Remove any existing triggers for this function to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onFormSubmitSetStatus') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onFormSubmitSetStatus')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log('Trigger created successfully.');
}
function readResources_() {

  const sheet = getSpreadsheet_()
    .getSheetByName(CONFIG.DATA_SHEET);

  if (!sheet) {
    throw new Error(
      'Sheet "' + CONFIG.DATA_SHEET + '" was not found.'
    );
  }

  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) {
    return [];
  }

  // Normalize column headers
  const normalizeHeader = function(header) {
    return String(header || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ');
  };

  const headers = values[0].map(normalizeHeader);

  function findColumn(names) {

    const normalizedNames =
      names.map(normalizeHeader);

    for (let i = 0; i < headers.length; i++) {

      if (
        normalizedNames.indexOf(headers[i]) !== -1
      ) {
        return i;
      }

    }

    return -1;
  }

  // ==========================================
  // COLUMN DETECTION
  // ==========================================

  const c = {

    timestamp: findColumn([
      'timestamp',
      'date'
    ]),

    subject: findColumn([
      'subject'
    ]),

    className: findColumn([
      'class',
      'class name'
    ]),

    volume: findColumn([
      'volume'
    ]),

    // CHAPTER NUMBER
    unit: findColumn([
      'chapter',
      'chapter number',
      'chapter no'
    ]),

    // CHAPTER NAME
    unitName: findColumn([
      'chapter name',
      'chapter name (in english)',
      'chapter name in english'
    ]),

    // TOPIC
    topic: findColumn([
      'topic',
      'topic name',
      'topic (in english)',
      'topic in english'
    ]),

    // CHAPTER NAME IN TAMIL
    unitNameTa: findColumn([
      'chapter name (தமிழில்)',
      'chapter name (tamil)',
      'chapter name tamil',
      'அத்தியாயத்தின் பெயர்',
      'அத்தியாயத்தின் பெயர் (தமிழில்)',
      'அத்தியாயத்தின் பெயர் தமிழில்'
    ]),

    // TOPIC IN TAMIL
    topicTa: findColumn([
      'topic (தமிழில்)',
      'topic (tamil)',
      'topic tamil',
      'தலைப்பு',
      'தலைப்பு (தமிழில்)',
      'தலைப்பு தமிழில்'
    ]),

    medium: findColumn([
      'medium',
      'language'
    ]),

    teacher: findColumn([
      'teacher name',
      'teacher'
    ]),

    school: findColumn([
      'school name',
      'school'
    ]),

    link: findColumn([
      'simulation link',
      'simulation',
      'link',
      'simulation url'
    ]),

    status: findColumn([
      'status'
    ])

  };


  // ==========================================
  // SAFE VALUE READER
  // ==========================================

  const get = function(row, index) {

    if (index === -1) {
      return '';
    }

    return String(row[index] || '').trim();
  };


  // ==========================================
  // BUILD RESOURCE LIST
  // ==========================================

  const resources = [];

  for (let i = 1; i < values.length; i++) {

    const row = values[i];

    const resource = {

      id: i + 1,

      timestamp:
        get(row, c.timestamp),

      subject:
        get(row, c.subject),

      className:
        get(row, c.className),

      volume:
        get(row, c.volume),

      // Sheet Chapter
      unit:
        get(row, c.unit),

      // Sheet Chapter Name
      unitName:
        get(row, c.unitName),

      topic:
        get(row, c.topic),

      unitName_ta:
        get(row, c.unitNameTa),

      topic_ta:
        get(row, c.topicTa),

      medium:
        get(row, c.medium),

      teacher:
        get(row, c.teacher),

      school:
        get(row, c.school),

      link:
        get(row, c.link),

      status:
        get(row, c.status)

    };


    // Add only valid resource rows
    if (
      resource.subject ||
      resource.unit ||
      resource.unitName ||
      resource.topic ||
      resource.teacher ||
      resource.school
    ) {

      resources.push(resource);

    }

  }

  return resources.reverse();
}

function testCurrentUser() {
  const result = getCurrentUser();

  Logger.log(
    JSON.stringify(result, null, 2)
  );

  return result;
}

function testCompletePortal() {
  const ss = getSpreadsheet_();

  const loginSheet =
    ss.getSheetByName(CONFIG.LOGIN_SHEET);

  const dataSheet =
    ss.getSheetByName(CONFIG.DATA_SHEET);

  if (!loginSheet) {
    throw new Error('Login mails sheet not found.');
  }

  if (!dataSheet) {
    throw new Error('Sheet2 not found.');
  }

  const user = getCurrentUser();

  Logger.log('Spreadsheet: ' + ss.getName());
  Logger.log('Google account: ' + user.email);
  Logger.log('Authenticated: ' + user.authenticated);
  Logger.log('Authorized: ' + user.authorized);

  if (user.authorized) {
    Logger.log(
      'Resources loaded: ' +
      readResources_().length
    );
  }

  Logger.log(
    '========== TEST COMPLETE =========='
  );

  return {
    spreadsheet: ss.getName(),
    email: user.email,
    authenticated: user.authenticated,
    authorized: user.authorized
  };
}


/*
========================================================
OPTIONAL DEBUG FUNCTION
========================================================

Run this function once from Apps Script if you want
to check exactly which column numbers were detected.

It will show:
Subject
Class
Volume
Unit
Unit Name
Topic
etc.

in the Execution log.
*/

function testColumnDetection() {
  const sheet =
    getSpreadsheet_().getSheetByName(CONFIG.DATA_SHEET);

  if (!sheet) {
    throw new Error(
      'Sheet "' + CONFIG.DATA_SHEET + '" was not found.'
    );
  }

  const values =
    sheet.getDataRange().getDisplayValues();

  if (values.length === 0) {
    Logger.log('No data found.');
    return;
  }

  const headers = values[0].map(function(header) {
    return String(header || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ');
  });

  Logger.log('========== SHEET HEADERS ==========');

  for (let i = 0; i < headers.length; i++) {
    Logger.log(
      'Column ' +
      (i + 1) +
      ': "' +
      headers[i] +
      '"'
    );
  }

  Logger.log(
    '==================================='
  );
}