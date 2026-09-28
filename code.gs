/**
 * TRIBERA JD MAILER - SINGLE BACKEND FILE
 * =======================================
 */

var TRIBERA_JD = {
  SHEET: 'Sep-26',
  GRAPH: 'https://graph.microsoft.com/v1.0',

  DEFAULT_HOST: 'gadher.sharepoint.com',
  DEFAULT_SITE_PATH: '/sites/TriberaTest',
  DEFAULT_DRIVE_NAME: 'Eragina Digital - Documents',
  DEFAULT_JD_PATH: 'Assets/Requirements',
  DEFAULT_SENDER: 'anirudh@Gadher.onmicrosoft.com',
  LOG_SHEET: 'JD Mailer Log',

  PROP_ROWS: 'TRIBERA_JD_SELECTED_ROWS_FINAL',
  PROP_SHEET_ID: 'TRIBERA_JD_SELECTED_SHEET_FINAL',
  PROP_TEMPLATES: 'TRIBERA_JD_TEMPLATES_FINAL',
  PROP_DEFAULT_TEMPLATE: 'TRIBERA_JD_DEFAULT_TEMPLATE_FINAL',

  TOKEN_CACHE: 'TRIBERA_JD_GRAPH_TOKEN_FINAL',
  NOTIFICATION_TOKEN_CACHE: 'TRIBERA_JD_NOTIFICATION_TOKEN_FINAL',

  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 1500,
  ROLE_CONTEXT_CACHE_SECONDS: 21600,
  JD_META_CACHE_SECONDS: 21600,

  SYSTEM_TOKENS: [
    'CandidateName',
    'CandidateEmail',
    'Role',
    'RoleID',
    'CustomerName',
    'TaName',
    'JDFileName'
  ],

  HEADERS: {
    name: ['Name','Candidate Name','Candidate','CandidateName','Candidate Name '],
    email: ['Email','Candidate Email','Email ID','Email Address','CandidateEmail'],
    role: ['Role','Role Name','Job Role','Position','Job Title','Designation','Role/Position'],
    roleId: ['Role-ID','Role ID','RoleId','Role Id','Role_ID','Requirement ID','Requirement-ID','Role ID (E.g. - EDS-CAP-SCALA-040126)'],
    customer: ['Customer','Customer Name','Client','Client Name','Account'],
    taName: ['TA','TA Name','TA Owner','Recruiter','Recruiter Name','TA/Recruiter','TA / Recruiter'],
    taEmail: ['TA Email','TA Email ID','TA/Recruiter Email','Recruiter Email'],
    roleFiles: ['Role Files','Role Files Path','Role File Path','JD File Path','JD Path','Job Description Path','Job File Path','Job Files Path','Job Description File','JD File'],
    mobile: ['Mobile Number','Mobile','Phone','Phone Number','Contact Number'],
    totalExp: ['Total Exp','Total Experience','Total Exp (Years)'],
    relevantExp: ['Relevant Exp','Relevant Experience','Relevant Exp (Years)']
  }
};


/* ================================================================
 * MENU / OPEN
 * ================================================================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('JD Mailer')
    .addItem('Send JD To Selected Candidates', 'openSendJD')
    .addItem('Test  Users JSON Requests', 'testMultipleUsersSelectedCandidates')
    .addToUi();
}

function openSendJD() {
  try {
    var rows = captureJDSelection();
    if (!rows.length) {
      throw new Error('No candidate rows are selected. In the candidate sheet, select the candidate rows first, then open JD Mailer.');
    }
    var initialData = buildInitialDialogData_(rows);

    // Refuse to open the dialog when the selection is invalid
    // (e.g. rows from two different roles).
    if (!initialData.success) {
      throw new Error(initialData.error || 'Unable to prepare the JD Mailer dialog.');
    }

    var template = HtmlService.createTemplateFromFile('JdMailer');
    template.initialDataB64 = Utilities.base64Encode(
      Utilities.newBlob(JSON.stringify(initialData), 'application/json').getBytes()
    );

    var html = template.evaluate()
      .setWidth(1280)
      .setHeight(850);

    SpreadsheetApp.getUi().showModalDialog(html, 'Tribera JD Mailer');
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'Tribera JD Mailer',
      jdFriendlyError(err),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function buildInitialDialogData_(rows) {
  rows = (rows || []).map(Number).filter(function(r){ return r >= 2; });

  try {
    var sheet = getCandidateSheet_();
    var table = getCandidateTable_();

    // Build lightweight candidate objects (no JD lookup yet) so we can
    // validate the role selection cheaply.
    var contextCache = {};
    var lightweight = rows.map(function(rowNumber) {
      var idx = Number(rowNumber) - 2;
      if (idx < 0 || idx >= table.values.length) {
        throw new Error('Candidate row ' + rowNumber + ' does not exist.');
      }
      return candidateFromTableRow_(table.headers, table.values[idx], rowNumber, contextCache, true);
    });

    // Refuse mixed-role selections before doing any Graph work.
    validateSingleRoleSelection_(lightweight);

    var candidates = buildCandidatesForUI_(sheet, rows);

    // IMPORTANT: resolve the JD BEFORE the HTML dialog is shown.
    // The dialog must open only after SharePoint resolution has completed.
    // This is intentionally done once for the selected rows and the results
    // are passed into the HTML as initial data. The browser therefore does
    // not make a second JD request when the dialog opens.
    var jdResults = resolveCandidateJDsBeforeDialog_(candidates);

    return {
      success: true,
      sheetName: sheet.getName(),
      selectedRows: rows,
      candidates: candidates,
      jdResults: jdResults,
      templates: getEmailTemplates(),
      systemTokens: TRIBERA_JD.SYSTEM_TOKENS,
      tokenCatalog: getTokenCatalog_(),
      initialPreview: null
    };
  } catch (err) {
    return {
      success: false,
      error: jdFriendlyError(err),
      candidates: [],
      selectedRows: rows,
      templates: getEmailTemplatesSafe_(),
      systemTokens: TRIBERA_JD.SYSTEM_TOKENS,
      tokenCatalog: getTokenCatalogSafe_(),
      initialPreview: null
    };
  }
}

function getEmailTemplatesSafe_() {
  try { return getEmailTemplates(); } catch (e) { return [defaultTemplate_()]; }
}

function getTokenCatalogSafe_() {
  try { return getTokenCatalog_(); } catch (e) { return {}; }
}

function bootstrapMailerData() {
  try {
    var rows = captureJDSelection();
    if (!rows.length) rows = getStoredJDSelection().rows;
    return buildInitialDialogData_(rows);
  } catch (err) {
    return {success:false, error:jdFriendlyError(err), candidates:[], selectedRows:[], templates:getEmailTemplatesSafe_(), systemTokens:TRIBERA_JD.SYSTEM_TOKENS, tokenCatalog:getTokenCatalogSafe_()};
  }
}

function refreshJDSelection() {
  try {
    var rows = captureJDSelection();
    SpreadsheetApp.getUi().alert(
      'Tribera JD Mailer',
      rows.length
        ? rows.length + ' candidate row(s) selected.'
        : 'No candidate rows selected. Select rows from the selected candidate tab and try again.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'Tribera JD Mailer',
      jdFriendlyError(err),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


/* ================================================================
 * SELECTION
 * ================================================================ */

function captureJDSelection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var selection = ss.getSelection();
  var ranges = [];

  try {
    var rangeList = selection && selection.getActiveRangeList();
    if (rangeList) ranges = rangeList.getRanges() || [];
  } catch (e) {}

  if (!ranges.length) {
    try {
      var activeRange = selection && selection.getActiveRange();
      if (activeRange) ranges = [activeRange];
    } catch (e2) {}
  }

  if (!ranges.length) {
    try {
      var spreadsheetRange = SpreadsheetApp.getActiveRange();
      if (spreadsheetRange) ranges = [spreadsheetRange];
    } catch (e3) {}
  }

  if (!ranges.length) {
    var stored = getStoredJDSelection();
    if (stored.rows.length) return stored.rows;
    return [];
  }

  var selectedSheet = ranges[0].getSheet();

  if (!isCandidateSheet_(selectedSheet)) {
    throw new Error(
      'The selected tab "' +
      selectedSheet.getName() +
      '" does not look like a candidate tab. It must contain a Name/Candidate column and an Email column.'
    );
  }

  var rowSet = {};

  ranges.forEach(function(range) {
    if (range.getSheet().getSheetId() !== selectedSheet.getSheetId()) {
      throw new Error('Select rows from only one candidate tab.');
    }

    var firstRow = Math.max(2, range.getRow());
    var lastRow = range.getLastRow();

    for (var r = firstRow; r <= lastRow; r++) {
      rowSet[r] = true;
    }
  });

  var rows = Object.keys(rowSet)
    .map(Number)
    .filter(function(r) { return r >= 2; })
    .sort(function(a, b) { return a - b; });

  if (!rows.length) {
    clearJDSelection_();
    return [];
  }

  var props = PropertiesService.getUserProperties();
  props.setProperty(TRIBERA_JD.PROP_ROWS, JSON.stringify(rows));
  props.setProperty(TRIBERA_JD.PROP_SHEET_ID, String(selectedSheet.getSheetId()));

  return rows;
}

function getStoredJDSelection() {
  var props = PropertiesService.getUserProperties();
  var rowsRaw = props.getProperty(TRIBERA_JD.PROP_ROWS);
  var sheetId = props.getProperty(TRIBERA_JD.PROP_SHEET_ID);
  var rows = [];

  if (rowsRaw) {
    try {
      rows = JSON.parse(rowsRaw)
        .map(Number)
        .filter(function(r) { return r >= 2; })
        .sort(function(a, b) { return a - b; });
    } catch (e) {
      rows = [];
    }
  }

  return { rows: rows, sheetId: sheetId || '' };
}

function updateJDSelection(rows) {
  rows = (rows || []).map(Number).filter(function(r) { return r >= 2; });
  rows = uniqueNumbers_(rows).sort(function(a, b) { return a - b; });

  var props = PropertiesService.getUserProperties();
  props.setProperty(TRIBERA_JD.PROP_ROWS, JSON.stringify(rows));
  props.setProperty(TRIBERA_JD.PROP_SHEET_ID, String(getCandidateSheet_().getSheetId()));

  return { success: true, rows: rows };
}

function clearJDSelection_() {
  var props = PropertiesService.getUserProperties();
  props.deleteProperty(TRIBERA_JD.PROP_ROWS);
  props.deleteProperty(TRIBERA_JD.PROP_SHEET_ID);
}

function getCandidateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stored = getStoredJDSelection();

  if (stored.sheetId) {
    var storedSheet = getSheetById_(Number(stored.sheetId));
    if (storedSheet) return storedSheet;
  }

  var active = ss.getActiveSheet();
  if (active) return active;

  var configured = ss.getSheetByName(TRIBERA_JD.SHEET);
  if (configured) return configured;

  throw new Error(
    'No source sheet is available. Select candidate rows in Sep-26 and open JD Mailer again.'
  );
}

function getSheetById_(sheetId) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === Number(sheetId)) return sheets[i];
  }
  return null;
}

function isCandidateSheet_(sheet) {
  if (!sheet) return false;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return false;

  var headers = sheet.getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(function(h) { return String(h || '').trim(); });

  return findHeaderIndex_(headers, TRIBERA_JD.HEADERS.name) >= 0 &&
         findHeaderIndex_(headers, TRIBERA_JD.HEADERS.email) >= 0;
}


/* ================================================================
 * INITIAL DATA
 * ================================================================ */

function getInitialData() {
  try {
    var sheet = getCandidateSheet_();
    var selection = getStoredJDSelection();
    var candidates = buildCandidatesForUI_(sheet, selection.rows);

    return {
      success: true,
      sheetName: sheet.getName(),
      selectedRows: selection.rows,
      candidates: candidates,
      templates: getEmailTemplates(),
      systemTokens: TRIBERA_JD.SYSTEM_TOKENS,
      tokenCatalog: getTokenCatalog_()
    };

  } catch (err) {
    return {
      success: false,
      error: jdFriendlyError(err),
      candidates: [],
      selectedRows: [],
      templates: getEmailTemplates(),
      systemTokens: TRIBERA_JD.SYSTEM_TOKENS,
      tokenCatalog: getTokenCatalog_()
    };
  }
}

function getMailerData() {
  return getInitialData();
}


/* ================================================================
 * CANDIDATE READING
 * ================================================================ */

function getCandidateTable_() {
  var sheet = getCandidateSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    throw new Error('Candidate sheet is empty.');
  }

  var headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(function(h) { return String(h || '').trim(); });

  var values = [];

  if (lastRow >= 2) {
    values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  }

  return { sheet: sheet, headers: headers, values: values };
}

function loadCandidate_(rowNumber) {
  return loadCandidates_([Number(rowNumber)])[0];
}

function resolveRoleContextByRoleId_(roleId) {
  var target = normalizeRoleId_(roleId);
  if (!target) return null;

  /* RoleJD master is the source of truth for Role, Customer and Role Files.
   * Read it fresh instead of caching it, because Role Files can be changed
   * in the sheet and the next send must use the current path. */
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var preferred = ['RoleJD master', 'Role/JD master', 'Role/JD Master', 'Notes', 'Role Master', 'Roles', 'Dashboard'];

  for (var p = 0; p < preferred.length; p++) {
    var sheet = ss.getSheetByName(preferred[p]);
    var found = scanRoleSheetForRoleId_(sheet, target);
    if (found) {
      Logger.log('ROLE CONTEXT | sheet=' + found.sourceSheet +
        ' | row=' + found.sourceRow +
        ' | roleId=' + found.roleId +
        ' | role=' + found.role +
        ' | roleFiles=' + found.roleFiles);
      return found;
    }
  }

  return null;
}

function scanRoleSheetForRoleId_(sheet, target) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return null;

  var data = sheet.getDataRange().getDisplayValues();
  if (!data.length) return null;

  var headers = data[0].map(function(h) { return String(h || '').trim(); });

  var roleIdCol = findHeaderIndex_(headers, [
    'Role-ID','Role ID','RoleId','Role Id','Role_ID',
    'Requirement ID','Requirement-ID',
    'Role ID (E.g. - EDS-CAP-SCALA-040126)'
  ]);
  var roleCol = findHeaderIndex_(headers, [
    'Role','Role Name','Job Role','Position','Job Title','Designation','Role/Position'
  ]);
  var customerCol = findHeaderIndex_(headers, [
    'Customer','Customer Name','Client','Client Name','Account'
  ]);
  var roleFilesCol = findHeaderIndex_(headers, [
    'Role Files','Role Files Path','Role File Path','Role File',
    'JD File','JD File Name','JD Filename','Job Description File',
    'Job Description','JD'
  ]);
  var taEmailCol = findHeaderIndex_(headers, [
    'TA Email','TA Email ID','TA/Recruiter Email','Recruiter Email',
    'TA E-mail','Recruiter E-mail'
  ]);
  var spocNameCol = findHeaderIndex_(headers, [
    'Customer SpoC Name','Customer SPOC Name','Customer Spoc Name'
  ]);
  var spocEmailCol = findHeaderIndex_(headers, [
    'Customer SpoC Email','Customer SPOC Email','Customer Spoc Email'
  ]);

  if (roleIdCol >= 0) {
    for (var r = 1; r < data.length; r++) {
      if (normalizeRoleId_(data[r][roleIdCol]) !== target) continue;
      return {
        roleId: String(data[r][roleIdCol] || '').trim(),
        role: roleCol >= 0 ? String(data[r][roleCol] || '').trim() : '',
        customer: customerCol >= 0 ? String(data[r][customerCol] || '').trim() : '',
        roleFiles: roleFilesCol >= 0 ? String(data[r][roleFilesCol] || '').trim() : '',
        taEmail: taEmailCol >= 0 ? String(data[r][taEmailCol] || '').trim() : '',
        customerSPOCName: spocNameCol >= 0 ? String(data[r][spocNameCol] || '').trim() : '',
        customerSPOCEmail: spocEmailCol >= 0 ? String(data[r][spocEmailCol] || '').trim() : '',
        sourceSheet: sheet.getName(),
        sourceRow: r + 1
      };
    }
  }

  for (var rr = 1; rr < data.length; rr++) {
    for (var c = 0; c < data[rr].length; c++) {
      if (normalizeRoleId_(data[rr][c]) === target) {
        return {
          roleId: String(data[rr][c] || '').trim(),
          role: roleCol >= 0 ? String(data[rr][roleCol] || '').trim() : '',
          customer: customerCol >= 0 ? String(data[rr][customerCol] || '').trim() : '',
          roleFiles: roleFilesCol >= 0 ? String(data[rr][roleFilesCol] || '').trim() : '',
          taEmail: taEmailCol >= 0 ? String(data[rr][taEmailCol] || '').trim() : '',
          customerSPOCName: spocNameCol >= 0 ? String(data[rr][spocNameCol] || '').trim() : '',
          customerSPOCEmail: spocEmailCol >= 0 ? String(data[rr][spocEmailCol] || '').trim() : '',
          sourceSheet: sheet.getName(),
          sourceRow: rr + 1
        };
      }
    }
  }

  return null;
}

function normalizeRoleId_(value) {
  return String(value == null ? '' : value)
    .replace(/\u00A0/g, '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim()
    .toLowerCase()
    .replace(/[\s_\-./\\:]+/g, '');
}

/**
 * Refuses selections that contain more than one Role-ID or more than one
 * Role name. Called from buildInitialDialogData_ (dialog) and
 * prepareBatchEmail (send) so the user is stopped as early as possible.
 */
function validateSingleRoleSelection_(candidates) {
  candidates = candidates || [];
  if (!candidates.length) return;

  var roleIdMap = {};
  var roleNameMap = {};

  candidates.forEach(function(c) {
    var ridRaw = String(c.roleId || '').trim();
    var roleRaw = String(c.role || '').trim();
    var rid = normalizeRoleId_(ridRaw);
    var rn  = normalizeSearch_(roleRaw);

    if (rid) {
      if (!roleIdMap[rid]) {
        roleIdMap[rid] = { roleId: ridRaw, role: roleRaw, name: c.name, row: c.row };
      }
    } else if (rn) {
      if (!roleNameMap[rn]) {
        roleNameMap[rn] = { role: roleRaw, name: c.name, row: c.row };
      }
    }
  });

  var idKeys = Object.keys(roleIdMap);
  var nameKeys = Object.keys(roleNameMap);

  if (idKeys.length > 1) {
    var samples = idKeys.slice(0, 3).map(function(k) {
      var e = roleIdMap[k];
      return '"' + e.roleId + '" (row ' + e.row + ')';
    });
    throw new Error(
      'Select candidates belonging to ONE Role-ID at a time. ' +
      'You selected ' + idKeys.length + ' different Role-IDs: ' + samples.join(', ') +
      (idKeys.length > 3 ? ', …' : '') + '.'
    );
  }

  if (idKeys.length === 0 && nameKeys.length > 1) {
    var samples2 = nameKeys.slice(0, 3).map(function(k) {
      var e = roleNameMap[k];
      return '"' + e.role + '" (row ' + e.row + ')';
    });
    throw new Error(
      'Select candidates belonging to ONE Role at a time. ' +
      'You selected ' + nameKeys.length + ' different Roles: ' + samples2.join(', ') +
      (nameKeys.length > 3 ? ', …' : '') + '.'
    );
  }
}

function sanitizeForDisplay_(value) {
  return String(value == null ? '' : value)
    .replace(/[\u2013\u2014\u2015]/g, '-')
    .replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'")
    .replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"')
    .replace(/\u2026/g, '...')
    .replace(/[\u00A0\u202F\u2007]/g, ' ')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u2022/g, '*')
    .replace(/[^\x00-\x7F]/g, '?');
}

function buildCandidatesForUI_(sheet, rows) {
  var table = getCandidateTable_();
  var contextCache = {};
  var out = [];

  (rows || []).forEach(function(rowNumber) {
    try {
      var index = Number(rowNumber) - 2;
      if (index < 0 || index >= table.values.length) {
        throw new Error('Candidate row ' + rowNumber + ' does not exist.');
      }

      // IMPORTANT: this function only prepares sheet data. SharePoint/Graph
      // JD resolution is deliberately NOT performed here. The HTML dialog
      // calls resolveSelectedCandidateJDs() immediately after it opens.
      var row = table.values[index];
      var candidate = candidateFromTableRow_(table.headers, row, rowNumber, contextCache, true);

      out.push({
        row: candidate.row,
        name: sanitizeForDisplay_(candidate.name),
        email: sanitizeForDisplay_(candidate.email),
        role: sanitizeForDisplay_(candidate.role),
        roleId: sanitizeForDisplay_(candidate.roleId),
        customer: sanitizeForDisplay_(candidate.customer),
        taName: sanitizeForDisplay_(candidate.taName),
        taEmail: sanitizeForDisplay_(candidate.taEmail),
        mobile: sanitizeForDisplay_(candidate.mobile),
        totalExp: sanitizeForDisplay_(candidate.totalExp),
        relevantExp: sanitizeForDisplay_(candidate.relevantExp),
        roleFiles: sanitizeForDisplay_(candidate.roleFiles || ''),
        jdFileName: '',
        jdOptions: [],
        jdAvailable: null,
        jdError: ''
      });
    } catch (e) {
      out.push({
        row: Number(rowNumber), name: 'Row ' + rowNumber,
        email: '', role: '', roleId: '', customer: '', taName: '', taEmail: '',
        mobile: '', totalExp: '', relevantExp: '', roleFiles: '',
        jdFileName: '', jdOptions: [], jdAvailable: null,
        jdError: jdFriendlyError(e)
      });
    }
  });

  return out;
}

function loadCandidates_(rows) {
  var table = getCandidateTable_();
  var contextCache = {};
  return (rows || []).map(function(rowNumber) {
    var index = Number(rowNumber) - 2;
    if (index < 0 || index >= table.values.length) throw new Error('Candidate row ' + rowNumber + ' does not exist.');
    return candidateFromTableRow_(table.headers, table.values[index], rowNumber, contextCache, true);
  });
}

function candidateFromTableRow_(headers, row, rowNumber, contextCache, includeRoleContext) {
  function valueFor(type) {
    var column = findHeaderIndex_(headers, TRIBERA_JD.HEADERS[type]);
    return column < 0 ? '' : cleanText_(row[column]);
  }
  var candidate = {
    row:Number(rowNumber), name:valueFor('name') || ('Row '+rowNumber), email:valueFor('email'),
    role:valueFor('role'), roleId:valueFor('roleId'), customer:valueFor('customer'), taName:valueFor('taName'),
    taEmail:valueFor('taEmail'), mobile:valueFor('mobile'), totalExp:valueFor('totalExp'), relevantExp:valueFor('relevantExp'),
    roleFiles:valueFor('roleFiles'), roleSourceSheet:'', roleSourceRow:''
  };
  var key = normalizeRoleId_(candidate.roleId);
  var roleContext = null;
  if (includeRoleContext !== false) {
    if (key && !Object.prototype.hasOwnProperty.call(contextCache, key)) contextCache[key] = resolveRoleContextByRoleId_(candidate.roleId);
    roleContext = key ? contextCache[key] : null;
  }
  if (roleContext) {
    candidate.role = roleContext.role || candidate.role || '';
    candidate.customer = roleContext.customer || candidate.customer || '';
    candidate.taEmail = roleContext.taEmail || candidate.taEmail || '';
    candidate.roleFiles = roleContext.roleFiles || '';
    candidate.roleSourceSheet = roleContext.sourceSheet || '';
    candidate.roleSourceRow = roleContext.sourceRow || '';
    candidate.customerSPOCName = roleContext.customerSPOCName || '';
    candidate.customerSPOCEmail = roleContext.customerSPOCEmail || '';
  }
  candidate.raw = {}; headers.forEach(function(header, col) { if (header) candidate.raw[header] = cleanText_(row[col]); });
  return candidate;
}

function buildCandidateForUI_(sheet, rowNumber) {
  try {
    var c = loadCandidate_(rowNumber);
    // No SharePoint/Graph lookup here. The UI resolves a JD only after an
    // explicit Role/Role-ID edit or when the user actually sends the email.
    return {
      row: c.row,
      name: sanitizeForDisplay_(c.name),
      email: sanitizeForDisplay_(c.email),
      role: sanitizeForDisplay_(c.role),
      roleId: sanitizeForDisplay_(c.roleId),
      customer: sanitizeForDisplay_(c.customer),
      taName: sanitizeForDisplay_(c.taName),
      taEmail: sanitizeForDisplay_(c.taEmail),
      mobile: sanitizeForDisplay_(c.mobile),
      totalExp: sanitizeForDisplay_(c.totalExp),
      relevantExp: sanitizeForDisplay_(c.relevantExp),
      roleFiles: sanitizeForDisplay_(c.roleFiles || ''),
      jdFileName: '',
      jdOptions: [],
      jdAvailable: null,
      jdError: ''
    };
  } catch (err) {
    return {
      row: Number(rowNumber), name: 'Row ' + rowNumber, email: '', role: '', roleId: '',
      customer: '', taName: '', taEmail: '', mobile: '', totalExp: '', relevantExp: '',
      roleFiles: '', jdFileName: '', jdOptions: [], jdAvailable: null, jdError: jdFriendlyError(err)
    };
  }
}

function resolveCandidateJDsBeforeDialog_(candidates) {
  candidates = Array.isArray(candidates) ? candidates : [];
  var results = [];
  var jdByKey = {};

  candidates.forEach(function(candidate) {
    var key = normalizeRoleId_(candidate.roleId) || normalizeSearch_(candidate.role) || String(candidate.row);
    try {
      if (!candidate.roleId && !candidate.role) {
        throw new Error('Role-ID/Role is missing for row ' + candidate.row + '.');
      }

      var result;
      if (Object.prototype.hasOwnProperty.call(jdByKey, key)) {
        result = jdByKey[key];
      } else {
        // Re-read the candidate so Role Files comes from the current sheet /
        // RoleJD master at the moment the dialog is opened.
        var liveCandidate = loadCandidate_(candidate.row);
        result = resolveJD_(liveCandidate);
        jdByKey[key] = result;
      }

      var selected = result && result.options && result.options.length ? result.options[0] : null;
      results.push({
        row: candidate.row,
        status: selected ? 'found' : 'missing',
        fileName: selected ? sanitizeForDisplay_(selected.name) : '',
        folderPath: sanitizeForDisplay_((result && result.folderPath) || candidate.roleFiles || ''),
        webUrl: selected ? (selected.webUrl || '') : '',
        mimeType: selected ? (selected.mimeType || 'application/pdf') : '',
        driveId: selected ? (selected.driveId || '') : '',
        fileId: selected ? (selected.id || '') : '',
        error: selected ? '' : sanitizeForDisplay_((result && result.error) || 'JD not found in the Role Files path.')
      });
    } catch (e) {
      results.push({
        row: candidate.row,
        status: 'error',
        fileName: '',
        folderPath: sanitizeForDisplay_(candidate.roleFiles || ''),
        webUrl: '',
        mimeType: '',
        driveId: '',
        fileId: '',
        error: jdFriendlyError(e)
      });
    }
  });

  return results;
}

function resolveSelectedCandidateJDs() {
  var selection = getStoredJDSelection();
  if (!selection.rows.length) return [];
  var candidates = loadCandidates_(selection.rows);
  var jdByRole = {};
  var output = [];
  candidates.forEach(function(candidate) {
    var key = normalizeRoleId_(candidate.roleId) || normalizeSearch_(candidate.role);
    try {
      if (!candidate.roleId && !candidate.role) throw new Error('Role-ID/Role could not be resolved for this candidate.');
      if (!Object.prototype.hasOwnProperty.call(jdByRole, key)) jdByRole[key] = resolveJD_(candidate);
      var result = jdByRole[key], selected = result.options && result.options.length ? result.options[0] : null;
      output.push({
        row: candidate.row,
        status: selected ? 'found' : 'missing',
        fileName: selected ? sanitizeForDisplay_(selected.name) : 'JD',
        folderPath: sanitizeForDisplay_(candidate.roleFiles || ''),
        webUrl: selected ? (selected.webUrl || '') : '',
        mimeType: selected ? selected.mimeType : '',
        error: selected ? '' : sanitizeForDisplay_(result.error || 'JD not found in SharePoint.')
      });
    } catch (e) {
      output.push({row:candidate.row,status:'access-error',fileName:'JD',folderPath:sanitizeForDisplay_(candidate.roleFiles||''),webUrl:'',error:jdFriendlyError(e)});
    }
  });
  return output;
}

function resolveJDForContext(roleId, role) {
  try {
    var c = {roleId:String(roleId||'').trim(), role:String(role||'').trim(), roleFiles:''};
    var ctx = c.roleId ? resolveRoleContextByRoleId_(c.roleId) : null;
    if (ctx) {
      c.roleFiles = ctx.roleFiles || '';
      if (!c.role) c.role = ctx.role || '';
    }

    var result = resolveJD_(c);
    return {
      success:true,
      selectedName:result.selectedName||'',
      role:ctx && ctx.role ? ctx.role : c.role,
      customer:ctx && ctx.customer ? ctx.customer : '',
      roleFiles:ctx && ctx.roleFiles ? ctx.roleFiles : c.roleFiles,
      options:(result.options||[]).map(function(o){
        return {
          id:o.id,
          name:o.name,
          webUrl:o.webUrl||'',
          mimeType:o.mimeType||'application/pdf',
          driveId:o.driveId||''
        };
      }),
      error:result.error||''
    };
  } catch(e) {
    return {success:false,selectedName:'',role:String(role||''),customer:'',roleFiles:'',options:[],error:jdFriendlyError(e)};
  }
}

function getJDForCandidate(rowNumber) {
  try {
    var c = loadCandidate_(rowNumber);
    if (!c.roleId && !c.role) {
      return { success: false, row: rowNumber, options: [], selectedName: '', error: 'Role-ID/Role is missing in the Candidate sheet.' };
    }
    var result = resolveJD_(c);
    return {
      success: true,
      row: rowNumber,
      options: (result.options || []).map(function(o) {
        return { id:o.id, name:sanitizeForDisplay_(o.name), webUrl:o.webUrl||'', mimeType:o.mimeType||'application/pdf' };
      }),
      selectedName: sanitizeForDisplay_(result.selectedName || ''),
      error: result.error ? sanitizeForDisplay_(result.error) : ''
    };
  } catch (err) {
    return { success: false, row: rowNumber, options: [], selectedName: '', error: jdFriendlyError(err) };
  }
}

function uploadFolderCacheKey_(path) {
  var normalized = normalizeSharePointFolderPath_(path || '');
  if (!normalized) return '';
  return 'TRIBERA_JD_UPLOAD_FOLDER_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, normalized, Utilities.Charset.UTF_8)
  ).replace(/=+$/g, '').substring(0, 80);
}

function cacheUploadFolder_(folder, path) {
  if (!folder || !folder.id || !path) return;
  var key = uploadFolderCacheKey_(path);
  if (!key) return;
  var value = JSON.stringify({id: String(folder.id), path: String(path)});
  CacheService.getUserCache().put(key, value, TRIBERA_JD.ROLE_CONTEXT_CACHE_SECONDS);
}

function getCachedUploadFolder_(path) {
  var key = uploadFolderCacheKey_(path);
  if (!key) return null;
  var raw = CacheService.getUserCache().get(key);
  if (!raw) return null;
  try {
    var parsed = JSON.parse(raw);
    return parsed && parsed.id ? parsed : null;
  } catch (e) {
    return null;
  }
}

function resolveJD_(candidate) {
  candidate = candidate || {};

  var roleId = String(candidate.roleId || '').trim();
  var role = String(candidate.role || '').trim();
  var roleFiles = String(candidate.roleFiles || '').trim();

  if (!roleFiles) {
    return {
      options: [],
      selectedName: '',
      error: 'Role Files path is empty for Role-ID "' + roleId + '".'
    };
  }

  /*
   * IMPORTANT DESIGN RULE
   * ----------------------
   * The spreadsheet is the source of truth for the JD location.
   * We DO NOT invent a folder from the role name and we DO NOT search the
   * whole SharePoint drive first.
   *
   * RoleJD master -> Role Files -> exact SharePoint path -> JD file.
   *
   * The workbook contains paths such as:
   *
   *   Assets/Requirements/Tag8/Business Development Manager
   *
   * and:
   *
   *   Eragina Digital Solutions Private Limited\\Eragina Digital - Documents\\Assets\\Requirements\\Tag8
   *
   * The first form is already relative to the document library.
   * The second form contains the company/library display-name prefix.
   * buildSharePointPathCandidates_() removes ONLY those known prefixes.
   */

  var token = getGraphToken_(false);
  var drive = getSharePointDrive_(token, roleFiles);
  var paths = buildSharePointPathCandidates_(roleFiles);
  var errors = [];

  Logger.log('============================================================');
  Logger.log('TRIBERA JD RESOLUTION');
  Logger.log('Role ID       : ' + roleId);
  Logger.log('Role          : ' + role);
  Logger.log('Role Files    : ' + roleFiles);
  Logger.log('Drive         : ' + drive.name + ' | ' + drive.id);
  Logger.log('Graph paths   : ' + JSON.stringify(paths));
  Logger.log('============================================================');

  for (var i = 0; i < paths.length; i++) {
    var path = paths[i];

    try {
      var item = getDriveItemByPath_(drive.id, path, token);

      if (!item) {
        errors.push(path + ' -> not found');
        continue;
      }

      Logger.log('PATH FOUND: ' + path + ' | ITEM: ' + String(item.name || ''));

      /* ----------------------------------------------------------
       * Role Files may point directly to the JD file.
       * ---------------------------------------------------------- */
      if (item.file) {
        if (!isJDFile_(item)) {
          errors.push(path + ' -> found item is not PDF/DOC/DOCX');
          continue;
        }

        var direct = toJDOption_(item);
        direct.driveId = drive.id;

        var directResult = {
          options: [direct],
          selectedName: direct.name,
          error: ''
        };

        cacheJDResolution_(roleId, role, directResult);
        Logger.log('JD SELECTED (DIRECT FILE): ' + direct.name);
        return directResult;
      }

      /* ----------------------------------------------------------
       * Role Files points to a folder.
       * Search only inside THAT folder and its subfolders.
       * ---------------------------------------------------------- */
      if (item.folder) {
        var allItems = listChildrenRecursive_(drive.id, item.id, token, 6);

        var jdFiles = allItems
          .filter(isJDFile_)
          .filter(function(file) {
            return !looksLikeResume_(file.name);
          })
          .map(function(file) {
            var option = toJDOption_(file);
            option.driveId = drive.id;
            return option;
          });

        jdFiles = uniqueJDOptions_(jdFiles);

        Logger.log(
          'JD FILES UNDER PATH: ' + path +
          ' => ' +
          JSON.stringify(jdFiles.map(function(x) { return x.name; }))
        );

        if (!jdFiles.length) {
          errors.push(path + ' -> folder found, but no JD-compatible files were found');
          continue;
        }

        var selected = chooseJDFromCandidates_(jdFiles, roleId, role, true);

        if (selected) {
          var folderResult = {
            options: [selected].concat(jdFiles.filter(function(x) {
              return x.id !== selected.id;
            })),
            selectedName: selected.name,
            error: ''
          };

          cacheJDResolution_(roleId, role, folderResult);
          Logger.log('JD SELECTED (FOLDER): ' + selected.name);
          return folderResult;
        }

        errors.push(
          path +
          ' -> files found but no unambiguous JD matched Role-ID/Role. Files: ' +
          jdFiles.map(function(x) { return x.name; }).join(', ')
        );
      }
    } catch (e) {
      errors.push(path + ' -> ' + jdFriendlyError(e));
    }
  }

  /*
   * Do NOT search the whole drive as a substitute for the sheet path.
   * That was one of the reasons incorrect JDs could be selected in earlier
   * versions. The Role Files column is authoritative.
   */
  var message =
    'JD could not be fetched from the Role Files path in the sheet.\n\n' +
    'Role-ID: ' + roleId + '\n' +
    'Role: ' + role + '\n' +
    'Role Files: ' + roleFiles + '\n\n' +
    'Paths checked:\n' + paths.join('\n') + '\n\n' +
    'Details:\n' + errors.join('\n');

  Logger.log(message);

  return {
    options: [],
    selectedName: '',
    error: message
  };
}

function cacheJDResolution_(roleId, role, result) {
  if (!result || !result.selectedName || !result.options || !result.options.length) return;

  var key = normalizeRoleId_(roleId) + '__' + normalizeSearch_(role);
  if (!key) return;

  CacheService.getUserCache().put(
    'TRIBERA_JDMETA_PATH_' + key,
    JSON.stringify(result),
    TRIBERA_JD.JD_META_CACHE_SECONDS
  );
}

function buildSharePointPathCandidates_(roleFiles) {
  var rawOriginal = String(roleFiles || '')
    .trim()
    .replace(/[\r\n]+/g, ' ')
    .replace(/\\/g, '/');

  var paths = [];
  var seen = {};

  function add(value) {
    var path = String(value || '')
      .trim()
      .replace(/\\/g, '/')
      .replace(/^\/+/, '')
      .replace(/\/+$/, '');

    if (!path) return;
    path = path.replace(/\/+/g, '/');
    var key = path.toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      paths.push(path);
    }
  }

  if (!rawOriginal || rawOriginal === '-') return paths;

  /*
   * Role Files is authoritative. We only normalize prefixes that identify
   * the SharePoint tenant/site/library. We NEVER manufacture a role folder
   * from the Role name.
   */
  var raw = rawOriginal;

  // Support a complete SharePoint/Graph URL stored in Role Files.
  if (/^https?:\/\//i.test(raw)) {
    try {
      var urlPath = raw.replace(/^https?:\/\/[^/]+/i, '');
      urlPath = decodeURIComponent(urlPath);
      var sitePath = normalizeSitePath_(
        getConfig_('SP_SITE_PATH') || getConfig_('SITE_PATH') || TRIBERA_JD.DEFAULT_SITE_PATH
      ).replace(/^\/+|\/+$/g, '');
      var sitePrefix = '/' + sitePath;
      if (sitePath && urlPath.toLowerCase().indexOf(sitePrefix.toLowerCase() + '/') === 0) {
        urlPath = urlPath.substring(sitePrefix.length + 1);
      } else {
        urlPath = urlPath.replace(/^\/sites\/[^/]+\/?/i, '');
      }
      raw = urlPath;
    } catch (e) {}
  }

  // Strip tenant prefix when it is present.
  raw = raw.replace(/^Eragina Digital Solutions Private Limited\/?/i, '');

  // Strip the exact document-library prefix from the path. The drive itself
  // is selected separately by getSharePointDrive_().
  var libraryName = extractDocumentLibraryName_(raw);
  if (libraryName) {
    var libRegex = new RegExp('^' + escapeRegExp_(libraryName) + '\\/?', 'i');
    raw = raw.replace(libRegex, '');
  }

  // Generic library names used by SharePoint/OneDrive exports.
  raw = raw.replace(/^Shared Documents\/?/i, '');
  raw = raw.replace(/^Documents\/?/i, '');

  add(raw);

  /*
   * If the sheet accidentally stores a trailing "JD" folder/file marker,
   * allow the actual parent path too. This is still derived only from the
   * spreadsheet path.
   */
  var parent = raw.replace(/[\\\/]JD\/?$/i, '');
  if (parent !== raw) add(parent);

  return paths;
}

function escapeRegExp_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function chooseJDFromCandidates_(files, roleId, role, strictFolder) {
  if (!files || !files.length) return null;

  var rid = normalizeSearch_(roleId);
  var rn = normalizeSearch_(role);
  var roleWords = String(role || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter(function(word) { return word.length > 2; });

  var scored = files
    .filter(function(file) {
      return file && file.id && isJDFileName_(file.name) && !looksLikeResume_(file.name);
    })
    .map(function(file) {
      var name = String(file.name || '');
      var n = normalizeSearch_(name);
      var score = 0;
      var matchedWords = 0;

      /* Explicit JD names are strongest inside the sheet-specified folder. */
      if (isExplicitJDName_(name)) score += 50000;

      if (rid && n.indexOf(rid) >= 0) score += 10000;
      if (rn && rn.length >= 4 && n.indexOf(rn) >= 0) score += 8000;

      roleWords.forEach(function(word) {
        if (n.indexOf(normalizeSearch_(word)) >= 0) {
          matchedWords++;
          score += 600;
        }
      });

      if (roleWords.length && matchedWords === roleWords.length) score += 5000;
      if (/job[ _.-]?description/i.test(name)) score += 3000;
      if (/\bjd\b/i.test(name)) score += 2500;
      if (/\.pdf$/i.test(name)) score += 100;

      return {
        file: file,
        score: score,
        matchedWords: matchedWords
      };
    })
    .sort(function(a, b) {
      if (b.score !== a.score) return b.score - a.score;
      if (b.matchedWords !== a.matchedWords) return b.matchedWords - a.matchedWords;
      return String(a.file.name).localeCompare(String(b.file.name));
    });

  if (!scored.length) return null;

  var top = scored[0];

  /*
   * If the Role Files folder contains exactly one compatible file, that
   * file is the JD. This supports filenames such as:
   *   TAG8 Business Development Manager.pdf
   *   Customer Care Executive.pdf
   *   Field Sales Executive.pdf
   * without requiring the Role-ID to appear in the filename.
   */
  if (strictFolder && scored.length === 1) {
    return top.file;
  }

  if (isExplicitJDName_(top.file.name)) return top.file;

  if (top.score < (strictFolder ? 600 : 3000)) return null;

  if (
    scored.length > 1 &&
    scored[1].score === top.score
  ) {
    return null;
  }

  return top.file;
}

function isJDFileName_(name) {
  return /\.(pdf|docx|doc)$/i.test(String(name || ''));
}

function isExplicitJDName_(name) {
  var base = String(name || '')
    .trim()
    .replace(/\.(pdf|docx|doc)$/i, '')
    .trim()
    .toLowerCase();

  return base === 'jd' ||
    base === 'job description' ||
    /^jd[\s._-]+/.test(base) ||
    /^job[\s._-]*description/.test(base);
}

function looksLikeResume_(name) {
  var n = String(name || '').trim();

  return /(^|[ _.-])(resume|cv|curriculum[ _-]?vitae|biodata|profile)([ _.-]|$)/i.test(n) ||
    /(^|[ _.-])candidate([ _.-]|$)/i.test(n);
}


function getSharePointDrive_(token, roleFiles) {
  var host =
    getConfig_('SP_HOSTNAME') ||
    getConfig_('SITE_HOSTNAME') ||
    TRIBERA_JD.DEFAULT_HOST;

  var sitePath =
    getConfig_('SP_SITE_PATH') ||
    getConfig_('SITE_PATH') ||
    TRIBERA_JD.DEFAULT_SITE_PATH;

  var siteId = getConfig_('SP_SITE_ID');

  if (!siteId) {
    var siteUrl =
      TRIBERA_JD.GRAPH +
      '/sites/' +
      host +
      ':' +
      normalizeSitePath_(sitePath) +
      '?$select=id,displayName,webUrl';

    var siteResponse = graphRequest_(siteUrl, 'get', token);

    if (
      siteResponse.code >= 200 &&
      siteResponse.code < 300 &&
      siteResponse.data &&
      siteResponse.data.id
    ) {
      siteId = siteResponse.data.id;
    } else {
      throw new Error(
        'Could not resolve SharePoint site ' + host + ':' + sitePath +
        '. Graph HTTP ' + siteResponse.code + '.'
      );
    }
  }

  var drivesResponse = graphRequest_(
    TRIBERA_JD.GRAPH +
    '/sites/' + encodeURIComponent(siteId) +
    '/drives?$top=100',
    'get',
    token
  );

  if (drivesResponse.code < 200 || drivesResponse.code >= 300) {
    throw new Error(
      'Could not list SharePoint document libraries. Graph HTTP ' +
      drivesResponse.code + '.'
    );
  }

  var drives = drivesResponse.data.value || [];

  Logger.log('SHAREPOINT SITE: ' + host + ':' + sitePath);
  Logger.log('DOCUMENT LIBRARIES: ' + drives.map(function(d) {
    return String(d.name || '') + ' [' + String(d.id || '') + ']';
  }).join(' | '));

  /*
   * The Role Files column contains the library display name:
   * "Eragina Digital - Documents"
   *
   * Therefore prefer that exact library when it is present. This is the
   * critical fix: choosing the first documentLibrary can point Graph at the
   * wrong library when a SharePoint site has multiple libraries.
   */
  var libraryName = extractDocumentLibraryName_(roleFiles);

  if (libraryName) {
    var exact = drives.filter(function(drive) {
      return normalizeLibraryName_(drive.name) === normalizeLibraryName_(libraryName);
    })[0];

    if (exact && exact.id) {
      Logger.log('USING ROLE-FILES LIBRARY: ' + exact.name + ' | ' + exact.id);
      return { id: String(exact.id), name: String(exact.name || libraryName) };
    }

    Logger.log('ROLE-FILES LIBRARY NOT FOUND: ' + libraryName);
  }

  var configuredDriveId = getConfig_('SP_DRIVE_ID');
  if (configuredDriveId) {
    var configured = drives.filter(function(drive) {
      return String(drive.id) === String(configuredDriveId);
    })[0];

    if (configured) {
      Logger.log('USING CONFIGURED DRIVE: ' + configured.name + ' | ' + configured.id);
      return { id: String(configured.id), name: String(configured.name || 'Configured SharePoint Drive') };
    }
  }

  /*
   * Most rows in RoleJD master contain the library name explicitly. A small
   * number contain a relative path such as:
   *   Capco\TechBA_Reporting_Delivery
   *   Assets/Requirements/Tag8
   *
   * Those rows still belong to the same SharePoint library. Prefer the
   * known Tribera library BEFORE falling back to a generic "Documents"
   * library. Otherwise Graph can resolve the correct path in the wrong
   * document library.
   */
  var configuredLibraryName = getConfig_('SP_LIBRARY_NAME') || TRIBERA_JD.DEFAULT_DRIVE_NAME;
  var preferredDrive = drives.filter(function(drive) {
    return normalizeLibraryName_(drive.name) === normalizeLibraryName_(configuredLibraryName);
  })[0];

  if (preferredDrive && preferredDrive.id) {
    Logger.log('USING DEFAULT TRIBERA DRIVE: ' + preferredDrive.name + ' | ' + preferredDrive.id);
    return { id: String(preferredDrive.id), name: String(preferredDrive.name || configuredLibraryName) };
  }

  var documents = drives.filter(function(drive) {
    var name = normalizeLibraryName_(drive.name);
    return name === 'documents' || name === 'shareddocuments';
  })[0];

  if (!documents) {
    documents = drives.filter(function(drive) {
      return String(drive.driveType || '') === 'documentLibrary';
    })[0];
  }

  if (documents && documents.id) {
    return { id: String(documents.id), name: String(documents.name || 'Documents') };
  }

  throw new Error('No SharePoint document library was found for ' + host + ':' + sitePath + '.');
}

function extractDocumentLibraryName_(roleFiles) {
  var raw = String(roleFiles || '').trim().replace(/\\/g, '/');
  if (!raw) return '';

  var match = raw.match(/(?:^|\/)Eragina Digital - Documents(?:\/|$)/i);
  if (match) return 'Eragina Digital - Documents';

  /* Generic support for a path containing " - Documents" as the library. */
  var parts = raw.split('/').filter(Boolean);
  for (var i = 0; i < parts.length; i++) {
    if (/ - Documents$/i.test(parts[i])) return parts[i];
  }

  return '';
}

function normalizeLibraryName_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[\s_-]+/g, '')
    .trim();
}

function getDriveItemByPath_(driveId, path, token) {
  var requested = cleanPath_(String(path || '').replace(/\\/g, '/'));
  if (!requested) return null;

  var url =
    TRIBERA_JD.GRAPH +
    '/drives/' + encodeURIComponent(driveId) +
    '/root:/' + encodeGraphPath_(requested);

  var response = graphRequest_(url, 'get', token);

  if (response.code >= 200 && response.code < 300 && response.data) {
    return response.data;
  }

  return null;
}

function listChildrenRecursive_(driveId, folderId, token, depth) {
  var url = TRIBERA_JD.GRAPH +
    '/drives/' + encodeURIComponent(driveId) +
    '/items/' + encodeURIComponent(folderId) +
    '/children?$top=200';

  var response = graphRequest_(url, 'get', token);
  if (response.code < 200 || response.code >= 300) return [];

  var items = response.data.value || [];
  if (depth <= 0) return items;

  var nested = [];
  items.forEach(function(item) {
    if (item.folder && item.id) {
      nested = nested.concat(listChildrenRecursive_(driveId, item.id, token, depth - 1));
    }
  });

  return items.concat(nested);
}

function graphDriveSearch_(driveId, query, token) {
  var safe = String(query || '').replace(/'/g, "''");
  var url = TRIBERA_JD.GRAPH +
    '/drives/' + encodeURIComponent(driveId) +
    "/root/search(q='" + encodeURIComponent(safe) + "')?$top=100";

  var response = graphRequest_(url, 'get', token);
  if (response.code < 200 || response.code >= 300) return [];
  return response.data.value || [];
}

function isJDFile_(item) {
  return !!(item && item.file && item.name && /\.(pdf|docx|doc)$/i.test(String(item.name)));
}

function toJDOption_(item) {
  return {
    id: String(item.id || ''),
    name: String(item.name || ''),
    webUrl: String(item.webUrl || ''),
    parentPath: item.parentReference && item.parentReference.path ? String(item.parentReference.path) : '',
    mimeType: item.file && item.file.mimeType ? item.file.mimeType : 'application/pdf'
  };
}

function uniqueJDOptions_(items) {
  var seen = {};
  var result = [];
  (items || []).forEach(function(item) {
    var key = String(item.id || item.name);
    if (!seen[key]) {
      seen[key] = true;
      result.push(item);
    }
  });
  return result;
}


/* ================================================================
 * ROLE / JD EDITING FOR CURRENT SEND
 * ================================================================ */

function getRoleSelectionContext(rowNumber, field, value) {
  var c = loadCandidate_(Number(rowNumber));
  field = String(field || '').trim();
  value = String(value == null ? '' : value).trim();

  if (!value) throw new Error('The selected ' + (field || 'role detail') + ' cannot be empty.');

  if (field === 'role') c.role = value;
  else if (field === 'roleId') c.roleId = value;
  else if (field === 'customer') c.customer = value;
  else throw new Error('Unsupported role field: ' + field);

  if (field === 'roleId') {
    var ctx = resolveRoleContextByRoleId_(c.roleId);
    if (ctx) {
      c.roleFiles = ctx.roleFiles || c.roleFiles || '';
      c.role = ctx.role || c.role || '';
      c.customer = ctx.customer || c.customer || '';
      c.customerSPOCName = ctx.customerSPOCName || '';
      c.customerSPOCEmail = ctx.customerSPOCEmail || '';
    }
  }

  var jdResult = resolveJD_(c);
  return {
    success: true,
    candidate: {
      row: c.row, name: sanitizeForDisplay_(c.name), email: sanitizeForDisplay_(c.email),
      role: sanitizeForDisplay_(c.role || ''), roleId: sanitizeForDisplay_(c.roleId || ''),
      customer: sanitizeForDisplay_(c.customer || ''), taName: sanitizeForDisplay_(c.taName || ''),
      taEmail: sanitizeForDisplay_(c.taEmail || ''), roleFiles: sanitizeForDisplay_(c.roleFiles || '')
    },
    jdOptions: (jdResult.options || []).map(function(o) {
      return { id:o.id, name:sanitizeForDisplay_(o.name), webUrl:o.webUrl||'', mimeType:o.mimeType||'application/pdf' };
    }),
    selectedName: sanitizeForDisplay_(jdResult.selectedName || ''),
    jdError: jdResult.error ? sanitizeForDisplay_(jdResult.error) : ''
  };
}

function updateBatchRoleOnly(rowNumber, role, currentRoleId, customer) {
  role = String(role == null ? '' : role).trim();
  if (!role) throw new Error('Role is required.');
  return updateBatchRoleDetails(rowNumber, role, currentRoleId || '', customer || '');
}

function updateBatchRoleIdOnly(rowNumber, currentRole, roleId, customer) {
  currentRole = String(currentRole == null ? '' : currentRole).trim();
  roleId = String(roleId == null ? '' : roleId).trim();
  if (!roleId) throw new Error('Role ID is required.');
  if (!currentRole) throw new Error('Role is required before changing Role ID.');
  return updateBatchRoleDetails(rowNumber, currentRole, roleId, customer || '');
}

function updateBatchRoleDetails(rowNumber, role, roleId, customer) {
  var c = loadCandidate_(Number(rowNumber));
  role = String(role == null ? '' : role).trim();
  roleId = String(roleId == null ? '' : roleId).trim();
  customer = String(customer == null ? '' : customer).trim();

  if (!role) throw new Error('Role is required.');

  c.role = role;
  if (roleId) c.roleId = roleId;
  if (customer) c.customer = customer;

  if (c.roleId) {
    var ctx = resolveRoleContextByRoleId_(c.roleId);
    if (ctx) {
      c.roleFiles = ctx.roleFiles || c.roleFiles || '';
      if (!role) c.role = ctx.role || c.role;
      if (!customer) c.customer = ctx.customer || c.customer;
    }
  }

  var jdResult = resolveJD_(c);
  return {
    success: true,
    candidate: {
      row: c.row, name: sanitizeForDisplay_(c.name), email: sanitizeForDisplay_(c.email),
      role: sanitizeForDisplay_(c.role || ''), roleId: sanitizeForDisplay_(c.roleId || ''),
      customer: sanitizeForDisplay_(c.customer || ''), taName: sanitizeForDisplay_(c.taName || ''),
      taEmail: sanitizeForDisplay_(c.taEmail || ''), roleFiles: sanitizeForDisplay_(c.roleFiles || '')
    },
    jdOptions: (jdResult.options || []).map(function(o) {
      return { id:o.id, name:sanitizeForDisplay_(o.name), webUrl:o.webUrl||'', mimeType:o.mimeType||'application/pdf' };
    }),
    selectedName: sanitizeForDisplay_(jdResult.selectedName || ''),
    jdError: jdResult.error ? sanitizeForDisplay_(jdResult.error) : ''
  };
}

function warmJDUploadTarget(roleFiles) {
  var requestedFolder = normalizeSharePointFolderPath_(
    roleFiles || getConfig_('JD_FOLDER_PATH') || TRIBERA_JD.DEFAULT_JD_PATH
  );
  if (!requestedFolder) return {success:false, ready:false, error:'Job Description folder is not configured.'};

  var token = getGraphToken_(false);
  var drive = getSharePointDrive_(token);
  var configuredFolderId = getConfig_('SP_JD_FOLDER_ID') || getConfig_('SP_JD_FOLDER_ITEM_ID') || '';
  var folder = configuredFolderId ? {id:String(configuredFolderId), path:requestedFolder} : getCachedUploadFolder_(requestedFolder);

  if (!folder || !folder.id) {
    var candidates = buildJDFolderVariants_(roleFiles || getConfig_('JD_FOLDER_PATH') || TRIBERA_JD.DEFAULT_JD_PATH);
    for (var i = 0; i < candidates.length; i++) {
      var found = getDriveItemByPath_(drive.id, candidates[i], token);
      if (found && found.folder && found.id) {
        folder = {id:String(found.id), path:String(candidates[i])};
        cacheUploadFolder_(found, candidates[i]);
        break;
      }
    }
  }

  return {success:!!(folder && folder.id), ready:!!(folder && folder.id), driveId:String(drive.id||''), folderId:folder?String(folder.id||''):'', folderPath:folder?String(folder.path||requestedFolder):requestedFolder};
}

/* ------------------------------------------------------------------
 * UPLOAD NEW JD
 * ------------------------------------------------------------------ */
function uploadNewJD(formOrBase64, rowNumber, role, roleId, roleFiles) {
  var fileName = '';
  var bytes = null;
  var mime = '';

  if (formOrBase64 && typeof formOrBase64 === 'object' && !Array.isArray(formOrBase64)) {

    var uploadedBlob = formOrBase64.jdFile;
    if (uploadedBlob && typeof uploadedBlob.getBytes === 'function') {
      fileName = String(uploadedBlob.getName ? uploadedBlob.getName() : '').trim();
      mime = String(uploadedBlob.getContentType ? uploadedBlob.getContentType() : '').trim();
      bytes = uploadedBlob.getBytes();
    }
    else if (formOrBase64.base64 || formOrBase64.fileBase64) {
      fileName = String(formOrBase64.fileName || formOrBase64.name || '').trim();
      mime = String(formOrBase64.mimeType || formOrBase64.type || '').trim();
      var raw64 = String(formOrBase64.base64 || formOrBase64.fileBase64 || '')
        .replace(/^data:[^;]+;base64,/, '')
        .replace(/\s+/g, '');
      try {
        bytes = Utilities.base64Decode(raw64);
      } catch (eDecode) {
        throw new Error('The selected Job Description file could not be decoded. Please choose the file again.');
      }
    }

    rowNumber = Number(formOrBase64.rowNumber || rowNumber || 0);
    role = String(formOrBase64.role || role || '').trim();
    roleId = String(formOrBase64.roleId || roleId || '').trim();
    roleFiles = String(formOrBase64.roleFiles || roleFiles || '').trim();
  }

  if (!bytes && typeof formOrBase64 === 'string' && arguments.length >= 5) {
    var legacyBase64 = String(formOrBase64 || '');
    var legacyFileName = String(rowNumber || '').trim();
    var legacyRow = role;
    var legacyRole = roleId;
    var legacyRoleId = roleFiles;
    rowNumber = legacyRow;
    role = legacyRole;
    roleId = legacyRoleId;
    fileName = legacyFileName;
    try { bytes = Utilities.base64Decode(legacyBase64); }
    catch (eLegacy) { throw new Error('The selected Job Description file could not be read.'); }
  }

  fileName = String(fileName || '').trim();
  if (!fileName) throw new Error('Please select a Job Description file.');
  if (!/\.(pdf|docx|doc)$/i.test(fileName)) {
    throw new Error('Only PDF, DOCX or DOC Job Description files are supported.');
  }
  if (!bytes || !bytes.length) throw new Error('The selected Job Description file is empty.');
  if (bytes.length > 3 * 1024 * 1024) {
    throw new Error('Please upload a JD up to 3 MB so it can also be attached to the email.');
  }

  role = String(role || '').trim();
  roleId = String(roleId || '').trim();
  var requestedFolder = normalizeSharePointFolderPath_(
    roleFiles || getConfig_('JD_FOLDER_PATH') || TRIBERA_JD.DEFAULT_JD_PATH
  );
  if (!requestedFolder) throw new Error('The Job Description folder is not configured.');

  if (!mime) {
    mime = /\.pdf$/i.test(fileName)
      ? 'application/pdf'
      : (/\.docx$/i.test(fileName)
          ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          : 'application/msword');
  }

  var token = getGraphToken_(false);
  var drive = getSharePointDrive_(token);

  var configuredFolderId = getConfig_('SP_JD_FOLDER_ID') || getConfig_('SP_JD_FOLDER_ITEM_ID') || '';
  var folder = configuredFolderId ? {id:String(configuredFolderId), path:requestedFolder} : getCachedUploadFolder_(requestedFolder);

  if (!folder || !folder.id) {
    var candidates = buildJDFolderVariants_(roleFiles || getConfig_('JD_FOLDER_PATH') || TRIBERA_JD.DEFAULT_JD_PATH);
    for (var i = 0; i < candidates.length; i++) {
      var candidateFolder = getDriveItemByPath_(drive.id, candidates[i], token);
      if (candidateFolder && candidateFolder.folder && candidateFolder.id) {
        folder = {id:String(candidateFolder.id), path:String(candidates[i])};
        cacheUploadFolder_(candidateFolder, candidates[i]);
        break;
      }
    }
  }

  if (!folder || !folder.id) {
    throw new Error('The Job Description folder could not be found in SharePoint: ' + requestedFolder);
  }

  var folderPath = String(folder.path || requestedFolder);
  var url = TRIBERA_JD.GRAPH +
    '/drives/' + encodeURIComponent(drive.id) +
    '/items/' + encodeURIComponent(folder.id) +
    ':/' + encodeGraphPath_(fileName) + ':/content';

  var response = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: mime,
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json'
    },
    payload: bytes,
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('SharePoint JD upload failed. HTTP ' + code +
      (response.getContentText() ? ' - ' + truncate_(response.getContentText(), 500) : ''));
  }

  var item = {};
  try { item = JSON.parse(response.getContentText() || '{}'); } catch (ignore) {}

  var keys = uniqueStrings_([normalizeRoleId_(roleId), normalizeSearch_(role)]);
  keys.forEach(function(key) {
    if (key) CacheService.getUserCache().remove('TRIBERA_JDMETA_' + key);
  });

  return {
    success: true,
    id: String(item.id || ''),
    name: String(item.name || fileName),
    webUrl: String(item.webUrl || ''),
    mimeType: mime,
    parentPath: folderPath,
    message: 'Job Description uploaded successfully.'
  };
}


function testAllRoleJDsFromMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RoleJD master');

  if (!sheet) throw new Error('Sheet "RoleJD master" was not found.');

  var data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) throw new Error('RoleJD master has no role rows.');

  var headers = data[0];
  var roleIdCol = findHeaderIndex_(headers, TRIBERA_JD.HEADERS.roleId);
  var roleCol = findHeaderIndex_(headers, TRIBERA_JD.HEADERS.role);
  var roleFilesCol = findHeaderIndex_(headers, TRIBERA_JD.HEADERS.roleFiles);

  if (roleIdCol < 0) throw new Error('RoleJD master does not contain a Role ID column.');
  if (roleFilesCol < 0) throw new Error('RoleJD master does not contain a Role Files column.');

  var seen = {};
  var results = [];

  Logger.log('============================================================');
  Logger.log('TRIBERA - TEST ALL ROLEJD MASTER PATHS');
  Logger.log('This test uses ONLY the Role Files value from RoleJD master.');
  Logger.log('============================================================');

  for (var r = 1; r < data.length; r++) {
    var roleId = String(data[r][roleIdCol] || '').trim();
    var role = roleCol >= 0 ? String(data[r][roleCol] || '').trim() : '';
    var roleFiles = String(data[r][roleFilesCol] || '').trim();
    var key = normalizeRoleId_(roleId);

    if (!key || seen[key]) continue;
    seen[key] = true;

    var result = {
      row: r + 1,
      roleId: roleId,
      role: role,
      roleFiles: roleFiles,
      status: 'FAILED',
      jd: '',
      error: ''
    };

    try {
      var resolved = resolveJD_({
        roleId: roleId,
        role: role,
        roleFiles: roleFiles
      });

      if (resolved && resolved.selectedName) {
        result.status = 'PASS';
        result.jd = resolved.selectedName;
      } else {
        result.error = resolved && resolved.error
          ? resolved.error
          : 'JD not found.';
      }
    } catch (e) {
      result.error = jdFriendlyError(e);
    }

    results.push(result);

    Logger.log(
      '%s | row=%s | Role-ID=%s | Role=%s | Role Files=%s | JD=%s | Error=%s',
      result.status,
      result.row,
      result.roleId,
      result.role,
      result.roleFiles,
      result.jd,
      result.error
    );
  }

  var passed = results.filter(function(x) { return x.status === 'PASS'; }).length;
  var failed = results.length - passed;

  Logger.log('============================================================');
  Logger.log('ROLEJD MASTER TEST COMPLETE');
  Logger.log('Unique Role-IDs tested: ' + results.length);
  Logger.log('PASS: ' + passed);
  Logger.log('FAILED: ' + failed);
  Logger.log('============================================================');

  return {
    success: failed === 0,
    total: results.length,
    passed: passed,
    failed: failed,
    results: results
  };
}

/* ================================================================
 * TEMPLATE MANAGEMENT
 * ================================================================ */

function getEmailTemplates() {
  var defaultTemplate = defaultTemplate_();
  var defaultRaw = PropertiesService.getDocumentProperties().getProperty(TRIBERA_JD.PROP_DEFAULT_TEMPLATE);
  if (defaultRaw) {
    try {
      var savedDefault = JSON.parse(defaultRaw);
      if (savedDefault && savedDefault.id === 'default') {
        defaultTemplate = Object.assign(defaultTemplate, savedDefault, {
          id:'default',
          name:'Default JD Introduction',
          isDefault:true,
          type:'default',
          canDelete:false
        });
        defaultTemplate.body = String(defaultTemplate.body || '')
          .replace(/\{TAEmail\}|\{TaEmail\}/g, '')
          .replace(/Talent Acquisition team at \./g, 'Talent Acquisition team.');
      }
    } catch (e) {}
  }

  var raw = PropertiesService.getDocumentProperties().getProperty(TRIBERA_JD.PROP_TEMPLATES);
  var custom = [];
  if (raw) {
    try { custom = JSON.parse(raw) || []; } catch (e2) { custom = []; }
  }
  return [defaultTemplate].concat(custom.filter(function(t){
    return t && t.id && t.id !== 'default';
  }).map(function(t){
    t.canDelete = true;
    t.type = 'custom';
    return t;
  }));
}

function getTemplateDetail(templateId) {
  var template = findTemplate_(getEmailTemplates(), templateId);
  if (!template) throw new Error('Template not found.');
  return template;
}

function saveEmailTemplate(template) {
  template = template || {};
  if (String(template.id || '') === 'default') return saveDefaultTemplate_(template);
  return saveCustomTemplate(template);
}

function saveDefaultTemplate_(template) {
  template = template || {};
  var normalized = {
    id: 'default',
    name: 'Default JD Introduction',
    isDefault: true,
    type: 'default',
    canDelete: false,
    to: String(template.to || '{CandidateEmail}').trim(),
    cc: String(template.cc || '').trim(),
    heading: String(template.heading || 'Job Opportunity').trim(),
    subject: String(template.subject || '{Role} Opportunity').trim(),
    body: String(template.body || '').trim(),
    customTokens: Array.isArray(template.customTokens) ? template.customTokens : []
  };
  validateTemplate_(normalized);

  var lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    PropertiesService.getDocumentProperties().setProperty(TRIBERA_JD.PROP_DEFAULT_TEMPLATE, JSON.stringify(normalized));
    return { success: true, template: normalized, templates: getEmailTemplates() };
  } finally {
    lock.releaseLock();
  }
}

function saveCustomTemplate(template) {
  template = template || {};
  var requestedId = String(template.id || '').trim();

  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getDocumentProperties();
    var key = TRIBERA_JD.PROP_TEMPLATES;
    var raw = props.getProperty(key);
    var existing = [];
    if (raw) {
      try { existing = JSON.parse(raw) || []; }
      catch (e) { throw new Error('Saved template data could not be read.'); }
    }
    if (!Array.isArray(existing)) existing = [];

    var previous = null;
    existing.forEach(function(t) {
      if (t && String(t.id) === requestedId) previous = t;
    });

    if (requestedId && requestedId !== 'default' && !previous) {
      throw new Error('The custom template being edited no longer exists. Refresh templates and try again.');
    }

    var normalized = normalizeTemplate_(template);
    normalized.canDelete = true;
    normalized.createdAt = previous && previous.createdAt ? previous.createdAt : new Date().toISOString();
    normalized.updatedAt = new Date().toISOString();

    validateTemplate_(normalized);

    existing = existing.filter(function(t) {
      return t && t.id && String(t.id) !== String(normalized.id) && String(t.id) !== 'default';
    });
    existing.push(normalized);

    props.setProperty(key, JSON.stringify(existing));

    var verifyRaw = props.getProperty(key);
    var verified = [];
    try { verified = JSON.parse(verifyRaw || '[]') || []; }
    catch (e2) { throw new Error('Template save verification failed.'); }

    var savedTemplate = verified.filter(function(t) {
      return t && String(t.id) === String(normalized.id);
    })[0];
    if (!savedTemplate) throw new Error('Template was not saved.');

    return {
      success: true,
      saved: true,
      template: savedTemplate,
      templates: [defaultTemplate_()].concat(verified)
    };
  } finally {
    lock.releaseLock();
  }
}

function duplicateTemplate(templateId) {
  var source = findTemplate_(getEmailTemplates(), templateId);
  if (!source) throw new Error('Template not found.');

  var copy = JSON.parse(JSON.stringify(source));
  copy.id = 'tpl_' + Utilities.getUuid();
  copy.name = source.name + ' Copy';
  copy.isDefault = false;

  return saveCustomTemplate(copy);
}

function deleteCustomTemplate(templateId) {
  return deleteEmailTemplate(templateId);
}

function deleteEmailTemplate(templateId) {
  return deleteCustomTemplates([templateId]);
}

function deleteCustomTemplates(templateIds) {
  var ids = Array.isArray(templateIds) ? templateIds : [templateIds];
  ids = ids.map(function(id) { return String(id == null ? '' : id).trim(); })
    .filter(function(id) { return id && id !== 'default'; });
  ids = ids.filter(function(id, index) { return ids.indexOf(id) === index; });

  if (!ids.length) throw new Error('Select at least one editable custom template.');

  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getDocumentProperties();
    var key = TRIBERA_JD.PROP_TEMPLATES;
    var raw = props.getProperty(key);
    var existing = [];
    if (raw) {
      try { existing = JSON.parse(raw) || []; }
      catch (e) { throw new Error('Saved template data could not be read.'); }
    }
    if (!Array.isArray(existing)) existing = [];

    var found = {};
    existing.forEach(function(t) { if (t && t.id) found[String(t.id)] = t; });

    var invalid = ids.filter(function(id) {
      return !found[id] || found[id].type === 'default' || String(id) === 'default';
    });
    if (invalid.length) {
      throw new Error('One or more selected templates cannot be deleted. Refresh the template list and try again.');
    }

    var deleteSet = {};
    ids.forEach(function(id) { deleteSet[id] = true; });
    var remaining = existing.filter(function(t) {
      return !(t && deleteSet[String(t.id)]);
    });

    props.setProperty(key, JSON.stringify(remaining.filter(function(t) {
      return t && String(t.id) !== 'default';
    })));

    var verifyRaw = props.getProperty(key) || '[]';
    var verify = JSON.parse(verifyRaw);
    var stillThere = verify.filter(function(t) { return t && deleteSet[String(t.id)]; });
    if (stillThere.length) throw new Error('Template deletion could not be verified.');

    return { success: true, deletedIds: ids, templates: getEmailTemplates() };
  } finally {
    lock.releaseLock();
  }
}

function defaultTemplate_() {
  return {
    id:'default', name:'Default JD Introduction', isDefault:true, type:'default',
    to:'{CandidateEmail}',
    cc:getConfig_('DEFAULT_CC') || '',
    heading:'Job Opportunity',
    subject:'{Role} Opportunity',
    body:'Hello {CandidateName},\n\n' +
      'We are pleased to share an opportunity for the {Role} position with {CustomerName}.\n\n' +
      'Please find the Job Description attached for your review.\n\n' +
      'Please review the attached Job Description and let us know your interest.\n\n' +
      'If you have any queries regarding this opportunity, please reach out to the Talent Acquisition team.\n\n' +
      'Regards,\n{TaName}\nTalent Acquisition Team',
    customTokens:[]
  };
}

function normalizeTokenSyntax_(value) {
  return String(value == null ? '' : value)
    .replace(/\{\{\s*([A-Za-z][A-Za-z0-9_]*)\s*\}\}/g, '{$1}');
}

function normalizeTemplate_(template) {
  template = template || {};

  var id = String(template.id || '').trim();
  if (!id || id === 'default') {
    id = 'tpl_' + Utilities.getUuid();
  }

  var customTokens =
    Array.isArray(template.customTokens)
      ? template.customTokens.map(function(item) {
          var token = normalizeCustomToken_(item.token || item.name || '');
          return {
            token: token,
            label: String(item.label || token).trim(),
            value: normalizeTokenSyntax_(String(item.value == null ? '' : item.value))
          };
        }).filter(function(item) { return item.token; })
      : [];

  return {
    id: id,
    name: String(template.name || '').trim(),
    isDefault: false,
    type: 'custom',
    to: normalizeTokenSyntax_(String(template.to || '').trim()),
    cc: normalizeTokenSyntax_(String(template.cc || '').trim()),
    heading: normalizeTokenSyntax_(String(template.heading || 'Job Opportunity')),
    subject: normalizeTokenSyntax_(String(template.subject || '')),
    body: normalizeTokenSyntax_(String(template.body || '')),
    customTokens: customTokens,
    canDelete: template.canDelete === true
  };
}

function validateTemplate_(template) {
  if (!template.name) throw new Error('Template name is required.');
  if (!template.to) throw new Error('To field is required.');
  if (!template.heading.trim()) throw new Error('Email heading is required.');
  if (!template.subject.trim()) throw new Error('Subject is required.');
  if (!template.body.trim()) throw new Error('Body is required.');

  var allowed = {};
  TRIBERA_JD.SYSTEM_TOKENS.forEach(function(token) { allowed[token] = true; });

  var custom = {};
  (template.customTokens || []).forEach(function(item) {
    if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(item.token)) {
      throw new Error('Invalid custom token "' + item.token + '". Use letters, numbers and underscore only.');
    }
    if (allowed[item.token]) {
      throw new Error('Custom token conflicts with system token "' + item.token + '".');
    }
    if (custom[item.token]) {
      throw new Error('Duplicate custom token "' + item.token + '".');
    }
    custom[item.token] = true;
  });

  var tokens = extractTokens_(
    template.to + '\n' + template.cc + '\n' + template.subject + '\n' + template.body + '\n' +
    (template.customTokens || []).map(function(item){ return item.value || ''; }).join('\n')
  );

  tokens.forEach(function(token) {
    if (!allowed[token] && !custom[token]) {
      throw new Error('Unknown token {' + token + '}.');
    }
  });
}

function findTemplate_(templates, id) {
  return (templates || []).filter(function(t) { return t.id === id; })[0] || null;
}

function getTokenCatalog_() {
  var labels = {
    CandidateName: 'Candidate name · Sep-26 Name',
    CandidateEmail: 'Candidate email · Sep-26 Email',
    Role: 'Role · Notes/Dashboard',
    RoleID: 'Role ID · Sep-26 / Notes',
    CustomerName: 'Customer · Sep-26 / Notes',
    TaName: 'TA / recruiter · Sep-26 / Notes',
    JDFileName: 'JD filename · SharePoint'
  };

  return {
    system: TRIBERA_JD.SYSTEM_TOKENS.map(function(name) {
      return {name:name, label:labels[name] || name, source:'Approved'};
    }),
    sheetColumns: [],
    suggestions: []
  };
}


/* ================================================================
 * RENDER / PREVIEW
 * ================================================================ */

function getRenderedMailerData(templateId) {
  var template = findTemplate_(getEmailTemplates(), templateId || 'default');
  if (!template) throw new Error('Template not found.');

  var rows = getStoredJDSelection().rows;
  if (!rows.length) throw new Error('No candidate rows are selected.');

  var candidates = rows.map(function(row) {
    var c = loadCandidate_(row);
    return { row: c.row, candidate: c, template: template, unresolved: [], jd: null };
  });

  return {
    success: true,
    template: template,
    candidates: candidates,
    templates: getEmailTemplates(),
    tokenCatalog: getTokenCatalog_()
  };
}


function prepareEmails(templateId, options) {
  var template = findTemplate_(getEmailTemplates(), templateId || 'default');
  if (!template) throw new Error('Template not found.');
  options = options || {};
  return prepareBatchEmail({
    template: template,
    overrides: options.overrides || null,
    draft: options.draft || null
  });
}

function mergeSendDraftIntoTemplate_(template, draft) {
  template = template || {};
  draft = draft || {};

  var out = {
    id: template.id,
    name: template.name,
    to: template.to,
    cc: template.cc,
    heading: template.heading,
    subject: template.subject,
    body: template.body,
    customTokens: template.customTokens,
    isDefault: template.isDefault,
    type: template.type,
    canDelete: template.canDelete
  };

  if (draft.to != null && String(draft.to).trim()) out.to = String(draft.to);
  if (draft.cc != null) out.cc = String(draft.cc);
  if (draft.heading != null && String(draft.heading).trim()) out.heading = String(draft.heading);
  if (draft.subject != null && String(draft.subject).trim()) out.subject = String(draft.subject);
  if (draft.body != null && String(draft.body).trim()) out.body = String(draft.body);
  if (Array.isArray(draft.customTokens)) out.customTokens = draft.customTokens;

  return out;
}

function prepareBatchEmail(payload) {
  payload = payload || {};
  var template = payload.template || {};
  var draft = payload.draft || null;

  if (draft) template = mergeSendDraftIntoTemplate_(template, draft);

  var rows = getStoredJDSelection().rows;
  if (!rows.length) throw new Error('No selected candidates.');

  var normalized = normalizeTemplate_(template);
  validateTemplate_(normalized);

  var candidates = loadCandidates_(rows);
  var overrides = payload.overrides || {};

  candidates.forEach(function(c) {
    var roleEdited = overrides.roleEdited === true;
    var roleIdEdited = overrides.roleIdEdited === true;
    var customerEdited = overrides.customerEdited === true;

    if (overrides.role != null && String(overrides.role).trim()) {
      c.role = String(overrides.role).trim();
    }
    if (overrides.roleId != null) {
      c.roleId = String(overrides.roleId || '').trim();
    }
    if (overrides.customer != null) {
      c.customer = String(overrides.customer || '').trim();
    }

    /*
     * IMPORTANT: the JD was already resolved before the dialog opened.
     * Therefore Role / Role ID / Customer edits are presentation/email
     * overrides only. They must NOT trigger another RoleJD lookup or another
     * SharePoint JD resolution.
     *
     * Only legacy callers that do not provide overrides.jd need the old
     * RoleJD-master context refresh so prepareBatchEmail remains backwards
     * compatible.
     */
    if (!overrides.jd || !overrides.jd.id) {
      if (c.roleId) {
        var ctx = resolveRoleContextByRoleId_(c.roleId);
        if (ctx) {
          c.roleFiles = ctx.roleFiles || c.roleFiles || '';
          if (!roleEdited) c.role = ctx.role || c.role;
          if (!customerEdited && !roleEdited) c.customer = ctx.customer || c.customer;
          c.roleSourceSheet = ctx.sourceSheet || '';
          c.roleSourceRow = ctx.sourceRow || '';
        }
      }
    }
  });

  // Hard check: refuse mixed Role-ID / Role selections. This is the last
  // line of defence before emails go out.
  validateSingleRoleSelection_(candidates);

  var commonJDResult = {options: [], selectedName: '', error: ''};
  var commonJD = null;

  if (overrides.jd && overrides.jd.id) {
    commonJD = {
      id: String(overrides.jd.id),
      name: String(overrides.jd.name || overrides.jd.fileName || ''),
      webUrl: String(overrides.jd.webUrl || ''),
      mimeType: String(overrides.jd.mimeType || 'application/pdf'),
      driveId: String(overrides.jd.driveId || '')
    };
    commonJDResult.options = [commonJD];
    commonJDResult.selectedName = commonJD.name;
  } else {
    // Legacy/fallback path only. Normal UI sends overrides.jd, which is the
    // JD resolved before the dialog opened, so this branch is not used for
    // normal sends after the dialog has been displayed.
    commonJDResult = resolveJD_(candidates[0]);
    var wantedJD = String(overrides.jdFileName || '').trim();
    var options = commonJDResult.options || [];
    if (wantedJD) {
      var match = options.filter(function(x){ return String(x.name || '') === wantedJD; })[0];
      if (match) commonJD = match;
    }
    if (!commonJD) commonJD = options.length ? options[0] : null;
  }

  if (!commonJD || !commonJD.id) {
    throw new Error('Job Description could not be resolved for Role-ID "' + String(candidates[0].roleId || '') + '" / Role "' + String(candidates[0].role || '') + '". Email was not sent.');
  }

  var emails = candidates.map(function(c){
    var jd = commonJD;
    var customValues = {}, unresolved = [];
    (normalized.customTokens || []).forEach(function(item){
      customValues[item.token] = renderTokens_(item.value, c, jd, customValues, unresolved);
    });
    var heading = renderTokens_(normalized.heading, c, jd, customValues, unresolved);
    var to = renderTokens_(normalized.to, c, jd, customValues, unresolved);
    var cc = parseEmailList_(renderTokens_(normalized.cc, c, jd, customValues, unresolved)).join(',');
    var subject = renderTokens_(normalized.subject, c, jd, customValues, unresolved);
    var body = renderTokens_(normalized.body, c, jd, customValues, unresolved);
    unresolved = uniqueStrings_(unresolved);
    var errors = [];
    if (!isValidEmail_(to)) errors.push('Candidate email is missing or invalid.');
    parseEmailList_(cc).forEach(function(e){ if(!isValidEmail_(e)) errors.push('Invalid CC address: ' + e); });
    if (!heading.trim()) errors.push('Email heading is required.');
    if (!subject.trim()) errors.push('Subject is required.');
    if (!body.trim()) errors.push('Email body is required.');
    if (unresolved.length) errors.push('Unresolved token(s): ' + unresolved.map(function(t){return '{'+t+'}';}).join(', '));

    return {
      row:c.row, candidateName:c.name, email:c.email, role:c.role, roleId:c.roleId, customer:c.customer,
      taName:c.taName, taEmail:c.taEmail, roleFiles:c.roleFiles||'', heading:heading, to:to, cc:cc, subject:subject, body:body,
      jd:jd ? {status:'found',id:jd.id,fileName:jd.name,webUrl:jd.webUrl||'',mimeType:jd.mimeType||'application/pdf',driveId:jd.driveId||'',path:c.roleFiles||''} : {status:'missing',id:'',fileName:'',webUrl:'',mimeType:'',driveId:'',path:c.roleFiles||''},
      validation:{valid:errors.length===0,errors:errors,unresolved:unresolved}
    };
  });

  return {
    success:true,
    template:normalized,
    emails:emails,
    batch:{count:emails.length,role:emails[0].role||'',roleId:emails[0].roleId||'',customer:emails[0].customer||'',jdFileName:commonJD ? commonJD.name : ''}
  };
}

function previewFirstPageEmail(rowNumber, templateId, customerOverride, draft, roleOverride, roleIdOverride, jdOverride) {
  var template = findTemplate_(getEmailTemplates(), templateId || 'default');
  if (!template) throw new Error('Template not found.');

  var c = loadCandidate_(Number(rowNumber));

  if (roleOverride != null && String(roleOverride).trim()) c.role = String(roleOverride).trim();
  if (roleIdOverride != null) c.roleId = String(roleIdOverride || '').trim();
  if (customerOverride != null && String(customerOverride).trim()) c.customer = String(customerOverride).trim();
  if (c.roleId && roleIdOverride != null) {
    var roleCtx = resolveRoleContextByRoleId_(c.roleId);
    if (roleCtx && roleCtx.roleFiles) c.roleFiles = roleCtx.roleFiles;
  }

  draft = draft || {};

  var runtimeTemplate = {
    id: 'runtime_first_page',
    name: template.name,
    to: draft.to != null ? String(draft.to) : template.to,
    cc: draft.cc != null ? String(draft.cc) : template.cc,
    heading: template.heading,
    subject: draft.subject != null ? String(draft.subject) : template.subject,
    body: draft.body != null ? String(draft.body) : template.body,
    customTokens: template.customTokens || []
  };

  var normalized = normalizeTemplate_(runtimeTemplate);
  normalized.id = 'runtime_first_page';

  validateTemplate_(normalized);

  var jdResult = {options: [], selectedName: '', error: ''};
  var selectedJD = null;
  if (jdOverride && jdOverride.id) {
    selectedJD = {id:String(jdOverride.id), name:String(jdOverride.name || jdOverride.fileName || ''), webUrl:String(jdOverride.webUrl || ''), mimeType:String(jdOverride.mimeType || 'application/pdf')};
    jdResult.options = [selectedJD];
    jdResult.selectedName = selectedJD.name;
  } else {
    jdResult = resolveJD_(c);
    selectedJD = (jdResult.options && jdResult.options.length) ? jdResult.options[0] : null;
  }

  var customValues = {};
  var unresolved = [];

  (normalized.customTokens || []).forEach(function(item) {
    customValues[item.token] = renderTokens_(item.value, c, selectedJD, customValues, unresolved);
  });

  var to = renderTokens_(normalized.to, c, selectedJD, customValues, unresolved);
  var cc = renderTokens_(normalized.cc, c, selectedJD, customValues, unresolved);
  var subject = renderTokens_(normalized.subject, c, selectedJD, customValues, unresolved);
  var body = renderTokens_(normalized.body, c, selectedJD, customValues, unresolved);

  unresolved = uniqueStrings_(unresolved);

  var errors = [];
  if (!isValidEmail_(to)) errors.push('Candidate email is missing or invalid.');
  parseEmailList_(cc).forEach(function(e){ if(!isValidEmail_(e)) errors.push('Invalid CC address: ' + e); });
  if (!subject.trim()) errors.push('Subject is required.');
  if (!body.trim()) errors.push('Email body is required.');
  if (unresolved.length) {
    errors.push('Unresolved token(s): ' + unresolved.map(function(t) { return '{' + t + '}'; }).join(', '));
  }

  return {
    success: true,
    row: c.row,
    candidate: {
      row: c.row, name: sanitizeForDisplay_(c.name), email: sanitizeForDisplay_(c.email),
      role: sanitizeForDisplay_(c.role || ''), roleId: sanitizeForDisplay_(c.roleId || ''),
      customer: sanitizeForDisplay_(c.customer || ''),
      taName: sanitizeForDisplay_(c.taName || ''),
      taEmail: sanitizeForDisplay_(c.taEmail || ''),
      roleFiles: sanitizeForDisplay_(c.roleFiles || '')
    },
    to: to, cc: cc, subject: subject, body: body,
    jd: selectedJD ? {
      id: selectedJD.id, name: sanitizeForDisplay_(selectedJD.name),
      webUrl: selectedJD.webUrl || '', mimeType: selectedJD.mimeType || 'application/pdf'
    } : null,
    jdOptions: (jdResult.options || []).map(function(o) {
      return { id:o.id, name:sanitizeForDisplay_(o.name), webUrl:o.webUrl||'', mimeType:o.mimeType||'application/pdf' };
    }),
    jdError: jdResult.error ? sanitizeForDisplay_(jdResult.error) : '',
    unresolved: unresolved,
    errors: errors,
    valid: errors.length === 0
  };
}

function previewCandidateEmail(rowNumber, templateId, jdFileName) {
  var template = findTemplate_(getEmailTemplates(), templateId);
  if (!template) throw new Error('Template not found.');

  var c = loadCandidate_(rowNumber);
  var jdResult = resolveJD_(c);
  var selectedJD = null;

  if (jdFileName) {
    selectedJD = jdResult.options.filter(function(jd) { return jd.name === jdFileName; })[0] || null;
  }
  if (!selectedJD && jdResult.options.length) selectedJD = jdResult.options[0];

  var customValues = {};
  (template.customTokens || []).forEach(function(item) {
    customValues[item.token] = resolveTextTokens_(item.value, c, selectedJD, customValues);
  });

  var unresolved = [];
  var to = renderTokens_(template.to, c, selectedJD, customValues, unresolved);
  var cc = renderTokens_(template.cc, c, selectedJD, customValues, unresolved);
  var subject = renderTokens_(template.subject, c, selectedJD, customValues, unresolved);
  var body = renderTokens_(template.body, c, selectedJD, customValues, unresolved);
  unresolved = uniqueStrings_(unresolved);

  return {
    success: true,
    row: c.row,
    candidate: {
      row: c.row, name: sanitizeForDisplay_(c.name), email: sanitizeForDisplay_(c.email),
      role: sanitizeForDisplay_(c.role || ''), roleId: sanitizeForDisplay_(c.roleId || ''),
      customer: sanitizeForDisplay_(c.customer || '')
    },
    to: to, cc: cc, subject: subject, body: body,
    jd: selectedJD ? { id:selectedJD.id, name:sanitizeForDisplay_(selectedJD.name), webUrl:selectedJD.webUrl||'', mimeType:selectedJD.mimeType||'application/pdf' } : null,
    jdOptions: (jdResult.options || []).map(function(o) {
      return { id:o.id, name:sanitizeForDisplay_(o.name), webUrl:o.webUrl||'', mimeType:o.mimeType||'application/pdf' };
    }),
    jdError: jdResult.error ? sanitizeForDisplay_(jdResult.error) : '',
    unresolved: unresolved,
    valid: !unresolved.length && !!selectedJD && isValidEmail_(to)
  };
}

function previewRuntimeTemplate(rowNumber, template, jdFileName) {
  var normalized = normalizeTemplate_(template || {});
  validateTemplate_(normalized);
  var c = loadCandidate_(Number(rowNumber));
  var jdResult = resolveJD_(c);
  var selectedJD = null;
  if (jdFileName) {
    selectedJD = (jdResult.options || []).filter(function(x){ return x.name === jdFileName; })[0] || null;
  }
  if (!selectedJD && jdResult.options && jdResult.options.length) selectedJD = jdResult.options[0];
  var customValues = {}, unresolved = [];
  (normalized.customTokens || []).forEach(function(item){
    customValues[item.token] = renderTokens_(item.value, c, selectedJD, customValues, unresolved);
  });
  return {
    success:true,
    candidateName:sanitizeForDisplay_(c.name),
    to:renderTokens_(normalized.to,c,selectedJD,customValues,unresolved),
    cc:renderTokens_(normalized.cc,c,selectedJD,customValues,unresolved),
    heading:renderTokens_(normalized.heading,c,selectedJD,customValues,unresolved),
    subject:renderTokens_(normalized.subject,c,selectedJD,customValues,unresolved),
    body:renderTokens_(normalized.body,c,selectedJD,customValues,unresolved),
    jd: selectedJD ? { id:selectedJD.id, name:sanitizeForDisplay_(selectedJD.name), webUrl:selectedJD.webUrl||'', mimeType:selectedJD.mimeType||'application/pdf' } : null,
    jdFileName: selectedJD ? sanitizeForDisplay_(selectedJD.name) : '',
    unresolved:uniqueStrings_(unresolved)
  };
}

function resolveTextTokens_(text, candidate, jd, customValues) {
  var unresolved = [];
  return renderTokens_(String(text || ''), candidate, jd, customValues || {}, unresolved);
}

function renderTokens_(text, candidate, jd, customValues, unresolved) {
  var values = {
    CandidateName: candidate.name,
    CandidateEmail: candidate.email,
    Role: candidate.role,
    RoleID: candidate.roleId,
    CustomerName: candidate.customer,
    TaName: candidate.taName,
    JDFileName: jd ? jd.name : ''
  };

  var input = normalizeTokenSyntax_(String(text || ''));

  return input.replace(
    /\{([A-Za-z][A-Za-z0-9_]*)\}/g,
    function(full, token) {
      if (Object.prototype.hasOwnProperty.call(values, token)) {
        return String(values[token] == null ? '' : values[token]);
      }
      if (customValues && Object.prototype.hasOwnProperty.call(customValues, token)) {
        return String(customValues[token] == null ? '' : customValues[token]);
      }
      unresolved.push(token);
      return full;
    }
  );
}



/* ================================================================
 * MULTI-USER PAYLOAD TEST MODE
 *
 * TEST ONLY:
 * - Does NOT send email.
 * - Does NOT call Graph / Notification API.
 * - Does NOT write shared mail logs.
 * - Receives the payload from the current user's browser execution,
 *   logs it to the Apps Script Execution log, and returns it.
 *
 * This deliberately avoids ScriptProperties/DocumentProperties for
 * temporary payload state so concurrent users do not overwrite each other.
 * ================================================================ */

function logJDMailerPayload(payload) {
  payload = payload || {};

  var activeUser = '';
  var temporaryUserKey = '';
  try {
    activeUser = String(Session.getActiveUser().getEmail() || '');
  } catch (e) {}
  try {
    temporaryUserKey = String(Session.getTemporaryActiveUserKey() || '');
  } catch (e2) {}

  var record = {
    testMode: true,
    action: 'JD_MAILER_PAYLOAD_TEST',
    executionId: Utilities.getUuid(),
    serverTimestamp: new Date().toISOString(),
    activeUser: activeUser || '(email unavailable for this deployment)',
    temporaryUserKey: temporaryUserKey || '(unavailable)',
    payload: payload
  };

  Logger.log('========== JD MAILER MULTI-USER PAYLOAD TEST ==========' );
  Logger.log(JSON.stringify(record, null, 2));
  Logger.log('=========================================================' );

  return {
    success: true,
    testMode: true,
    executionId: record.executionId,
    serverTimestamp: record.serverTimestamp,
    activeUser: record.activeUser,
    temporaryUserKey: record.temporaryUserKey,
    message: 'Payload received and written to the Apps Script Execution log. No email was sent.'
  };
}

/**
 * Manual editor test.
 * Select candidate rows in the Sheet, run this function from Apps Script,
 * then open Executions to inspect the JSON payload.
 *
 * This is useful for testing selection isolation before using the HTML UI.
 */
function testSelectedCandidatesPayload() {
  var rows = captureJDSelection();
  if (!rows.length) {
    throw new Error('Select one or more candidate rows first.');
  }

  var candidates = loadCandidates_(rows);
  validateSingleRoleSelection_(candidates);

  var payload = {
    version: 'multi-user-payload-test-v1',
    clientTimestamp: new Date().toISOString(),
    source: 'Apps Script editor test',
    sheet: getCandidateSheet_().getName(),
    selectedRows: rows.slice(),
    candidateCount: candidates.length,
    candidates: candidates.map(function(c) {
      return {
        row: Number(c.row || 0),
        name: String(c.name || ''),
        email: String(c.email || ''),
        role: String(c.role || ''),
        roleId: String(c.roleId || ''),
        customer: String(c.customer || ''),
        taName: String(c.taName || ''),
        taEmail: String(c.taEmail || ''),
        roleFiles: String(c.roleFiles || '')
      };
    })
  };

  return logJDMailerPayload(payload);
}

/* ================================================================
 * SEND - MODE DISPATCH
 * ================================================================ */

function getSendMode_() {
  var props = PropertiesService.getScriptProperties();

  var explicit = String(props.getProperty('JD_MAILER_SEND_MODE') || '')
    .trim()
    .toLowerCase();

  if (explicit === 'test' || explicit === 'graph' || explicit === 'notification') {
    return explicit;
  }

  var testMode = String(props.getProperty('JD_MAILER_TEST_MODE') || '')
    .toLowerCase() === 'true';

  return testMode ? 'test' : 'notification';
}

function sendBatchEmail(payload) {
  payload = payload || {};
  var prepared = prepareBatchEmail({
    template: payload.template || {},
    overrides: payload.overrides || {},
    draft: payload.draft || null
  });
  if (!prepared || !prepared.success || !prepared.emails || !prepared.emails.length) {
    throw new Error((prepared && prepared.error) || 'No candidate emails were prepared.');
  }
  var result = sendPreparedEmails({ items: prepared.emails });
  result.prepared = prepared.emails.length;
  return result;
}

function sendPreparedEmails(payload) {
  payload = payload || {};
  var items = Array.isArray(payload.items) ? payload.items : [];

  if (!items.length) throw new Error('No candidate emails were prepared.');

  var invalid = items.filter(function(item) {
    return !isValidEmail_(item.to || item.email || '');
  });

  if (invalid.length) {
    throw new Error(
      'Some selected candidates have an invalid email address: ' +
      invalid.map(function(item) {
        return item.candidateName || item.name || 'Candidate';
      }).join(', ')
    );
  }

  var mode = getSendMode_();

  if (mode === 'test') return sendPreparedEmailsToHttpbin_(items);
  if (mode === 'graph') return sendPreparedEmailsViaGraph_(items);
  return sendPreparedEmailsViaNotificationApi_(items);
}


/* ================================================================
 * SEND - NOTIFICATION API
 * ================================================================ */

function sendPreparedEmailsViaNotificationApi_(items) {
  items = Array.isArray(items) ? items : [];

  if (!items.length) {
    throw new Error('No candidate emails were prepared.');
  }

  /*
   * IMPORTANT PAYLOAD RULE
   * ----------------------
   * 1 candidate  -> one JSON object
   * 2+ candidates -> ONE JSON array containing all candidate objects
   *
   * Example for two candidates:
   * [
   *   { ...Nihal payload... },
   *   { ...Arpit payload... }
   * ]
   *
   * This is ONE HTTP request to Notification Service.
   */
  var payloads = items.map(function(item) {
    return buildNotificationPayload_(item);
  });

  var outboundPayload = payloads.length === 1
    ? payloads[0]
    : payloads;

  var token = getNotificationAccessToken();
  var request = buildNotificationBatchRequest_(outboundPayload, token);

  var response;

  try {
    response = UrlFetchApp.fetch(request.url, {
      method: request.method,
      contentType: request.contentType,
      headers: request.headers,
      payload: request.payload,
      muteHttpExceptions: true
    });
  } catch (e) {
    throw new Error(
      'Notification API batch request failed: ' + jdFriendlyError(e)
    );
  }

  var code = response.getResponseCode();
  var body = response.getContentText() || '';
  var data = {};

  try {
    data = body ? JSON.parse(body) : {};
  } catch (e2) {
    data = {};
  }

  var ok = code >= 200 && code < 300;

  /*
   * The API response belongs to the ONE batch request.
   * If the service returns candidate-level results, use them.
   * Otherwise apply the HTTP result to every selected candidate.
   */
  var responseResults = Array.isArray(data.results)
    ? data.results
    : (Array.isArray(data.items) ? data.items : null);

  var results = items.map(function(item, index) {
    var candidateResponse = responseResults && responseResults[index]
      ? responseResults[index]
      : {};

    var candidateOk =
      responseResults && responseResults[index] &&
      typeof candidateResponse.success === 'boolean'
        ? candidateResponse.success
        : ok;

    return {
      success: candidateOk,
      row: Number(item.row || 0),
      name: item.candidateName || item.name || 'Candidate',
      email: item.to || item.email || '',
      roleId: item.roleId || '',
      status: candidateOk
        ? (candidateResponse.status || data.status || 'queued')
        : 'Failed',
      messageId:
        candidateResponse.message_id ||
        candidateResponse.messageId ||
        data.message_id ||
        data.messageId ||
        data.id ||
        '',
      httpCode: code,
      jdFileName: item.jd && item.jd.fileName || '',
      attachment: false,
      reason: candidateOk
        ? ''
        : (
            candidateResponse.reason ||
            'Notification API failed. HTTP ' + code +
            (body ? ' - ' + truncate_(body, 700) : '')
          ),
      details: candidateOk
        ? 'Notification Service accepted the JD notification.'
        : 'Notification Service rejected the JD notification.'
    };
  });

  writeMailLogsBatch_(items, results);

  var sent = results.filter(function(r) {
    return r.success;
  }).length;

  var failed = results.length - sent;

  Logger.log('');
  Logger.log('============================================================');
  Logger.log('TRIBERA JD MAILER - NOTIFICATION BATCH REQUEST');
  Logger.log('============================================================');
  Logger.log('Candidates: ' + items.length);
  Logger.log('HTTP requests sent: 1');
  Logger.log('Payload shape: ' + (payloads.length === 1 ? 'OBJECT' : 'ARRAY'));
  Logger.log('HTTP status: ' + code);
  Logger.log('Successful candidates: ' + sent);
  Logger.log('Failed candidates: ' + failed);
  Logger.log('Response received: YES');
  Logger.log('============================================================');

  return {
    success: failed === 0,
    total: results.length,
    sent: sent,
    successful: sent,
    failed: failed,
    requestCount: 1,
    candidateCount: items.length,
    payloadShape: payloads.length === 1 ? 'object' : 'array',
    results: results,
    response: data,
    timing: 'Notification API single batch request',
    realEmailSent: false
  };
}

function buildNotificationBatchRequest_(payload, token) {
  var props = PropertiesService.getScriptProperties();
  var baseUrl = String(
    props.getProperty('NOTIFICATION_API_URL') || ''
  ).trim();

  if (!baseUrl) {
    throw new Error('NOTIFICATION_API_URL is missing.');
  }

  baseUrl = baseUrl.replace(/\/+$/, '');

  return {
    url: baseUrl + '/api/v1/notifications/send',
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
}


/* ================================================================
 * SEND - MICROSOFT GRAPH (JD attached)
 * ================================================================ */

function sendPreparedEmailsViaGraph_(items){
  var token=getGraphToken_(false),firstJD=items[0]&&items[0].jd;if(!firstJD||!firstJD.id)throw new Error('No valid SharePoint Job Description was prepared. Email was not sent.');
  var drive=firstJD.driveId ? {id:String(firstJD.driveId),name:'Role Files drive'} : getSharePointDrive_(token),blob=downloadJD_(drive.id,{id:firstJD.id,name:firstJD.fileName||firstJD.name},token),attachmentName=blob.getName()||firstJD.fileName||firstJD.name||'Job Description.pdf',attachmentMime=blob.getContentType()||firstJD.mimeType||'application/pdf',attachmentBase64=Utilities.base64Encode(blob.getBytes());
  if(!attachmentBase64)throw new Error('The Job Description attachment is empty. Email was not sent.');
  var requests=items.map(function(item){return {item:item,request:buildGraphRequest_(item,attachmentBase64,attachmentName,attachmentMime,token)};}),responses;
  try{responses=UrlFetchApp.fetchAll(requests.map(function(x){return x.request;}));}catch(batchError){responses=requests.map(function(x){try{var r=sendGraphRequestWithRetry_(x.request);return {fallback:true,code:r.getResponseCode(),body:r.getContentText()};}catch(e){return {fallback:false,code:0,body:jdFriendlyError(e)};}});}
  var results=requests.map(function(x,i){var r=responses[i],code=r&&r.fallback?Number(r.code||0):Number(r&&r.getResponseCode?r.getResponseCode():r&&r.code||0),body=r&&r.getContentText?r.getContentText():String(r&&r.body||'');if(code>=400&&isTemporaryGraphCode_(code)){try{r=sendGraphRequestWithRetry_(x.request);code=r.getResponseCode();body=r.getContentText();}catch(e){code=0;body=jdFriendlyError(e);}}var ok=code>=200&&code<300;return {success:ok,row:Number(x.item.row||0),name:x.item.candidateName||x.item.name||'Candidate',email:x.item.to||x.item.email||'',roleId:x.item.roleId||'',status:ok?'Sent to Exchange Online':'Failed',httpCode:code,jdFileName:attachmentName,attachment:ok,reason:ok?'':'Microsoft Graph send failed. HTTP '+code+(body?' - '+truncate_(body,700):''),details:ok?'Exchange Online accepted the message with the Job Description attached.':'Microsoft Graph rejected the message.'};});
  writeMailLogsBatch_(items,results);var sent=results.filter(function(r){return r.success;}).length,failed=results.length-sent;return {success:failed===0,total:results.length,sent:sent,successful:sent,failed:failed,results:results,timing:'One SharePoint download + parallel Graph fetchAll'};
}

function buildGraphMessage_(item, attachmentBase64, attachmentName, attachmentMimeType) {
  var message = {
    subject: String(item.subject || '').trim(),
    body: {
      contentType: 'HTML',
      content:
        '<div style="font-family:Segoe UI,Arial,sans-serif">' +
        '<h2 style="margin:0 0 18px;font-size:20px">' +
        htmlEscape_(item.heading || '') +
        '</h2>' +
        textToSimpleHtml_(item.body || '') +
        '</div>'
    },
    toRecipients: [
      { emailAddress: { address: String(item.to || '').trim() } }
    ]
  };

  if (attachmentBase64 && attachmentName) {
    message.attachments = [{
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: attachmentName,
      contentType: attachmentMimeType || 'application/pdf',
      contentBytes: attachmentBase64
    }];
  }

  var ccList = parseEmailList_(item.cc || '');
  if (ccList.length) {
    message.ccRecipients = ccList.map(function(e) {
      return { emailAddress: { address: e } };
    });
  }

  return message;
}


/* ================================================================
 * HTTPBIN BACKEND TEST MODE
 * ================================================================ */

function sendPreparedEmailsToHttpbin_(items) {
  var endpoint = 'https://httpbin.org/post';
  items = Array.isArray(items) ? items : [];

  if (!items.length) {
    throw new Error('No candidate emails were prepared.');
  }

  var selection = getStoredJDSelection();
  var selectedRows = (selection.rows || [])
    .map(Number)
    .filter(function(row) { return row >= 2; })
    .sort(function(a, b) { return a - b; });

  if (!selectedRows.length) {
    throw new Error('No candidate rows are selected in the Google Sheet.');
  }

  var sheet = getCandidateSheet_();

  if (
    selection.sheetId &&
    String(sheet.getSheetId()) !== String(selection.sheetId)
  ) {
    throw new Error(
      'The candidate sheet changed after selection. Please select the candidates again.'
    );
  }

  if (items.length !== selectedRows.length) {
    throw new Error(
      'Prepared candidate count (' + items.length +
      ') does not match selected sheet row count (' + selectedRows.length + ').'
    );
  }

  var selectedRowMap = {};
  selectedRows.forEach(function(row) {
    selectedRowMap[String(row)] = true;
  });

  var preparedRowMap = {};
  var payloads = [];
  var candidateMeta = [];

  items.forEach(function(item) {
    var row = Number(item.row || 0);

    if (!row) {
      throw new Error('Prepared candidate is missing its Sheet row number.');
    }

    if (!selectedRowMap[String(row)]) {
      throw new Error(
        'Prepared candidate row ' + row +
        ' was not part of the current Google Sheet selection.'
      );
    }

    if (preparedRowMap[String(row)]) {
      throw new Error('Duplicate prepared candidate row: ' + row);
    }

    preparedRowMap[String(row)] = true;

    var candidate = loadCandidate_(row);
    var preparedEmail = String(item.to || item.email || '').trim();
    var sheetEmail = String(candidate.email || '').trim();

    if (!isValidEmail_(preparedEmail)) {
      throw new Error(
        'Invalid prepared email at row ' + row + ': ' + preparedEmail
      );
    }

    if (sheetEmail.toLowerCase() !== preparedEmail.toLowerCase()) {
      throw new Error(
        'Candidate email mismatch at row ' + row +
        '. Sheet email: ' + sheetEmail +
        '. Prepared email: ' + preparedEmail
      );
    }

    payloads.push(buildNotificationPayload_(item));

    candidateMeta.push({
      row: row,
      candidate: String(
        item.candidateName ||
        item.name ||
        candidate.name ||
        'Candidate'
      ),
      email: preparedEmail
    });
  });

  /*
   * EXACT REQUEST SHAPE:
   *
   * 1 candidate:
   * {
   *   candidate_name: "...",
   *   ...
   * }
   *
   * 2+ candidates:
   * [
   *   {
   *     candidate_name: "...",
   *     ...
   *   },
   *   {
   *     candidate_name: "...",
   *     ...
   *   }
   * ]
   *
   * There is exactly ONE HTTP request in both cases.
   */
  var outboundPayload = payloads.length === 1
    ? payloads[0]
    : payloads;

  var request = {
    url: endpoint,
    method: 'post',
    contentType: 'application/json',
    headers: {
      Accept: 'application/json'
    },
    payload: JSON.stringify(outboundPayload),
    muteHttpExceptions: true
  };

  Logger.log('');
  Logger.log('============================================================');
  Logger.log('TRIBERA JD MAILER - EXACT JSON REQUEST TEST');
  Logger.log('============================================================');
  Logger.log('Endpoint: ' + endpoint);
  Logger.log('Sheet: ' + sheet.getName());
  Logger.log('Selected candidates: ' + payloads.length);
  Logger.log('HTTP requests sent: 1');
  Logger.log(
    'Payload shape: ' +
    (payloads.length === 1 ? 'OBJECT' : 'ARRAY')
  );
  Logger.log('Real Notification/AWS API called: NO');
  Logger.log('Real emails sent: 0');
  Logger.log('------------------------------------------------------------');
  Logger.log('JSON REQUEST:');
  Logger.log(JSON.stringify(outboundPayload, null, 2));
  Logger.log('------------------------------------------------------------');

  var started = Date.now();
  var response;

  try {
    response = UrlFetchApp.fetch(endpoint, request);
  } catch (err) {
    throw new Error(
      'JSON HTTP request failed: ' + jdFriendlyError(err)
    );
  }

  var elapsedMs = Date.now() - started;
  var httpCode = response.getResponseCode();
  var responseBody = response.getContentText() || '';
  var parsedResponse = null;

  try {
    parsedResponse = responseBody
      ? JSON.parse(responseBody)
      : null;
  } catch (e) {
    parsedResponse = null;
  }

  var ok = httpCode >= 200 && httpCode < 300;

  Logger.log('JSON RESPONSE:');
  if (parsedResponse !== null) {
    Logger.log(JSON.stringify(parsedResponse, null, 2));
  } else {
    Logger.log(responseBody || '(empty response)');
  }

  Logger.log('------------------------------------------------------------');

  var results = candidateMeta.map(function(meta) {
    return {
      success: ok,
      requestNumber: 1,
      row: meta.row,
      candidate: meta.candidate,
      email: meta.email,
      httpCode: httpCode,
      responseSize: responseBody.length
    };
  });

  Logger.log('RESULT: ' + (ok ? 'SUCCESS' : 'FAILED'));
  Logger.log('Total candidates: ' + payloads.length);
  Logger.log('HTTP requests: 1');
  Logger.log('Elapsed time: ' + elapsedMs + ' ms');
  Logger.log('============================================================');

  return {
    success: ok,
    requestCount: 1,
    total: payloads.length,
    candidateCount: payloads.length,
    payloadShape: payloads.length === 1 ? 'object' : 'array',
    successful: ok ? payloads.length : 0,
    passed: ok ? payloads.length : 0,
    failed: ok ? 0 : payloads.length,
    elapsedMs: elapsedMs,
    results: results,
    payload: outboundPayload,
    response: parsedResponse !== null
      ? parsedResponse
      : responseBody,
    payloadPrinted: true,
    realEmailSent: false,
    realNotificationApiCalled: false
  };
}

function buildNotificationPayload_(item) {
  item = item || {};

  /*
   * THIS IS THE EXACT NOTIFICATION/AWS BUSINESS PAYLOAD.
   *
   * Do not add transport/debug fields here.
   * The Notification Service receives only these 11 fields.
   */
  return {
    candidate_name: String(item.candidateName || item.name || '').trim(),
    candidate_email: String(item.to || item.email || '').trim(),
    role: String(item.role || '').trim(),
    role_id: String(item.roleId || '').trim(),
    customer: String(item.customer || '').trim(),
    ta_name: String(item.taName || '').trim(),
    ta_email: String(item.taEmail || '').trim(),
    cc: String(item.cc || '').trim(),
    subject: String(item.subject || ''),
    body: String(item.body || ''),
    jd_filename: String(item.jd && item.jd.fileName || '').trim()
  };
}

function buildNotificationRequest_(item, token) {
  var props = PropertiesService.getScriptProperties();
  var baseUrl = String(props.getProperty('NOTIFICATION_API_URL') || '').trim();

  if (!baseUrl) throw new Error('NOTIFICATION_API_URL is missing.');
  baseUrl = baseUrl.replace(/\/+$/, '');

  return {
    url: baseUrl + '/api/v1/notifications/send',
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json'
    },
    payload: JSON.stringify(buildNotificationPayload_(item)),
    muteHttpExceptions: true
  };
}

/*
 * TEST ONLY.
 *
 * Builds the same prepared email data used by the real send flow and logs
 * the exact Notification/AWS payload. It never calls the Notification API
 * and never sends an email.
 */
function logNotificationPayloadFromDialog(payload) {
  payload = payload || {};

  var prepared = prepareBatchEmail({
    template: payload.template || {},
    overrides: payload.overrides || {},
    draft: payload.draft || null
  });

  if (
    !prepared ||
    !prepared.success ||
    !prepared.emails ||
    !prepared.emails.length
  ) {
    throw new Error(
      (prepared && prepared.error) ||
      'No candidate emails were prepared.'
    );
  }

  var payloads = prepared.emails.map(function(item) {
    return buildNotificationPayload_(item);
  });

  var outboundPayload = payloads.length === 1
    ? payloads[0]
    : payloads;

  Logger.log('============================================================');
  Logger.log('TRIBERA JD MAILER - EXACT NOTIFICATION/AWS PAYLOAD');
  Logger.log('============================================================');
  Logger.log('EMAIL SENDING: DISABLED');
  Logger.log('CANDIDATE COUNT: ' + payloads.length);
  Logger.log(
    'PAYLOAD SHAPE: ' +
    (payloads.length === 1 ? 'OBJECT' : 'ARRAY')
  );
  Logger.log('HTTP REQUEST COUNT: 1');
  Logger.log('------------------------------------------------------------');
  Logger.log(JSON.stringify(outboundPayload, null, 2));
  Logger.log('------------------------------------------------------------');
  Logger.log('NO EMAIL SENT.');
  Logger.log('============================================================');

  return {
    success: true,
    sent: 0,
    requestCount: 1,
    candidateCount: payloads.length,
    payloadShape: payloads.length === 1 ? 'object' : 'array',
    payload: outboundPayload
  };
}


function sendGraphRequestWithRetry_(request) {
  var lastError = '';
  for (var attempt = 1; attempt <= TRIBERA_JD.MAX_RETRIES; attempt++) {
    var response;
    try {
      response = UrlFetchApp.fetch(request.url, {
        method: request.method,
        contentType: request.contentType,
        headers: request.headers,
        payload: request.payload,
        muteHttpExceptions: true
      });
    } catch (e) {
      lastError = 'Network error while calling Microsoft Graph: ' + jdFriendlyError(e);
      if (attempt < TRIBERA_JD.MAX_RETRIES) {
        Utilities.sleep(Math.min(3000, TRIBERA_JD.RETRY_DELAY_MS * attempt));
        continue;
      }
      throw new Error(lastError);
    }

    var code = response.getResponseCode();
    if (code >= 200 && code < 300) return response;

    var body = response.getContentText();
    lastError = 'Microsoft Graph send failed. HTTP ' + code + (body ? ' - ' + truncate_(body, 700) : '');
    if (!isTemporaryGraphCode_(code) || attempt >= TRIBERA_JD.MAX_RETRIES) {
      throw new Error(lastError);
    }

    var headers = response.getHeaders() || {};
    var retryAfter = Number(headers['Retry-After'] || headers['retry-after'] || 0);
    Utilities.sleep(retryAfter > 0 ? Math.min(retryAfter * 1000, 10000) : Math.min(3000, TRIBERA_JD.RETRY_DELAY_MS * attempt));
  }
  throw new Error(lastError || 'Microsoft Graph send failed.');
}

function buildGraphRequest_(item, attachmentBase64, attachmentName, attachmentMimeType, token) {
  var sender = getConfig_('GRAPH_SENDER_UPN') || TRIBERA_JD.DEFAULT_SENDER;
  var url = TRIBERA_JD.GRAPH + '/users/' + encodeURIComponent(sender) + '/sendMail';

  return {
    url: url,
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json'
    },
    payload: JSON.stringify({
      message: buildGraphMessage_(item, attachmentBase64, attachmentName, attachmentMimeType),
      saveToSentItems: true
    }),
    muteHttpExceptions: true
  };
}


function writeMailLogsBatch_(items, results) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(TRIBERA_JD.LOG_SHEET) || ss.insertSheet(TRIBERA_JD.LOG_SHEET);
    var headers = ['Timestamp', 'Candidate', 'Email', 'Role-ID', 'Customer', 'JD File', 'Subject', 'Graph Status', 'Delivery Status', 'Trace ID', 'Error'];

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }

    var rows = (results || []).map(function (r, i) {
      var item = items[i] || {};
      return [
        new Date(),
        item.candidateName || item.name || '',
        item.to || item.email || '',
        item.roleId || '',
        item.customer || '',
        r.jdFileName || (item.jd && item.jd.fileName) || '',
        item.subject || '',
        r.status || '',
        r.success ? 'CHECKING DELIVERY' : '',
        r.messageId || '',
        r.success ? '' : (r.reason || '')
      ];
    });

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    }
  } catch (e) {
    Logger.log('Batch log write failed: ' + jdFriendlyError(e));
  } finally {
    lock.releaseLock();
  }
}

function writeMailLog_(candidate, item, jdFileName, status, errorText, traceStatus, traceId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRIBERA_JD.LOG_SHEET) || ss.insertSheet(TRIBERA_JD.LOG_SHEET);
  var headers = ['Timestamp','Candidate','Email','Role-ID','Customer','JD File','Subject','Graph Status','Delivery Status','Trace ID','Error'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  } else {
    var current = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(), headers.length)).getDisplayValues()[0];
    if (current.indexOf('Delivery Status') < 0 || current.indexOf('Trace ID') < 0) {
      sheet.getRange(1,1,1,headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
  sheet.appendRow([
    new Date(),
    (candidate && candidate.name) || (candidate && candidate.candidateName) || '',
    item && item.to || candidate && candidate.email || candidate && candidate.email || '',
    (candidate && candidate.roleId) || '',
    (candidate && candidate.customer) || '',
    jdFileName || '',
    item && item.subject || '',
    status || '',
    traceStatus || 'CHECKING DELIVERY',
    traceId || '',
    errorText || ''
  ]);
  var row = sheet.getLastRow();
  sheet.getRange(row,1).setNumberFormat('dd-mmm-yyyy hh:mm:ss');
  return row;
}

function refreshMailDeliveryStatuses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRIBERA_JD.LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {updated:0,message:'No JD Mailer Log entries found.'};

  var values = sheet.getDataRange().getDisplayValues();
  var headers = values[0].map(function(h){return String(h||'').trim();});
  var idx = {};
  headers.forEach(function(h,i){idx[h]=i;});
  var required = ['Timestamp','Email','Subject','Graph Status','Delivery Status'];
  for (var i=0;i<required.length;i++){ if(idx[required[i]]==null) return {updated:0,message:'JD Mailer Log needs the latest columns. Send one new email or recreate the log sheet.'}; }

  var token;
  try { token = getGraphToken_(false); } catch(e) { return {updated:0,message:jdFriendlyError(e)}; }
  var sender = getConfig_('GRAPH_SENDER_UPN') || TRIBERA_JD.DEFAULT_SENDER;
  var updated = 0;

  for (var r=1;r<values.length;r++){
    var graphStatus = String(values[r][idx['Graph Status']]||'');
    if (!/SENT_TO_EXCHANGE/i.test(graphStatus)) continue;
    var email = String(values[r][idx['Email']]||'').trim();
    var subject = String(values[r][idx['Subject']]||'').trim();
    if (!email||!subject) continue;

    var filter = "senderAddress eq '"+escapeOData_(sender)+"' and recipientAddress eq '"+escapeOData_(email)+"' and contains(subject, '"+escapeOData_(subject)+"')";
    var url = TRIBERA_JD.GRAPH+'/admin/exchange/tracing/messageTraces?$top=20&$filter='+encodeURIComponent(filter);
    var response = graphRequest_(url,'get',token);
    if (response.code === 401 || response.code === 403) {
      return {updated:updated,permissionRequired:true,message:'Delivery tracking is not authorized. Add ExchangeMessageTrace.Read.All (application permission) and admin consent, then refresh delivery status again.'};
    }
    if (response.code<200||response.code>=300) continue;
    var traces = response.data && response.data.value || [];
    if (!traces.length) continue;
    traces.sort(function(a,b){return String(b.receivedDateTime||'').localeCompare(String(a.receivedDateTime||''));});
    var trace = traces[0];
    var status = String(trace.status||'pending').toUpperCase();
    var display = {DELIVERED:'DELIVERED',PENDING:'PENDING',FAILED:'FAILED',QUARANTINED:'QUARANTINED',FILTEREDASSPAM:'FILTERED AS SPAM',FILTERED_AS_SPAM:'FILTERED AS SPAM',EXPANDED:'EXPANDED',GETTINGSTATUS:'CHECKING DELIVERY'}[status]||'CHECKING DELIVERY';
    sheet.getRange(r+1, idx['Delivery Status']+1).setValue(display);
    if (idx['Trace ID']!=null) sheet.getRange(r+1, idx['Trace ID']+1).setValue(trace.id||'');
    updated++;
  }
  return {updated:updated,permissionRequired:false,message:updated ? ('Delivery status refreshed for '+updated+' log entr'+(updated===1?'y':'ies')+'.') : 'No new Exchange delivery trace was available yet. Entries remain CHECKING DELIVERY.'};
}

function escapeOData_(value){ return String(value||'').replace(/'/g,"''"); }

function htmlEscape_(text){
  return String(text||'')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function sendGraphMailWithRetry_(message, token) {
  var sender = getConfig_('GRAPH_SENDER_UPN') || TRIBERA_JD.DEFAULT_SENDER;
  var url = TRIBERA_JD.GRAPH+'/users/'+encodeURIComponent(sender)+'/sendMail';
  var lastError = '';
  for (var attempt=1;attempt<=TRIBERA_JD.MAX_RETRIES;attempt++){
    var response;
    try {
      response = UrlFetchApp.fetch(url,{
        method:'post',
        contentType:'application/json',
        headers:{Authorization:'Bearer '+token, Accept:'application/json'},
        payload:JSON.stringify({message:message, saveToSentItems:true}),
        muteHttpExceptions:true
      });
    } catch(e) {
      lastError = 'Network error while calling Microsoft Graph: '+jdFriendlyError(e);
      if (attempt<TRIBERA_JD.MAX_RETRIES){Utilities.sleep(TRIBERA_JD.RETRY_DELAY_MS*attempt);continue;}
      throw new Error(lastError);
    }
    var code = response.getResponseCode(), body = response.getContentText();
    if (code>=200 && code<300) return true;
    lastError = 'Microsoft Graph send failed. HTTP '+code+(body?' - '+truncate_(body,500):'');
    if (!isTemporaryGraphCode_(code) || attempt>=TRIBERA_JD.MAX_RETRIES) throw new Error(lastError);
    var retryAfter = Number(response.getHeaders()['Retry-After']||response.getHeaders()['retry-after']||0);
    Utilities.sleep(retryAfter>0?retryAfter*1000:TRIBERA_JD.RETRY_DELAY_MS*attempt);
  }
  throw new Error(lastError||'Microsoft Graph send failed.');
}

function isTemporaryGraphCode_(code) {
  return [429, 500, 502, 503, 504].indexOf(Number(code)) >= 0;
}


/* ================================================================
 * SHAREPOINT DOWNLOAD
 * ================================================================ */

function downloadJD_(driveId, jd, token) {
  var url = TRIBERA_JD.GRAPH +
    '/drives/' + encodeURIComponent(driveId) +
    '/items/' + encodeURIComponent(jd.id) +
    '/content';

  if (!/\.pdf$/i.test(String(jd.name || ''))) {
    url += '?format=pdf';
  }

  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/pdf'
    },
    followRedirects: true,
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('SharePoint JD download failed. HTTP ' + code + '.');
  }

  var blob = response.getBlob();
  var bytes = blob.getBytes();

  if (!bytes.length) throw new Error('The SharePoint JD file is empty.');

  var name = String(jd.name || 'Job Description.pdf').replace(/\.(docx|doc)$/i, '') + '.pdf';

  return blob.setName(name).setContentType('application/pdf');
}


/* ================================================================
 * GRAPH AUTH
 * ================================================================ */

function getGraphToken_(forceRefresh) {
  var cache = CacheService.getScriptCache();

  if (!forceRefresh) {
    var cached = cache.get(TRIBERA_JD.TOKEN_CACHE);
    if (cached) return cached;
  }

  // Double-check after potential race
  var cached = cache.get(TRIBERA_JD.TOKEN_CACHE);
  if (cached && !forceRefresh) return cached;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Second check inside lock
    cached = cache.get(TRIBERA_JD.TOKEN_CACHE);
    if (cached && !forceRefresh) return cached;

    var tenant = getRequiredConfig_('MS_TENANT_ID');
    var clientId = getRequiredConfig_('MS_CLIENT_ID');
    var secret = getRequiredConfig_('MS_CLIENT_SECRET');

    var url = 'https://login.microsoftonline.com/' + encodeURIComponent(tenant) + '/oauth2/v2.0/token';

    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      payload: {
        client_id: clientId,
        client_secret: secret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code < 200 || code >= 300) {
      throw new Error('Microsoft Graph authentication failed. HTTP ' + code + '.');
    }

    var json;
    try { json = JSON.parse(body); }
    catch (e) { throw new Error('Microsoft Graph returned an invalid token response.'); }

    if (!json.access_token) throw new Error('Microsoft Graph did not return an access token.');

    var expires = Number(json.expires_in || 3600);

    cache.put(
      TRIBERA_JD.TOKEN_CACHE,
      json.access_token,
      Math.max(60, Math.min(3300, expires - 120))
    );

    return json.access_token;
  } finally {
    lock.releaseLock();
  }
}

function getJDPdfPreview(rowNumber, jdFileName) {
  try {
    var c = loadCandidate_(Number(rowNumber));
    var result = resolveJD_(c);
    var jd = null;
    if (jdFileName) jd = (result.options||[]).filter(function(x){return x.name===jdFileName;})[0]||null;
    if (!jd && result.options && result.options.length) jd = result.options[0];
    if (!jd) throw new Error(result.error||'JD not found.');
    var token = getGraphToken_(false);
    var drive = getSharePointDrive_(token);
    var blob = downloadJD_(drive.id, jd, token);
    return {success:true, fileName:blob.getName(), mimeType:'application/pdf', base64:Utilities.base64Encode(blob.getBytes())};
  } catch(e) {
    return {success:false, error:jdFriendlyError(e)};
  }
}

function testJDGraphConnection() {
  var token = getGraphToken_(true);
  SpreadsheetApp.getUi().alert(
    'Microsoft Graph',
    token ? 'Graph authentication successful.' : 'Graph authentication failed.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function testJDSharePoint() {
  try {
    var token = getGraphToken_(false);
    var drive = getSharePointDrive_(token);
    SpreadsheetApp.getUi().alert(
      'SharePoint',
      'SharePoint drive resolved successfully.\n\nDrive: ' + drive.name + '\nID: ' + drive.id,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert('SharePoint', jdFriendlyError(err), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function testGraphConnection() { return testJDGraphConnection(); }
function testSharePointConnection() { return testJDSharePoint(); }


/* ================================================================
 * LOW-LEVEL GRAPH
 * ================================================================ */

function graphRequest_(url, method, token) {
  var response = UrlFetchApp.fetch(url, {
    method: method || 'get',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var text = response.getContentText();
  var data = {};

  try { data = text ? JSON.parse(text) : {}; }
  catch (e) { data = { raw: text }; }

  return { code: code, data: data, raw: text };
}


/* ================================================================
 * CONFIG / HELPERS
 * ================================================================ */

function getConfig_(key) {
  var props = PropertiesService.getScriptProperties();
  var value = props.getProperty(key);
  return value ? String(value).trim() : '';
}

function getRequiredConfig_(key) {
  var value = getConfig_(key);
  if (!value) throw new Error('Script Property "' + key + '" is missing.');
  return value;
}

function findHeaderIndex_(headers, aliases) {
  var normalized = headers.map(normalizeHeader_);
  var wanted = aliases.map(normalizeHeader_);

  for (var i = 0; i < wanted.length; i++) {
    var exact = normalized.indexOf(wanted[i]);
    if (exact >= 0) return exact;
  }
  return -1;
}

function normalizeHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function cleanText_(value) {
  return String(value == null ? '' : value).trim();
}

function cleanPath_(path) {
  return String(path || '').trim().replace(/^\/+/, '').replace(/\/+$/, '');
}

function normalizeSharePointFolderPath_(value) {
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return '';

  var lines = raw.split(/[\r\n;]+/).map(function(s) { return s.trim(); }).filter(Boolean);
  if (lines.length) raw = lines[0];

  raw = raw.replace(/\\/g, '/');

  if (/^https?:\/\//i.test(raw)) {
    try {
      var urlPath = raw.replace(/^https?:\/\/[^/]+/i, '');
      urlPath = decodeURIComponent(urlPath);

      var sitePath = normalizeSitePath_(
        getConfig_('SITE_PATH') || TRIBERA_JD.DEFAULT_SITE_PATH
      ).replace(/\/+$/, '');

      if (sitePath && urlPath.indexOf(sitePath + '/') === 0) {
        urlPath = urlPath.substring(sitePath.length + 1);
      } else if (urlPath.indexOf('/sites/') === 0) {
        var parts = urlPath.split('/').filter(Boolean);
        if (parts.length > 2) urlPath = parts.slice(2).join('/');
      }
      raw = urlPath;
    } catch (e) {}
  }

  var prefixes = [
    /^Eragina Digital Solutions Private Limited\//i,
    /^Eragina Digital - Documents\//i,
    /^Shared Documents\//i,
    /^Documents\//i,
    /^SiteAssets\//i
  ];
  var changed = true;
  while (changed) {
    changed = false;
    for (var p = 0; p < prefixes.length; p++) {
      var next = raw.replace(prefixes[p], '');
      if (next !== raw) { raw = next; changed = true; }
    }
  }

  raw = cleanPath_(raw);
  if (!raw) return '';

  if (/^Requirements\//i.test(raw)) {
    raw = 'Assets/' + raw;
  } else if (!/^Assets\//i.test(raw)) {
    raw = 'Assets/Requirements/' + raw;
  }

  return raw;
}

function buildJDFolderVariants_(rawPath) {
  var variants = [];
  var seen = {};

  function add(p) {
    var cleaned = cleanPath_(String(p || '').replace(/\\/g, '/'));
    if (!cleaned) return;
    var key = cleaned.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    variants.push(cleaned);
  }

  var base = String(rawPath || '').trim();
  if (!base) return variants;

  // Preserve the literal path first. For OneDrive, Documents/JobDescription
  // is often the real drive-root path; for a SharePoint document library the
  // equivalent path is usually JobDescription/.... Do not destroy that
  // distinction before Graph gets a chance to resolve it.
  add(base);
  if (/^Documents\//i.test(base)) add(base.replace(/^Documents\//i, ''));
  if (/^Shared Documents\//i.test(base)) add(base.replace(/^Shared Documents\//i, ''));
  add(normalizeSharePointFolderPath_(base));

  var withoutTrailingJD = base.replace(/[\\\/]\s*JD\s*$/i, '');
  if (withoutTrailingJD !== base) {
    add(normalizeSharePointFolderPath_(withoutTrailingJD));
  }

  var slashed = base.replace(/\\/g, '/');
  var stripPrefixes = [
    /^Eragina Digital Solutions Private Limited\//i,
    /^Eragina Digital - Documents\//i,
    /^Shared Documents\//i,
    /^Documents\//i
  ];

  var stage = slashed;
  var stages = [stage];
  var changedAny = true;
  while (changedAny) {
    changedAny = false;
    for (var i = 0; i < stripPrefixes.length; i++) {
      var next = stage.replace(stripPrefixes[i], '');
      if (next !== stage) { stage = next; stages.push(stage); changedAny = true; }
    }
  }

  stages.forEach(function(s) {
    add(s);
    add('Assets/' + s);
    add('Assets/Requirements/' + s);

    var noJD = s.replace(/[\\\/]\s*JD\s*$/i, '');
    if (noJD !== s) {
      add(noJD);
      add('Assets/' + noJD);
      add('Assets/Requirements/' + noJD);
    }
  });

  return variants;
}

function normalizeSitePath_(path) {
  var clean = String(path || '').trim();
  if (!clean) return '';
  return clean.charAt(0) === '/' ? clean : '/' + clean;
}

function encodeGraphPath_(path) {
  return String(path || '')
    .split('/')
    .map(function(part) { return encodeURIComponent(part); })
    .join('/');
}

function normalizeSearch_(value) {
  return String(value || '').toLowerCase().replace(/[-_\s]+/g, '');
}

function uniqueStrings_(items) {
  var seen = {};
  var result = [];

  (items || []).forEach(function(item) {
    var value = cleanText_(item);
    if (value && !seen[value]) {
      seen[value] = true;
      result.push(value);
    }
  });

  return result;
}

function uniqueNumbers_(items) {
  var seen = {};
  var result = [];

  (items || []).forEach(function(item) {
    var n = Number(item);
    if (isFinite(n) && n >= 2 && !seen[n]) {
      seen[n] = true;
      result.push(n);
    }
  });

  return result;
}

function extractTokens_(text) {
  var result = [];
  text = normalizeTokenSyntax_(String(text || ''));
  var re = /\{([A-Za-z][A-Za-z0-9_]*)\}/g;
  var match;
  while ((match = re.exec(String(text || '')))) {
    var token = match[1];
    if (result.indexOf(token) < 0) result.push(token);
  }
  return result;
}

function normalizeCustomToken_(value) {
  return String(value || '').trim().replace(/^\{/, '').replace(/\}$/, '');
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function parseEmailList_(value){if(!value)return[];var seen={};return String(value).split(/[;,]/).map(function(e){return e.trim().toLowerCase();}).filter(Boolean).filter(function(e){if(seen[e])return false;seen[e]=true;return true;});}

function textToSimpleHtml_(text) {
  var escaped = String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  return escaped
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n/g, '<br>');
}

function getCandidateNameSafely_(row) {
  try { return loadCandidate_(row).name; }
  catch (e) { return 'Row ' + row; }
}

function truncate_(text, max) {
  var s = String(text || '');
  return s.length > max ? s.substring(0, max) + '...' : s;
}

function jdFriendlyError(err) {
  if (!err) return 'Unknown error.';
  return String(err.message || err).replace(/^Exception:\s*/i, '');
}


/* ================================================================
 * DIAGNOSTIC HELPERS
 * ================================================================ */

function clearJDMailerCaches() {
  var userCache = CacheService.getUserCache();
  var scriptCache = CacheService.getScriptCache();

  ['TRIBERA_SP_DRIVE_FINAL', TRIBERA_JD.TOKEN_CACHE, TRIBERA_JD.NOTIFICATION_TOKEN_CACHE].forEach(function(key) {
    try { userCache.remove(key); } catch (e) {}
    try { scriptCache.remove(key); } catch (e) {}
  });

  Logger.log('Tribera JD Mailer caches cleared.');

  try {
    SpreadsheetApp.getUi().alert(
      'JD Mailer',
      'Tribera JD Mailer caches cleared.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) { /* Not in a UI context */ }

  return { success: true, message: 'Tribera JD Mailer caches cleared.' };
}

function debugJDForRow(rowNumber) {
  rowNumber = Number(rowNumber || 2);
  var c = loadCandidate_(rowNumber);

  Logger.log('===== CANDIDATE =====');
  Logger.log(JSON.stringify({
    row: c.row, name: c.name, roleId: c.roleId, role: c.role,
    roleFiles: c.roleFiles, roleSourceSheet: c.roleSourceSheet, roleSourceRow: c.roleSourceRow
  }, null, 2));

  var token = getGraphToken_(false);
  var drive = getSharePointDrive_(token);
  Logger.log('===== DRIVE =====');
  Logger.log(JSON.stringify(drive, null, 2));

  var variants = buildJDFolderVariants_(c.roleFiles || '');
  Logger.log('===== FOLDER VARIANTS =====');
  Logger.log(JSON.stringify(variants, null, 2));

  for (var i = 0; i < variants.length; i++) {
    var folder = getDriveItemByPath_(drive.id, variants[i], token);
    if (folder && folder.folder && folder.id) {
      Logger.log('FOUND: ' + variants[i]);
      var items = listChildrenRecursive_(drive.id, folder.id, token, 5);
      Logger.log(JSON.stringify(items.map(function(it) {
        return {
          name: it.name,
          isFile: !!it.file,
          isFolder: !!it.folder,
          isJD: isJDFile_(it),
          looksLikeResume: looksLikeResume_(it.name)
        };
      }), null, 2));
      break;
    } else {
      Logger.log('MISS : ' + variants[i]);
    }
  }

  var resolved = resolveJD_(c);
  Logger.log('===== resolveJD RESULT =====');
  Logger.log(JSON.stringify(resolved, null, 2));

  return resolved;
}

function debugCandidateRow(rowNumber) {
  rowNumber = Number(rowNumber || 2);

  var sheet = getCandidateSheet_();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var rowValues = sheet.getRange(rowNumber, 1, 1, lastCol).getDisplayValues()[0];

  Logger.log('===== CANDIDATE SHEET =====');
  Logger.log('Sheet name: ' + sheet.getName());
  Logger.log('Last row: ' + sheet.getLastRow());
  Logger.log('Last col: ' + lastCol);

  Logger.log('===== HEADERS vs VALUES for row ' + rowNumber + ' =====');
  for (var c = 0; c < headers.length; c++) {
    Logger.log(
      'col ' + (c + 1) +
      '  |  header: "' + headers[c] + '"' +
      '  |  value: "' + rowValues[c] + '"'
    );
  }

  Logger.log('===== HEADER MATCH CHECK =====');
  Logger.log('name      col index: ' + findHeaderIndex_(headers, TRIBERA_JD.HEADERS.name));
  Logger.log('email     col index: ' + findHeaderIndex_(headers, TRIBERA_JD.HEADERS.email));
  Logger.log('role      col index: ' + findHeaderIndex_(headers, TRIBERA_JD.HEADERS.role));
  Logger.log('roleId    col index: ' + findHeaderIndex_(headers, TRIBERA_JD.HEADERS.roleId));
  Logger.log('customer  col index: ' + findHeaderIndex_(headers, TRIBERA_JD.HEADERS.customer));
}


/* ================================================================
 * TESTING HELPERS
 * ================================================================ */

function getNotificationAccessToken() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(TRIBERA_JD.NOTIFICATION_TOKEN_CACHE);
  if (cached) return cached;

  // Double-check after potential race
  cached = cache.get(TRIBERA_JD.NOTIFICATION_TOKEN_CACHE);
  if (cached) return cached;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Second check inside lock
    cached = cache.get(TRIBERA_JD.NOTIFICATION_TOKEN_CACHE);
    if (cached) return cached;

    var p = PropertiesService.getScriptProperties();
    var base = String(p.getProperty('NOTIFICATION_API_URL') || '').trim();
    var id = String(p.getProperty('NOTIFICATION_CLIENT_ID') || '').trim();
    var secret = String(p.getProperty('NOTIFICATION_CLIENT_SECRET') || '').trim();

    if (!base) throw new Error('NOTIFICATION_API_URL is missing.');
    if (!id) throw new Error('NOTIFICATION_CLIENT_ID is missing.');
    if (!secret) throw new Error('NOTIFICATION_CLIENT_SECRET is missing.');
    base = base.replace(/\/+$/, '');

    var r = UrlFetchApp.fetch(base + '/api/v1/auth/login', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ client_id: id, client_secret: secret }),
      muteHttpExceptions: true
    });
    var code = r.getResponseCode();
    var body = r.getContentText();

    if (code !== 200) throw new Error('Notification API authentication failed. HTTP ' + code + ': ' + body);

    var data;
    try { data = JSON.parse(body); }
    catch (e) { throw new Error('Notification API returned an invalid authentication response.'); }
    if (!data.access_token) throw new Error('Notification API did not return an access token.');

    cache.put(TRIBERA_JD.NOTIFICATION_TOKEN_CACHE, data.access_token, Math.max(60, Math.min(3300, Number(data.expires_in || 1800) - 60)));
    return data.access_token;
  } finally {
    lock.releaseLock();
  }
}

function testNotificationJDSend() {
  var result = sendPreparedEmails({
    items: [
      {
        row: 2,
        candidateName: 'Anirudh Test',
        to: 'anirudhgadhar2004@gmail.com',
        roleId: 'EDS-CAP-TBARD-310826'
      }
    ]
  });
  Logger.log(JSON.stringify(result, null, 2));
}


/* ================================================================
 * HARNESS / EVALS
 * ================================================================ */
function runJdMailerEvals() {
  var allowed = TRIBERA_JD.SYSTEM_TOKENS.slice();
  var expected = ['CandidateName','CandidateEmail','Role','RoleID','CustomerName','TaName','JDFileName'];
  var checks = [];
  checks.push({name:'approved token contract', pass: JSON.stringify(allowed) === JSON.stringify(expected), actual:allowed, expected:expected});
  checks.push({name:'TaName enabled', pass: allowed.indexOf('TaName') >= 0});
  checks.push({name:'TaEmail removed', pass: allowed.indexOf('TaEmail') < 0 && allowed.indexOf('TAEmail') < 0});
  var d = defaultTemplate_();
  checks.push({name:'default template name', pass:d.name === 'Default JD Introduction', actual:d.name});
  checks.push({name:'default template has TaName', pass:/\{TaName\}/.test(d.body)});
  checks.push({name:'default template has no TaEmail', pass:!/(\{TaEmail\}|\{TAEmail\})/.test(d.body)});
  var failed = checks.filter(function(x){ return !x.pass; });
  var result = {success: failed.length === 0, checks:checks, failed:failed};
  Logger.log('TRIBERA JD MAILER EVALS\n' + JSON.stringify(result, null, 2));
  return result;
}


/* ================================================================
 * SHORTCUTS
 * ================================================================ */

function openJDMailerDialog() {
  openSendJD();
}

/**
 * MACRO-SAFE: assign this function to a Google Sheets macro/keyboard shortcut
 * when you want to open the JD Mailer dialog for the currently selected rows.
 * No underscore is used in the function name intentionally.
 */
function openJDMailerShortcut() {
  openSendJD();
}

/**
 * MACRO-SAFE: sends the JD payload directly without opening the JD Mailer UI.
 *
 * Flow:
 * 1. Capture current selected candidate rows.
 * 2. Resolve the JD once BEFORE sending.
 * 3. Use the selected JD as an explicit override so prepareBatchEmail()
 *    NEVER performs a second JD lookup.
 * 4. Prepare + send in the same Apps Script execution.
 *
 * This is the fast keyboard/macro path.
 */
function sendJDPayloadShortcut() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var rows = captureJDSelection();
    if (!rows.length) {
      ss.toast('Select candidate rows first.', 'JD Mailer', 4);
      return { success: false, error: 'No candidate rows are selected.' };
    }

    ss.toast('Preparing JD payload...', 'JD Mailer', 3);

    // Read and validate the selected candidates.
    var candidates = loadCandidates_(rows);
    validateSingleRoleSelection_(candidates);

    if (!candidates.length) {
      throw new Error('No candidate rows were found.');
    }

    // Resolve the JD ONCE for this shortcut execution.
    // Role/Role-ID edits inside the dialog are irrelevant here because this
    // shortcut sends directly from the selected sheet rows.
    var jdResult = resolveJD_(candidates[0]);
    var selected = jdResult && jdResult.options && jdResult.options.length
      ? jdResult.options[0]
      : null;

    if (!selected || !selected.id) {
      throw new Error(
        'Job Description could not be resolved for Role-ID "' +
        String(candidates[0].roleId || '') +
        '" / Role "' +
        String(candidates[0].role || '') +
        '". The payload was not sent.'
      );
    }

    var selectedId = String(selected.id);
    var selectedName = String(selected.name || selected.fileName || '');
    var selectedDriveId = String(selected.driveId || '');

    // validateSingleRoleSelection_ guarantees one Role-ID/Role. Therefore
    // the resolved JD is reused for the entire batch instead of performing
    // another SharePoint lookup for every candidate.
    var templates = getEmailTemplates();
    var lastId = PropertiesService.getUserProperties()
      .getProperty('TRIBERA_JD_LAST_TEMPLATE_ID') || 'default';

    var template =
      templates.filter(function(t) { return String(t.id) === String(lastId); })[0] ||
      templates.filter(function(t) { return String(t.id) === 'default'; })[0] ||
      templates[0];

    if (!template) {
      throw new Error('No email template is available.');
    }

    // Explicit JD override is critical: prepareBatchEmail() will NOT call
    // resolveJD_() again when overrides.jd.id exists.
    var prepared = prepareBatchEmail({
      template: template,
      overrides: {
        jd: {
          id: selectedId,
          name: selectedName,
          fileName: selectedName,
          webUrl: String(selected.webUrl || ''),
          mimeType: String(selected.mimeType || 'application/pdf'),
          driveId: selectedDriveId
        }
      }
    });

    if (!prepared || !prepared.success || !prepared.emails || !prepared.emails.length) {
      throw new Error((prepared && prepared.error) || 'No candidate emails were prepared.');
    }

    var invalid = prepared.emails.filter(function(e) {
      return e.validation && e.validation.valid === false;
    });

    if (invalid.length) {
      var invalidMessage = invalid.slice(0, 5).map(function(e) {
        return (e.candidateName || 'Candidate') +
          ' — ' +
          ((e.validation.errors || []).join('; ') || 'invalid');
      }).join('\n');

      throw new Error(
        'Send blocked. ' + invalid.length +
        ' candidate(s) have validation errors:\n\n' + invalidMessage
      );
    }

    ss.toast('Sending JD payload...', 'JD Mailer', 3);

    // One send execution. Notification mode uses fetchAll(); Graph mode
    // downloads the JD once and uses parallel Graph requests.
    var result = sendPreparedEmails({ items: prepared.emails });
    var mode = getSendMode_();
    var sent = Number(result.sent || result.successful || 0);
    var failed = Number(result.failed || 0);

    var summary =
      sent + ' sent · ' + failed + ' failed | ' +
      'JD: ' + selectedName;

    if (result.timing) summary += ' | ' + result.timing;

    ss.toast(summary, 'JD Mailer', 8);
    return result;

  } catch (e) {
    var message = jdFriendlyError(e);
    ss.toast(message, 'JD Mailer - Send failed', 8);
    return { success: false, error: message };
  }
}

/**
 * Backward-compatible macro name. If this name was already imported as a
 * Sheets macro, it now uses the optimized direct-send implementation.
 */
function quickSendJD() {
  return sendJDPayloadShortcut();
}

function rememberLastTemplate(templateId) {
  try {
    PropertiesService.getUserProperties()
      .setProperty("TRIBERA_JD_LAST_TEMPLATE_ID", String(templateId || "default"));
    return { success: true };
  } catch (e) {
    return { success: false, error: String(e) };
  }
}

function testAllSelectedJDResolution() {
  var rows = getStoredJDSelection().rows;
  Logger.log('===== TRIBERA JD RESOLUTION TEST =====');
  Logger.log('Selected rows: ' + JSON.stringify(rows));

  if (!rows.length) {
    Logger.log('RESULT: No selected rows. Select candidate rows first.');
    return {success:false,total:0,resolved:0,failed:0,results:[]};
  }

  var results = [];
  rows.forEach(function(row) {
    var item = {row:Number(row),candidate:'',roleId:'',role:'',roleFiles:'',success:false,jdFileName:'',optionCount:0,error:''};
    try {
      var c = loadCandidate_(row);
      item.candidate = c.name || '';
      item.roleId = c.roleId || '';
      item.role = c.role || '';
      item.roleFiles = c.roleFiles || '';
      Logger.log('');
      Logger.log('ROW %s | %s | Role-ID=%s | Role=%s | Role Files=%s', row, c.name, c.roleId, c.role, c.roleFiles);

      var r = resolveJD_(c);
      item.success = !!(r && r.selectedName);
      item.jdFileName = r && r.selectedName || '';
      item.optionCount = r && r.options ? r.options.length : 0;
      item.error = r && r.error || '';

      if (item.success) {
        Logger.log('PASS | JD=%s | options=%s', item.jdFileName, item.optionCount);
        (r.options || []).slice(0,10).forEach(function(o, i) {
          Logger.log('  option[%s] %s | %s', i + 1, o.name, o.webUrl || '');
        });
      } else {
        Logger.log('FAIL | %s', item.error || 'No JD selected');
      }
    } catch (e) {
      item.error = jdFriendlyError(e);
      Logger.log('FAIL | %s', item.error);
    }
    results.push(item);
  });

  var resolved = results.filter(function(x){return x.success;}).length;
  var failedCount = results.length - resolved;
  Logger.log('');
  Logger.log('===== SUMMARY =====');
  Logger.log('Resolved: %s/%s', resolved, results.length);
  Logger.log('Failed: %s', failedCount);
  Logger.log(JSON.stringify(results, null, 2));
  return {success:failedCount === 0,total:results.length,resolved:resolved,failed:failedCount,results:results};
}

function debugPreparedEmails() {
  var rows = getStoredJDSelection().rows;
  Logger.log('Selected rows: ' + JSON.stringify(rows));
  Logger.log('');

  if (!rows.length) {
    Logger.log('No rows selected. Select rows in Sep-26, open JD Mailer, then run this again.');
    return;
  }

  rows.forEach(function(row) {
    Logger.log('================ ROW ' + row + ' ================');
    var c = loadCandidate_(row);
    Logger.log('  name:      "' + c.name + '"');
    Logger.log('  roleId:    "' + c.roleId + '"');
    Logger.log('  role:      "' + c.role + '"');
    Logger.log('  roleFiles: "' + c.roleFiles + '"');
    Logger.log('  roleSourceSheet: "' + c.roleSourceSheet + '"');
    Logger.log('  roleSourceRow:   "' + c.roleSourceRow + '"');

    var jd = resolveJD_(c);
    Logger.log('  resolveJD:');
    Logger.log('    selectedName: "' + (jd.selectedName || '') + '"');
    Logger.log('    error:        "' + (jd.error || '') + '"');
    Logger.log('    option count: ' + (jd.options ? jd.options.length : 0));
    if (jd.options && jd.options.length) {
      jd.options.forEach(function(o, i) {
        Logger.log('      [' + i + '] ' + o.name);
      });
    }
    Logger.log('');
  });
}

function debugSendPath() {
  var rows = getStoredJDSelection().rows;
  Logger.log('Selected rows: ' + JSON.stringify(rows));

  var candidates = loadCandidates_(rows);
  Logger.log('Candidates loaded: ' + candidates.length);

  var first = candidates[0];
  Logger.log('');
  Logger.log('Candidate[0]:');
  Logger.log('  name:      ' + first.name);
  Logger.log('  roleId:    ' + first.roleId);
  Logger.log('  roleFiles: ' + first.roleFiles);

  var jdResult = resolveJD_(first);
  Logger.log('');
  Logger.log('resolveJD result for candidate[0]:');
  Logger.log(JSON.stringify(jdResult, null, 2));

  var prepared = prepareBatchEmail({
    template: getEmailTemplates()[0],
    overrides: {},
    draft: null
  });

  Logger.log('');
  Logger.log('prepareBatchEmail output:');
  Logger.log('  batch: ' + JSON.stringify(prepared.batch));
  prepared.emails.forEach(function(e, i) {
    Logger.log('  email[' + i + ']: ' + e.candidateName +
               ' | jd.status=' + (e.jd && e.jd.status) +
               ' | jd.fileName="' + (e.jd && e.jd.fileName || '') + '"');
  });
}

function sendJDFromSheet() {
  var ui = SpreadsheetApp.getUi();

  try {
    // Get currently selected spreadsheet rows
    var rows = captureJDSelection();

    if (!rows.length) {
      ui.alert(
        'JD Mailer',
        'No candidate rows are selected.\n\nSelect the candidates you want to send the JD to.',
        ui.ButtonSet.OK
      );
      return;
    }

    // Get default template
    var template = findTemplate_(
      getEmailTemplates(),
      'default'
    );

    if (!template) {
      throw new Error('Default template not found.');
    }

    // Prepare personalized emails
    var prepared = prepareBatchEmail({
      template: template
    });

    var emails = prepared.emails || [];

    if (!emails.length) {
      throw new Error('No candidate emails were prepared.');
    }

    // Check for invalid candidates
    var invalidItems = emails.filter(function(item) {
      return item.validation &&
        item.validation.valid === false;
    });

    if (invalidItems.length) {
      var invalidMessage = invalidItems.map(function(item) {
        return '• ' +
          (item.candidateName || 'Candidate') +
          ' <' +
          (item.to || '') +
          '>\n  ' +
          ((item.validation.errors || []).join('; ') || 'Invalid data');
      }).join('\n\n');

      ui.alert(
        'JD Mailer',
        'Email was NOT sent.\n\n' +
        invalidItems.length +
        ' candidate(s) have issues:\n\n' +
        invalidMessage,
        ui.ButtonSet.OK
      );

      return;
    }

    // Send emails
    var result = sendPreparedEmails({
      items: emails
    });

    var sent = result.sent || result.successful || 0;
    var failed = result.failed || 0;

    // Build confirmation message
    var message = '';

    message += 'JD Mailer - Send Result\n';
    message += '────────────────────────\n\n';

    message += 'Total selected: ' + emails.length + '\n';
    message += 'Successfully sent: ' + sent + '\n';
    message += 'Failed: ' + failed + '\n\n';

    message += 'Recipients:\n\n';

    (result.results || []).forEach(function(item) {

      if (item.success) {
        message +=
          '✓ SENT\n' +
          '  ' + (item.name || 'Candidate') + '\n' +
          '  ' + (item.email || '') + '\n\n';

      } else {
        message +=
          '✗ FAILED\n' +
          '  ' + (item.name || 'Candidate') + '\n' +
          '  ' + (item.email || '') + '\n' +
          '  Reason: ' + (item.reason || item.details || 'Unknown error') +
          '\n\n';
      }

    });

    ui.alert(
      'JD Mailer',
      message,
      ui.ButtonSet.OK
    );

  } catch (error) {

    ui.alert(
      'JD Mailer',
      'Send failed:\n\n' +
      jdFriendlyError(error),
      ui.ButtonSet.OK
    );
  }
}


/* ================================================================
 * CONCURRENCY TEST HELPERS
 * ================================================================ */

/**
 * Simulates concurrent sends from multiple "users" by running
 * prepareBatchEmail + sendPreparedEmails in parallel via separate
 * script executions (using UrlFetchApp to call self).
 *
 * Run this function to test: testConcurrentSends_()
 */
function testConcurrentSends_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var testResults = {
    startTime: new Date(),
    simulations: [],
    summary: {}
  };

  try {
    // Get current selection
    var rows = captureJDSelection();
    if (!rows.length) {
      throw new Error('No rows selected. Select candidate rows first.');
    }

    // Load candidates and validate
    var candidates = loadCandidates_(rows);
    validateSingleRoleSelection_(candidates);

    // Resolve JD once (same as shortcut does)
    var jdResult = resolveJD_(candidates[0]);
    var selected = jdResult && jdResult.options && jdResult.options.length ? jdResult.options[0] : null;
    if (!selected || !selected.id) {
      throw new Error('JD could not be resolved for test.');
    }

    var selectedJD = {
      id: String(selected.id),
      name: String(selected.name || selected.fileName || ''),
      fileName: String(selected.name || selected.fileName || ''),
      webUrl: String(selected.webUrl || ''),
      mimeType: String(selected.mimeType || 'application/pdf'),
      driveId: String(selected.driveId || '')
    };

    var template = getEmailTemplates()[0]; // default

    // Prepare base payload
    var basePayload = {
      template: template,
      overrides: {
        jd: selectedJD
      }
    };

    var prepared = prepareBatchEmail(basePayload);
    if (!prepared || !prepared.success || !prepared.emails.length) {
      throw new Error('Prepare failed: ' + (prepared && prepared.error));
    }

    Logger.log('=== CONCURRENCY TEST SETUP ===');
    Logger.log('Selected rows: ' + JSON.stringify(rows));
    Logger.log('Candidates: ' + prepared.emails.length);
    Logger.log('JD: ' + selectedJD.name);
    Logger.log('Mode: ' + getSendMode_());

    // Simulate N concurrent users by calling sendPreparedEmails directly
    // Each "user" gets a copy of the prepared emails
    var CONCURRENT_USERS = 5;
    var sendPromises = [];

    for (var u = 0; u < CONCURRENT_USERS; u++) {
      var userLabel = 'User' + (u + 1);
      var userStart = new Date().getTime();

      // Run send in this same execution (serial but measures lock contention)
      // For true parallel test, use testConcurrentSendsViaWebApp_() below
      try {
        var result = sendPreparedEmails({ items: prepared.emails });
        var userElapsed = new Date().getTime() - userStart;

        testResults.simulations.push({
          user: userLabel,
          success: result.success,
          sent: result.sent || result.successful || 0,
          failed: result.failed || 0,
          elapsedMs: userElapsed,
          timing: result.timing,
          error: null
        });

        Logger.log(userLabel + ': ' + (result.success ? 'OK' : 'FAIL') +
          ' | sent=' + (result.sent || result.successful || 0) +
          ' | failed=' + (result.failed || 0) +
          ' | ' + userElapsed + 'ms');
      } catch (e) {
        var userElapsed = new Date().getTime() - userStart;
        testResults.simulations.push({
          user: userLabel,
          success: false,
          sent: 0,
          failed: prepared.emails.length,
          elapsedMs: userElapsed,
          timing: '',
          error: jdFriendlyError(e)
        });
        Logger.log(userLabel + ': ERROR - ' + jdFriendlyError(e));
      }
    }

    testResults.endTime = new Date();
    testResults.totalMs = testResults.endTime - testResults.startTime;

    var totalSent = testResults.simulations.reduce(function(sum, s) { return sum + s.sent; }, 0);
    var totalFailed = testResults.simulations.reduce(function(sum, s) { return sum + s.failed; }, 0);
    var allOk = testResults.simulations.every(function(s) { return s.success; });

    testResults.summary = {
      concurrentUsers: CONCURRENT_USERS,
      totalCandidatesPerUser: prepared.emails.length,
      totalSent: totalSent,
      totalFailed: totalFailed,
      allSucceeded: allOk,
      totalTimeMs: testResults.totalMs,
      avgTimePerUserMs: Math.round(testResults.totalMs / CONCURRENT_USERS)
    };

    Logger.log('');
    Logger.log('=== CONCURRENCY TEST SUMMARY ===');
    Logger.log(JSON.stringify(testResults.summary, null, 2));
    Logger.log('');

    // Check mail log for row count
    var logSheet = ss.getSheetByName(TRIBERA_JD.LOG_SHEET);
    if (logSheet) {
      var logRows = logSheet.getLastRow() - 1; // minus header
      Logger.log('Mail log rows written: ' + logRows);
      Logger.log('Expected rows: ' + (CONCURRENT_USERS * prepared.emails.length));
      testResults.summary.logRowsWritten = logRows;
      testResults.summary.logRowsExpected = CONCURRENT_USERS * prepared.emails.length;
    }

    return testResults;

  } catch (err) {
    testResults.endTime = new Date();
    testResults.error = jdFriendlyError(err);
    Logger.log('TEST ERROR: ' + testResults.error);
    return testResults;
  }
}

/**
 * TRUE PARALLEL TEST: Deploys as Web App and calls itself via HTTP
 * to simulate real concurrent users from different sessions.
 *
 * Prerequisite: Deploy as Web App (Anyone, even anonymous)
 * Set WEB_APP_URL below to your deployed URL.
 */
function testConcurrentSendsViaWebApp_() {
  // REPLACE WITH YOUR DEPLOYED WEB APP URL
  var WEB_APP_URL = 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec';

  if (WEB_APP_URL === 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec') {
    throw new Error('Set WEB_APP_URL in the function to your deployed Web App URL.');
  }

  var rows = captureJDSelection();
  if (!rows.length) throw new Error('Select candidate rows first.');

  var candidates = loadCandidates_(rows);
  validateSingleRoleSelection_(candidates);

  var jdResult = resolveJD_(candidates[0]);
  var selected = jdResult && jdResult.options && jdResult.options.length ? jdResult.options[0] : null;
  if (!selected || !selected.id) throw new Error('JD not resolved.');

  var template = getEmailTemplates()[0];
  var prepared = prepareBatchEmail({
    template: template,
    overrides: { jd: { id: selected.id, name: selected.name, fileName: selected.name, webUrl: selected.webUrl || '', mimeType: selected.mimeType || 'application/pdf', driveId: selected.driveId || '' } }
  });

  if (!prepared || !prepared.success) throw new Error('Prepare failed.');

  var CONCURRENT_USERS = 5;
  var payload = {
    template: prepared.template,
    overrides: { jd: prepared.emails[0].jd },
    draft: null
  };

  Logger.log('=== PARALLEL WEB APP TEST ===');
  Logger.log('Users: ' + CONCURRENT_USERS);
  Logger.log('Candidates each: ' + prepared.emails.length);
  Logger.log('URL: ' + WEB_APP_URL);

  var requests = [];
  for (var u = 0; u < CONCURRENT_USERS; u++) {
    requests.push({
      url: WEB_APP_URL,
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  }

  var start = new Date().getTime();
  var responses = UrlFetchApp.fetchAll(requests);
  var totalMs = new Date().getTime() - start;

  var results = responses.map(function(r, i) {
    var code = r.getResponseCode();
    var body = r.getContentText();
    var parsed = null;
    try { parsed = JSON.parse(body); } catch (e) {}
    return {
      user: 'User' + (i + 1),
      httpCode: code,
      success: parsed && parsed.success,
      sent: parsed && (parsed.sent || parsed.successful || 0),
      failed: parsed && (parsed.failed || 0),
      error: parsed && !parsed.success ? (parsed.error || body) : null
    };
  });

  Logger.log('Total time: ' + totalMs + 'ms');
  Logger.log('Results: ' + JSON.stringify(results, null, 2));

  var allOk = results.every(function(r) { return r.success; });
  var totalSent = results.reduce(function(s, r) { return s + (r.sent || 0); }, 0);

  return {
    success: allOk,
    totalTimeMs: totalMs,
    users: CONCURRENT_USERS,
    candidatesPerUser: prepared.emails.length,
    totalSent: totalSent,
    results: results
  };
}

/**
 * Test token refresh under concurrent load
 * Clears cache then fires multiple token requests
 */
function testTokenRefreshConcurrency_() {
  clearJDMailerCaches();

  var ITERATIONS = 10;
  var start = new Date().getTime();
  var tokens = [];

  for (var i = 0; i < ITERATIONS; i++) {
    var t0 = new Date().getTime();
    var token = getGraphToken_(false);
    var elapsed = new Date().getTime() - t0;
    tokens.push({ iteration: i + 1, elapsedMs: elapsed, tokenPrefix: token ? token.substring(0, 20) : 'NULL' });
    Logger.log('Iteration ' + (i + 1) + ': ' + elapsed + 'ms');
  }

  var totalMs = new Date().getTime() - start;
  var uniqueTokens = [...new Set(tokens.map(function(t) { return t.tokenPrefix; }))].length;

  Logger.log('Total: ' + totalMs + 'ms | Unique tokens: ' + uniqueTokens + ' (should be 1)');

  return {
    iterations: ITERATIONS,
    totalMs: totalMs,
    avgMs: Math.round(totalMs / ITERATIONS),
    uniqueTokens: uniqueTokens,
    details: tokens
  };
}

/**
 * Test mail log batch write under concurrent load
 */
function testLogWriteConcurrency_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRIBERA_JD.LOG_SHEET) || ss.insertSheet(TRIBERA_JD.LOG_SHEET);
  var beforeRows = sheet.getLastRow();

  var ITERATIONS = 20;
  var items = [];
  var results = [];

  // Create dummy items
  for (var i = 0; i < 3; i++) {
    items.push({
      row: 2 + i,
      candidateName: 'Test Candidate ' + i,
      to: 'test' + i + '@example.com',
      roleId: 'TEST-ROLE-' + i,
      customer: 'Test Customer',
      subject: 'Test Subject ' + i,
      jd: { fileName: 'Test_JD_' + i + '.pdf' }
    });
    results.push({
      success: true,
      jdFileName: 'Test_JD_' + i + '.pdf',
      status: 'Sent',
      messageId: 'msg-' + i
    });
  }

  var start = new Date().getTime();
  for (var iter = 0; iter < ITERATIONS; iter++) {
    writeMailLogsBatch_(items, results);
  }
  var totalMs = new Date().getTime() - start;

  var afterRows = sheet.getLastRow();
  var written = afterRows - beforeRows;
  var expected = ITERATIONS * items.length;

  Logger.log('Iterations: ' + ITERATIONS);
  Logger.log('Items per iter: ' + items.length);
  Logger.log('Total time: ' + totalMs + 'ms');
  Logger.log('Rows written: ' + written + ' / Expected: ' + expected);
  Logger.log('Avg per write: ' + (totalMs / ITERATIONS).toFixed(1) + 'ms');

  return {
    iterations: ITERATIONS,
    itemsPerIter: items.length,
    totalMs: totalMs,
    rowsWritten: written,
    rowsExpected: expected,
    match: written === expected,
    avgMsPerWrite: totalMs / ITERATIONS
  };
}

/**
 * Run all concurrency tests
 */
function runAllConcurrencyTests() {
  var allResults = {};

  Logger.log('========== TEST 1: Token Refresh Concurrency ==========');
  allResults.tokenRefresh = testTokenRefreshConcurrency_();

  Logger.log('');
  Logger.log('========== TEST 2: Log Write Concurrency ==========');
  allResults.logWrite = testLogWriteConcurrency_();

  Logger.log('');
  Logger.log('========== TEST 3: Serial Send Concurrency (same execution) ==========');
  allResults.serialSend = testConcurrentSends_();

  Logger.log('');
  Logger.log('========== ALL TESTS COMPLETE ==========');
  Logger.log(JSON.stringify(allResults, null, 2));

  return allResults;
}

/**
 * DEBUG: Shows full prepared payload for current selection
 * Run after selecting candidate rows in Sheet
 */
function debugPreparedPayload_() {
  var rows = captureJDSelection();
  if (!rows.length) { Logger.log('ERROR: Select candidate rows in Sheet first'); return; }

  var candidates = loadCandidates_(rows);
  validateSingleRoleSelection_(candidates);

  var jdResult = resolveJD_(candidates[0]);
  var selected = jdResult && jdResult.options && jdResult.options.length ? jdResult.options[0] : null;
  if (!selected || !selected.id) { Logger.log('ERROR: JD not resolved - ' + (jdResult && jdResult.error)); return; }

  var template = getEmailTemplates()[0];
  var prepared = prepareBatchEmail({
    template: template,
    overrides: { jd: { id: selected.id, name: selected.name, fileName: selected.name, webUrl: selected.webUrl||'', mimeType: selected.mimeType||'application/pdf', driveId: selected.driveId||'' } }
  });

  Logger.log('=== PAYLOAD FOR SEND ===');
  Logger.log('Mode: ' + getSendMode_());
  Logger.log('Candidates: ' + prepared.emails.length);
  Logger.log('JD: ' + selected.name);
  Logger.log('');
  Logger.log('FIRST EMAIL FULL PAYLOAD:');
  Logger.log(JSON.stringify(prepared.emails[0], null, 2));
  Logger.log('');
  Logger.log('ALL EMAILS SUMMARY:');
  prepared.emails.forEach(function(e, i) {
    Logger.log((i+1) + '. ' + e.candidateName + ' | ' + e.email + ' | ' + e.roleId + ' | JD: ' + e.jd?.fileName);
  });

  return prepared;
}

/**
 * CONCURRENCY TEST: Shows each user's payload + result
 * Run after selecting candidate rows in Sheet
 */
function testConcurrentSendsWithPayload_() {
  var rows = captureJDSelection();
  if (!rows.length) { Logger.log('ERROR: Select candidate rows in Sheet first'); return; }

  var candidates = loadCandidates_(rows);
  validateSingleRoleSelection_(candidates);

  var jdResult = resolveJD_(candidates[0]);
  var selected = jdResult && jdResult.options && jdResult.options.length ? jdResult.options[0] : null;
  if (!selected || !selected.id) { Logger.log('ERROR: JD not resolved - ' + (jdResult && jdResult.error)); return; }

  var template = getEmailTemplates()[0];
  var prepared = prepareBatchEmail({
    template: template,
    overrides: { jd: { id: selected.id, name: selected.name, fileName: selected.name, webUrl: selected.webUrl||'', mimeType: selected.mimeType||'application/pdf', driveId: selected.driveId||'' } }
  });

  var CONCURRENT_USERS = 3;
  Logger.log('=== STARTING ' + CONCURRENT_USERS + ' CONCURRENT SENDS ===');
  Logger.log('Candidates per user: ' + prepared.emails.length);
  Logger.log('JD: ' + selected.name);
  Logger.log('Mode: ' + getSendMode_());
  Logger.log('');

  // Log first user's full payload as sample
  Logger.log('SAMPLE PAYLOAD (User1):');
  Logger.log(JSON.stringify(prepared.emails[0], null, 2));
  Logger.log('');

  var results = [];
  for (var u = 0; u < CONCURRENT_USERS; u++) {
    var label = 'User' + (u+1);
    var start = Date.now();
    try {
      var result = sendPreparedEmails({ items: prepared.emails });
      var ms = Date.now() - start;
      Logger.log(label + ': SUCCESS | ' + (result.sent||result.successful||0) + ' sent | ' + (result.failed||0) + ' failed | ' + ms + 'ms');
      results.push({user: label, success: true, sent: result.sent||result.successful||0, failed: result.failed||0, ms: ms});
    } catch (e) {
      var ms = Date.now() - start;
      Logger.log(label + ': FAILED | ' + e.message + ' | ' + ms + 'ms');
      results.push({user: label, success: false, error: e.message, ms: ms});
    }
  }

  var allOk = results.every(function(r){return r.success;});
  Logger.log('');
  Logger.log('=== SUMMARY ===');
  Logger.log('All succeeded: ' + allOk);
  Logger.log('Total sent: ' + results.reduce(function(s,r){return s+(r.sent||0);},0));
  Logger.log('Total failed: ' + results.reduce(function(s,r){return s+(r.failed||0);},0));

  // Check log sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(TRIBERA_JD.LOG_SHEET);
  if (logSheet) {
    var written = logSheet.getLastRow() - 1;
    Logger.log('Log rows written: ' + written + ' / Expected: ' + (CONCURRENT_USERS * prepared.emails.length));
  }
  return results;
}

function testSendPayload(payload) {
  var executionId = Utilities.getUuid();

  Logger.log('========================================');
  Logger.log('JD MAILER TEST PAYLOAD');
  Logger.log('Execution ID: ' + executionId);
  Logger.log('========================================');

  Logger.log(JSON.stringify(payload, null, 2));

  Logger.log('========================================');
  Logger.log('NO EMAIL SENT - TEST ONLY');
  Logger.log('========================================');

  return {
    success: true,
    executionId: executionId,
    message: 'Payload received successfully. No email was sent.'
  };
}

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);

  Logger.log('==============================');
  Logger.log('PAYLOAD RECEIVED');
  Logger.log('Time: ' + new Date().toISOString());
  Logger.log(JSON.stringify(payload, null, 2));
  Logger.log('==============================');

  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      received: true,
      requestId: Utilities.getUuid(),
      rows: payload.selectedRows || []
    }))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * TEST MULTIPLE USERS AT THE SAME TIME
 *
 * This sends 5 requests concurrently to the deployed
 * Apps Script Web App.
 *
 * NO EMAIL IS SENT.
 */
/**
 * ================================================================
 * CONCURRENT SELECTED-CANDIDATE REQUEST TEST
 * ================================================================
 *
 * Select exactly 2 or 3 candidate rows in the Sheet and run:
 *   JD Mailer -> Test 2-3 Concurrent Requests
 *
 * This test:
 * 1. Uses the real selected Sheet rows.
 * 2. Uses the real template preparation flow.
 * 3. Resolves the real JD before preparing the test.
 * 4. Creates one outbound request per selected candidate.
 * 5. Sends all requests together with fetchAll().
 * 6. NEVER calls Notification/AWS.
 * 7. NEVER sends a real email.
 * 8. NEVER prints the request payload.
 *
 * The Execution log shows:
 * - request number
 * - candidate
 * - email
 * - sheet row
 * - HTTP status
 * - whether a response was received
 * - response size
 * - total elapsed time
 *
 * IMPORTANT:
 * This proves concurrent outbound requests from this Apps Script
 * execution. It does not create fake USER_A/USER_B identities.
 * Actual independent Google users create independent executions.
 */
function testMultipleUsersSelectedCandidates() {
  var ui = SpreadsheetApp.getUi();

  try {
    /*
     * HOW THIS TEST WORKS
     * -------------------
     * 1. You select the candidates in the Sheet exactly as you normally do.
     * 2. If you select 2 candidates, EVERY simulated user request contains
     *    those same 2 candidate objects inside ONE JSON array.
     * 3. If you select 3 candidates, EVERY simulated user request contains
     *    those same 3 candidate objects inside ONE JSON array.
     * 4. We simulate 2 or 3 different users by creating 2 or 3 separate
     *    HTTP requests and sending them concurrently with fetchAll().
     *
     * IMPORTANT:
     * This does not create fake Google accounts or fake Apps Script users.
     * It tests the exact outbound request shape and concurrent HTTP handling.
     */

    var rows = captureJDSelection();

    if (!rows || rows.length < 1 || rows.length > 4) {
      throw new Error(
        'Select 1 to 4 candidate rows in the Sheet, then run this test.'
      );
    }

    var userCountResponse = ui.prompt(
      'Test Multiple Users',
      'How many users should be simulated? Enter 2 or 3.',
      ui.ButtonSet.OK_CANCEL
    );

    if (userCountResponse.getSelectedButton() !== ui.Button.OK) {
      return {
        cancelled: true
      };
    }

    var userCount = Number(
      String(userCountResponse.getResponseText() || '').trim()
    );

    if (userCount !== 2 && userCount !== 3) {
      throw new Error('Enter only 2 or 3 for the number of simulated users.');
    }

    var initialData = buildInitialDialogData_(rows);

    if (!initialData || !initialData.success) {
      throw new Error(
        (initialData && initialData.error) ||
        'Unable to prepare the selected candidates.'
      );
    }

    /*
     * Use the same template and JD preparation path as the real mailer.
     * No candidate names/emails are hardcoded anywhere in this test.
     */
    var templates = initialData.templates || [];
    var template = templates.filter(function(t) {
      return String(t.id) === 'default';
    })[0];

    if (!template && templates.length) {
      template = templates[0];
    }

    if (!template) {
      throw new Error('No email template is available.');
    }

    /*
     * buildInitialDialogData_ resolves the JD before the dialog opens.
     * Reuse that resolved JD for the test instead of resolving it again.
     */
    var firstJD =
      initialData.jdResults &&
      initialData.jdResults.length
        ? initialData.jdResults[0]
        : null;

    if (
      !firstJD ||
      firstJD.status !== 'found' ||
      !firstJD.fileId
    ) {
      throw new Error(
        'Job Description could not be resolved for the selected candidates.'
      );
    }

    var jd = {
      id: String(firstJD.fileId),
      name: String(firstJD.fileName || ''),
      fileName: String(firstJD.fileName || ''),
      webUrl: String(firstJD.webUrl || ''),
      mimeType: String(firstJD.mimeType || 'application/pdf'),
      driveId: String(firstJD.driveId || '')
    };

    var prepared = prepareBatchEmail({
      template: JSON.parse(JSON.stringify(template)),
      overrides: {
        jd: jd
      },
      draft: null
    });

    if (
      !prepared ||
      !prepared.success ||
      !prepared.emails ||
      !prepared.emails.length
    ) {
      throw new Error(
        (prepared && prepared.error) ||
        'No candidate emails were prepared.'
      );
    }

    if (prepared.emails.length !== rows.length) {
      throw new Error(
        'Prepared candidate count (' +
        prepared.emails.length +
        ') does not match selected row count (' +
        rows.length +
        ').'
      );
    }

    /*
     * Build the EXACT business payload once.
     *
     * For 2+ selected candidates this becomes:
     *
     * [
     *   { candidate_name: "...", ... },
     *   { candidate_name: "...", ... }
     * ]
     *
     * That same array is used as the body of EACH simulated user's
     * independent HTTP request.
     */
    var candidatePayloads = prepared.emails.map(function(item) {
      return buildNotificationPayload_(item);
    });

    var outboundPayload = candidatePayloads.length === 1
      ? candidatePayloads[0]
      : candidatePayloads;

    var endpoint = 'https://httpbin.org/post';

    /*
     * Create one HTTP request per simulated user.
     *
     * User 1 -> [candidate1, candidate2]
     * User 2 -> [candidate1, candidate2]
     * User 3 -> [candidate1, candidate2]
     *
     * With 3 selected candidates the same becomes:
     *
     * User 1 -> [candidate1, candidate2, candidate3]
     * User 2 -> [candidate1, candidate2, candidate3]
     * User 3 -> [candidate1, candidate2, candidate3]
     */
    var requests = [];

    for (var userIndex = 1; userIndex <= userCount; userIndex++) {
      requests.push({
        url: endpoint,
        method: 'post',
        contentType: 'application/json',
        headers: {
          Accept: 'application/json'
        },
        payload: JSON.stringify(outboundPayload),
        muteHttpExceptions: true
      });
    }

    Logger.log('');
    Logger.log('============================================================');
    Logger.log('TRIBERA JD MAILER - MULTI USER JSON REQUEST TEST');
    Logger.log('============================================================');
    Logger.log('Endpoint: ' + endpoint);
    Logger.log('Sheet: ' + getCandidateSheet_().getName());
    Logger.log('Selected rows: ' + JSON.stringify(rows));
    Logger.log('Selected candidates: ' + candidatePayloads.length);
    Logger.log('Simulated users: ' + userCount);
    Logger.log('HTTP requests sent: ' + requests.length);
    Logger.log(
      'Payload shape: ' +
      (candidatePayloads.length === 1 ? 'OBJECT' : 'ARRAY')
    );
    Logger.log(
      'Each user request contains the same selected candidates: YES'
    );
    Logger.log('Real Notification/AWS API called: NO');
    Logger.log('Real emails sent: 0');
    Logger.log('------------------------------------------------------------');

    /*
     * Print the request payload exactly as it will be sent.
     * We intentionally print it once here and then again beside every
     * response so it is easy to verify REQUEST 1 / REQUEST 2 / REQUEST 3.
     */
    Logger.log('BASE PAYLOAD USED BY ALL SIMULATED USERS:');
    Logger.log(JSON.stringify(outboundPayload, null, 2));
    Logger.log('------------------------------------------------------------');

    var started = Date.now();
    var responses;

    try {
      /*
       * fetchAll() dispatches the simulated user requests concurrently.
       */
      responses = UrlFetchApp.fetchAll(requests);
    } catch (err) {
      throw new Error(
        'Concurrent multi-user HTTP test failed: ' + jdFriendlyError(err)
      );
    }

    var elapsedMs = Date.now() - started;
    var successful = 0;
    var failed = 0;
    var results = [];

    for (var i = 0; i < responses.length; i++) {
      var response = responses[i];
      var httpCode = response.getResponseCode();
      var responseBody = response.getContentText() || '';
      var parsedResponse = null;

      try {
        parsedResponse = responseBody
          ? JSON.parse(responseBody)
          : null;
      } catch (e) {
        parsedResponse = null;
      }

      var ok = httpCode >= 200 && httpCode < 300;

      if (ok) {
        successful++;
      } else {
        failed++;
      }

      Logger.log('');
      Logger.log('============================================================');
      Logger.log('RESPONSE ' + (i + 1));
      Logger.log('============================================================');
      Logger.log('Simulated user: USER_' + (i + 1));
      Logger.log('HTTP request number: ' + (i + 1));
      Logger.log('HTTP status: ' + httpCode);
      Logger.log('Candidate count in payload: ' + candidatePayloads.length);
      Logger.log(
        'Payload shape: ' +
        (candidatePayloads.length === 1 ? 'OBJECT' : 'ARRAY')
      );
      Logger.log('REQUEST ' + (i + 1) + ' PAYLOAD:');
      Logger.log(JSON.stringify(outboundPayload, null, 2));
      Logger.log('RESPONSE ' + (i + 1) + ' BODY:');

      if (parsedResponse !== null) {
        /*
         * httpbin echoes the request body under json.
         * Printing the whole response lets you verify that the server
         * received the exact same object/array.
         */
        Logger.log(JSON.stringify(parsedResponse, null, 2));
      } else {
        Logger.log(responseBody || '(empty response)');
      }

      Logger.log('------------------------------------------------------------');

      results.push({
        success: ok,
        user: 'USER_' + (i + 1),
        requestNumber: i + 1,
        candidateCount: candidatePayloads.length,
        payloadShape: candidatePayloads.length === 1
          ? 'object'
          : 'array',
        httpCode: httpCode,
        responseSize: responseBody.length
      });
    }

    Logger.log('');
    Logger.log('============================================================');
    Logger.log('MULTI USER TEST SUMMARY');
    Logger.log('============================================================');
    Logger.log('Selected candidates: ' + candidatePayloads.length);
    Logger.log('Simulated users: ' + userCount);
    Logger.log('HTTP requests: ' + requests.length);
    Logger.log('Successful responses: ' + successful);
    Logger.log('Failed responses: ' + failed);
    Logger.log('Elapsed time: ' + elapsedMs + ' ms');
    Logger.log('Real Notification/AWS API called: NO');
    Logger.log('Real emails sent: 0');
    Logger.log('============================================================');

    ui.alert(
      'JD Mailer Multi-User Test',
      'Test completed successfully.\\n\\n' +
      'Selected candidates: ' + candidatePayloads.length + '\\n' +
      'Simulated users: ' + userCount + '\\n' +
      'HTTP requests: ' + requests.length + '\\n' +
      'Successful: ' + successful + '\\n' +
      'Failed: ' + failed + '\\n' +
      'Time: ' + elapsedMs + ' ms\\n\\n' +
      '1 candidate = JSON object.\\n' +
      '2+ candidates = one JSON array per user request.\\n\\n' +
      'Open Apps Script -> Executions to see REQUEST 1/2/3 payloads and responses.\\n' +
      'No real email was sent.',
      ui.ButtonSet.OK
    );

    return {
      success: failed === 0,
      selectedCandidateCount: candidatePayloads.length,
      simulatedUserCount: userCount,
      requestCount: requests.length,
      payloadShape: candidatePayloads.length === 1
        ? 'object'
        : 'array',
      successful: successful,
      failed: failed,
      elapsedMs: elapsedMs,
      payload: outboundPayload,
      results: results,
      realEmailSent: false,
      realNotificationApiCalled: false
    };

  } catch (err) {
    ui.alert(
      'JD Mailer Multi-User Test',
      jdFriendlyError(err),
      ui.ButtonSet.OK
    );
    throw err;
  }
}

/*
 * Backward-compatible aliases.
 * Keep these public names so existing macros/buttons do not break.
 */
function testSelectedCandidatesTwoOrThreeRequests() {
  return testMultipleUsersSelectedCandidates();
}

function testConcurrentRequestsNoDeployment() {
  return testMultipleUsersSelectedCandidates();
}

function testMultiUserPayloadsNoDeployment() {
  return testMultipleUsersSelectedCandidates();
}

function testMultipleUsersAtSameTime() {
  return testMultipleUsersSelectedCandidates();
}

