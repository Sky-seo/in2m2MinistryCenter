const SHEETS = {
  MEMBERS: 'Members',
  POSITIONS: 'Positions',
  AVAILABILITY: 'Availability',
  SUBMISSION_STATUS: 'Submission_Status',
  ASSIGNMENTS: 'Assignments',
  SETTINGS: 'Settings',
  EMAIL_LOG: 'Email_Log',
};

const HEADERS = {
  [SHEETS.MEMBERS]: ['Member_ID', 'Name', 'Email', 'Position', 'Active', 'Join_Date', 'Notes'],
  [SHEETS.POSITIONS]: ['Position_ID', 'Position_Name', 'Required_Per_Sunday', 'Active', 'Sort_Order'],
  [SHEETS.AVAILABILITY]: ['Submission_ID', 'Month', 'Member_ID', 'Name', 'Email', 'Position', 'Available_Sundays', 'Submitted_At', 'Notes'],
  [SHEETS.SUBMISSION_STATUS]: ['Month', 'Member_ID', 'Name', 'Position', 'Status', 'Available_Sundays', 'Submitted_At'],
  [SHEETS.ASSIGNMENTS]: ['Date'],
  [SHEETS.SETTINGS]: ['Key', 'Value'],
  [SHEETS.EMAIL_LOG]: ['Sent_At', 'Email_Type', 'Month', 'Recipient_Name', 'Recipient_Email', 'Status'],
};

const DEFAULT_SETTINGS = {
  Admin_Email: '',
  Current_Target_Month: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM'),
  Submission_Deadline: '',
  Web_App_URL: '',
  Church_Name: '',
  Team_Name: '진행팀',
};

function doGet(e) {
  const page = resolvePage_(e);
  const templateName = page === 'admin' ? 'Admin' : 'Index';
  const template = HtmlService.createTemplateFromFile(templateName);
  template.appUrl = ScriptApp.getService().getUrl();
  return template
    .evaluate()
    .setTitle('진행팀 Scheduler')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function resolvePage_(e) {
  const rawPage = (e && e.parameter && e.parameter.page) || (e && e.pathInfo) || 'index';
  const page = String(rawPage).replace(/^\/+/, '').replace(/\.html$/i, '').toLowerCase();
  return page === 'admin' ? 'admin' : 'index';
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEETS.MEMBERS) return;

  const editedCol = e.range.getColumn();
  const editedRow = e.range.getRow();
  if (editedRow <= 1) return;

  const nameCol = HEADERS[SHEETS.MEMBERS].indexOf('Name') + 1;
  if (editedCol !== nameCol) return;

  const value = typeof e.value === 'string' ? e.value.trim() : e.range.getValue();
  if (!value) return;

  const activeCol = HEADERS[SHEETS.MEMBERS].indexOf('Active') + 1;
  sheet.getRange(editedRow, activeCol).setValue('Yes');
}

function setupWorkbook() {
  ensureSetup(true);
  return getAppData();
}

function ensureSetup(forceHeaders) {
  ensureSheets_(Object.keys(HEADERS), forceHeaders);
  seedSettings_();
  seedPositions_();
  ensureMemberDefaults_();
  syncSubmissionStatus();
  refreshAssignmentHeaders_();
}

function ensurePublicSetup_(forceHeaders) {
  ensureSheets_([SHEETS.MEMBERS, SHEETS.SETTINGS], forceHeaders);
  seedSettings_();
}

function ensurePublicSubmissionSetup_(forceHeaders) {
  ensureSheets_([SHEETS.MEMBERS, SHEETS.SETTINGS, SHEETS.AVAILABILITY, SHEETS.SUBMISSION_STATUS], forceHeaders);
  seedSettings_();
  ensureMemberDefaults_();
}

function ensureAdminPrimarySetup_(forceHeaders) {
  ensureSheets_([SHEETS.MEMBERS, SHEETS.POSITIONS, SHEETS.AVAILABILITY, SHEETS.ASSIGNMENTS, SHEETS.SETTINGS], forceHeaders);
  seedSettings_();
  seedPositions_();
  ensureMemberDefaults_();
}

function ensureSheets_(sheetNames, forceHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetNames.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    ensureHeaders_(sheet, HEADERS[sheetName], forceHeaders);
  });
}

function getAppData() {
  ensureSetup();
  const settings = getSettings();
  const members = getMembers();
  const positions = getPositions();
  const activeMembers = members.filter(isActive_);
  const activePositions = positions.filter(isActive_);
  const targetMonth = getActiveTargetMonth_(settings);
  const sundays = getPublicSundays_();
  return {
    settings,
    members,
    positions,
    activeMembers,
    activePositions,
    targetMonth,
    sundays,
    status: getSubmissionStatus(targetMonth),
    availabilityRows: getAvailabilityMatrixRows_(targetMonth),
    assignments: getAssignments(targetMonth),
    appUrl: ScriptApp.getService().getUrl(),
  };
}

function getAdminPrimaryData() {
  ensureAdminPrimarySetup_();
  const settings = getSettings();
  const members = getMembers();
  const positions = getPositions();
  const activeMembers = members.filter(isActive_);
  const activePositions = positions.filter(isActive_);
  const targetMonth = getActiveTargetMonth_(settings);
  const sundays = getVisibleSundays_(targetMonth);
  const submissions = getRows_(SHEETS.AVAILABILITY).filter((row) => normalizeMonth_(row.Month) === targetMonth);
  const statusRows = buildSubmissionStatusRows_(activeMembers, submissions, targetMonth);
  return {
    settings,
    members,
    positions,
    activeMembers,
    activePositions,
    targetMonth,
    sundays,
    status: statusRows,
    availabilityRows: statusRows,
    assignments: getAssignmentsForPositions_(targetMonth, activePositions, sundays),
    appUrl: ScriptApp.getService().getUrl(),
  };
}

function getAdminStatusData() {
  ensureAdminPrimarySetup_();
  const settings = getSettings();
  const members = getMembers();
  const positions = getPositions();
  const activeMembers = members.filter(isActive_);
  const targetMonth = getActiveTargetMonth_(settings);
  const sundays = getVisibleSundays_(targetMonth);
  const submissions = getRows_(SHEETS.AVAILABILITY).filter((row) => normalizeMonth_(row.Month) === targetMonth);
  const statusRows = buildSubmissionStatusRows_(activeMembers, submissions, targetMonth);
  return {
    settings,
    members,
    positions,
    activeMembers,
    activePositions: positions.filter(isActive_),
    targetMonth,
    sundays,
    status: statusRows,
    availabilityRows: statusRows,
    appUrl: ScriptApp.getService().getUrl(),
  };
}

function getAdminAssignmentsData(month) {
  ensureSheets_([SHEETS.POSITIONS, SHEETS.ASSIGNMENTS, SHEETS.SETTINGS], false);
  const settings = getSettings();
  const positions = getPositions();
  const activePositions = positions.filter(isActive_);
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(settings));
  const sundays = getVisibleSundays_(targetMonth);
  return {
    positions,
    activePositions,
    targetMonth,
    sundays,
    assignments: getAssignmentsForPositions_(targetMonth, activePositions, sundays),
  };
}

function getAdminSecondaryData() {
  ensureAdminPrimarySetup_();
  const settings = getSettings();
  const members = getMembers();
  const positions = getPositions();
  return {
    settings,
    members,
    positions,
    activeMembers: members.filter(isActive_),
    activePositions: positions.filter(isActive_),
    appUrl: ScriptApp.getService().getUrl(),
  };
}

function getPublicData() {
  const cached = CacheService.getScriptCache().get('publicData:v1');
  if (cached) return JSON.parse(cached);

  const members = getPublicMembersFast_();
  const targetMonth = getCurrentMonth_();
  const data = {
    activeMembers: members.filter(isActive_).map((member) => Object.assign({}, member, {
      Member_ID: member.Member_ID || member.Name,
    })),
    targetMonth,
    sundays: getPublicSundays_(),
  };
  CacheService.getScriptCache().put('publicData:v1', JSON.stringify(data), 120);
  return data;
}

function getMemberAvailability(month, memberId) {
  ensurePublicSubmissionSetup_();
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings()));
  const member = getMembers().find((item) => item.Member_ID === memberId && isActive_(item));
  if (!member) return null;
  const submission = getRows_(SHEETS.AVAILABILITY)
    .find((row) => normalizeMonth_(row.Month) === targetMonth && sameMemberSubmission_(row, member));
  return submission || null;
}

function getSettings() {
  const values = getRows_(SHEETS.SETTINGS);
  return values.reduce((settings, row) => {
    if (row.Key) settings[row.Key] = row.Value;
    return settings;
  }, Object.assign({}, DEFAULT_SETTINGS));
}

function saveSettings(settings) {
  ensureSetup();
  const merged = Object.assign(getSettings(), settings || {});
  const rows = Object.keys(DEFAULT_SETTINGS).map((key) => [key, merged[key] || '']);
  const sheet = getSheet_(SHEETS.SETTINGS);
  sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).clearContent();
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  syncSubmissionStatus();
  refreshAssignmentHeaders_();
  return getAppData();
}

function getMembers() {
  return getRows_(SHEETS.MEMBERS);
}

function getPublicMembersFast_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MEMBERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(1, 1, sheet.getLastRow(), HEADERS[SHEETS.MEMBERS].length).getValues();
  const headers = values[0];
  return values.slice(1)
    .filter((row) => row.some((cell) => cell !== ''))
    .map((row) => headers.reduce((object, header, index) => {
      object[header] = row[index] instanceof Date ? formatDateCell_(row[index], header) : row[index];
      return object;
    }, {}));
}

function saveMember(member) {
  ensureSetup();
  const sheet = getSheet_(SHEETS.MEMBERS);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const id = member.Member_ID || createId_('M');
  const record = {
    Member_ID: id,
    Name: clean_(member.Name),
    Email: clean_(member.Email),
    Position: clean_(member.Position),
    Active: member.Active === 'No' ? 'No' : 'Yes',
    Join_Date: member.Join_Date || today_(),
    Notes: clean_(member.Notes),
  };
  validateRequired_(record.Name, '이름을 입력해 주세요.');
  validateRequired_(record.Position, '포지션을 선택해 주세요.');

  const rowIndex = findRowIndex_(rows, headers.indexOf('Member_ID'), id);
  if (rowIndex > -1) {
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([headers.map((key) => record[key] || '')]);
  } else {
    sheet.appendRow(headers.map((key) => record[key] || ''));
  }
  syncSubmissionStatus();
  clearPublicCache_();
  return getAppData();
}

function setMemberActive(memberId, active) {
  return updateMemberField_(memberId, 'Active', active ? 'Yes' : 'No');
}

function deleteMember(memberId) {
  ensureSetup();
  const sheet = getSheet_(SHEETS.MEMBERS);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = findRowIndex_(rows, headers.indexOf('Member_ID'), memberId);
  if (rowIndex > -1) sheet.deleteRow(rowIndex + 1);
  syncSubmissionStatus();
  clearPublicCache_();
  return getAppData();
}

function getPositions() {
  return getRows_(SHEETS.POSITIONS).sort((a, b) => Number(a.Sort_Order || 0) - Number(b.Sort_Order || 0));
}

function savePosition(position) {
  ensureSetup();
  const sheet = getSheet_(SHEETS.POSITIONS);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const id = position.Position_ID || createId_('P');
  const record = {
    Position_ID: id,
    Position_Name: clean_(position.Position_Name),
    Required_Per_Sunday: Number(position.Required_Per_Sunday || 1),
    Active: position.Active === 'No' ? 'No' : 'Yes',
    Sort_Order: Number(position.Sort_Order || getPositions().length + 1),
  };
  validateRequired_(record.Position_Name, '포지션 이름을 입력해 주세요.');

  const rowIndex = findRowIndex_(rows, headers.indexOf('Position_ID'), id);
  if (rowIndex > -1) {
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([headers.map((key) => record[key] || '')]);
  } else {
    sheet.appendRow(headers.map((key) => record[key] || ''));
  }
  refreshAssignmentHeaders_();
  syncSubmissionStatus();
  return getAppData();
}

function setPositionActive(positionId, active) {
  ensureSetup();
  const sheet = getSheet_(SHEETS.POSITIONS);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = findRowIndex_(rows, headers.indexOf('Position_ID'), positionId);
  if (rowIndex > -1) sheet.getRange(rowIndex + 1, headers.indexOf('Active') + 1).setValue(active ? 'Yes' : 'No');
  refreshAssignmentHeaders_();
  syncSubmissionStatus();
  return getAppData();
}

function deletePosition(positionId) {
  ensureSetup();
  const sheet = getSheet_(SHEETS.POSITIONS);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = findRowIndex_(rows, headers.indexOf('Position_ID'), positionId);
  if (rowIndex > -1) sheet.deleteRow(rowIndex + 1);
  refreshAssignmentHeaders_();
  syncSubmissionStatus();
  return getAppData();
}

function submitAvailability(payload) {
  ensurePublicSubmissionSetup_();
  const settings = getSettings();
  const month = normalizeMonth_(payload.Month || getActiveTargetMonth_(settings));
  const member = getMembers().find((item) => (
    item.Member_ID === payload.Member_ID || normalizeName_(item.Name) === normalizeName_(payload.Member_ID)
  ) && isActive_(item));
  if (!member) throw new Error('활성 멤버를 찾을 수 없습니다.');

  const selected = Array.isArray(payload.Available_Sundays) ? payload.Available_Sundays : [];
  const sheet = getSheet_(SHEETS.AVAILABILITY);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const memberCol = headers.indexOf('Member_ID');
  const nameCol = headers.indexOf('Name');
  const monthCol = headers.indexOf('Month');
  const availableCol = headers.indexOf('Available_Sundays');
  sheet.getRange(1, monthCol + 1, sheet.getMaxRows(), 1).setNumberFormat('@');
  sheet.getRange(1, availableCol + 1, sheet.getMaxRows(), 1).setNumberFormat('@');
  const existingIndex = rows.findIndex((row, index) => {
    if (index === 0 || normalizeMonth_(row[monthCol]) !== month) return false;
    return row[memberCol] === member.Member_ID || normalizeName_(row[nameCol]) === normalizeName_(member.Name);
  });
  const record = {
    Submission_ID: existingIndex > -1 ? rows[existingIndex][headers.indexOf('Submission_ID')] : createId_('S'),
    Month: month,
    Member_ID: member.Member_ID,
    Name: member.Name,
    Email: member.Email,
    Position: member.Position,
    Available_Sundays: selected.join(', '),
    Submitted_At: timestamp_(),
    Notes: clean_(payload.Notes),
  };

  if (existingIndex > -1) {
    sheet.getRange(existingIndex + 1, 1, 1, headers.length).setValues([headers.map((key) => record[key] || '')]);
  } else {
    sheet.appendRow(headers.map((key) => record[key] || ''));
  }
  syncSubmissionStatus(month);
  return { success: true, message: '제출이 완료되었습니다.', record };
}

function getSubmissionStatus(month) {
  syncSubmissionStatus(month);
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings()));
  return getRows_(SHEETS.SUBMISSION_STATUS).filter((row) => normalizeMonth_(row.Month) === targetMonth);
}

function syncSubmissionStatus(month) {
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings()));
  const members = getMembers().filter(isActive_);
  const submissions = getRows_(SHEETS.AVAILABILITY).filter((row) => normalizeMonth_(row.Month) === targetMonth);
  const existingOtherMonths = getRows_(SHEETS.SUBMISSION_STATUS).filter((row) => normalizeMonth_(row.Month) !== targetMonth);
  const statusRows = buildSubmissionStatusRows_(members, submissions, targetMonth);
  writeObjects_(SHEETS.SUBMISSION_STATUS, existingOtherMonths.concat(statusRows));
  return statusRows;
}

function getAvailabilityMatrixRows_(month) {
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings()));
  const members = getMembers().filter(isActive_);
  const submissions = getRows_(SHEETS.AVAILABILITY).filter((row) => normalizeMonth_(row.Month) === targetMonth);
  return buildSubmissionStatusRows_(members, submissions, targetMonth);
}

function buildSubmissionStatusRows_(members, submissions, targetMonth) {
  return members.map((member) => {
    const submission = submissions.find((row) => sameMemberSubmission_(row, member));
    return {
      Month: targetMonth,
      Member_ID: member.Member_ID,
      Name: member.Name,
      Position: member.Position,
      Status: submission ? 'Submitted' : 'Not Submitted',
      Available_Sundays: submission ? submission.Available_Sundays : '',
      Submitted_At: submission ? submission.Submitted_At : '',
    };
  });
}

function getAssignments(month) {
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings()));
  const activePositions = getPositions().filter(isActive_);
  const sundays = getVisibleSundays_(targetMonth);
  return getAssignmentsForPositions_(targetMonth, activePositions, sundays);
}

function getAssignmentsForPositions_(targetMonth, activePositions, sundays) {
  const visibleMonths = getMonthsForDates_(sundays);
  const rows = getRows_(SHEETS.ASSIGNMENTS).filter((row) => visibleMonths.indexOf(normalizeMonth_(row.Date)) > -1);
  return sundays.map((date) => {
    const existing = rows.find((row) => row.Date === date) || { Date: date };
    activePositions.forEach((position) => {
      if (existing[position.Position_Name] === undefined) existing[position.Position_Name] = '';
    });
    return existing;
  });
}

function saveAssignments(payload) {
  ensureSetup();
  refreshAssignmentHeaders_();
  const incoming = payload.Assignments || [];
  const incomingMonths = getMonthsForDates_(incoming.map((row) => row.Date));
  const allRows = getRows_(SHEETS.ASSIGNMENTS);
  const keepRows = allRows.filter((row) => incomingMonths.indexOf(normalizeMonth_(row.Date)) === -1);
  writeObjects_(SHEETS.ASSIGNMENTS, keepRows.concat(incoming));
  return getAppData();
}

function generateLineup(month) {
  ensureSetup();
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings()));
  const sundays = getVisibleSundays_(targetMonth);
  const positions = getPositions().filter(isActive_);
  const members = getMembers().filter(isActive_);
  const visibleMonths = getMonthsForDates_(sundays);
  const submissions = getRows_(SHEETS.AVAILABILITY).filter((row) => visibleMonths.indexOf(normalizeMonth_(row.Month)) > -1);
  const assignmentCounts = {};
  members.forEach((member) => assignmentCounts[member.Member_ID] = 0);

  const assignments = sundays.map((date) => {
    const row = { Date: date };
    positions.forEach((position) => {
      const needed = Number(position.Required_Per_Sunday || 1);
      const candidates = members
        .filter((member) => member.Position === position.Position_Name)
        .filter((member) => {
          const submission = submissions.find((item) => sameMemberSubmission_(item, member));
          return submission && parseDateList_(submission.Available_Sundays).indexOf(date) > -1;
        })
        .sort((a, b) => (assignmentCounts[a.Member_ID] || 0) - (assignmentCounts[b.Member_ID] || 0));
      const selected = candidates.slice(0, needed);
      selected.forEach((member) => assignmentCounts[member.Member_ID] += 1);
      row[position.Position_Name] = selected.map((member) => member.Name).join(', ');
    });
    return row;
  });
  saveAssignments({ Month: targetMonth, Assignments: assignments });
  return getAppData();
}

function sendReminderEmails(month) {
  ensureSetup();
  const settings = getSettings();
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(settings));
  const webAppUrl = settings.Web_App_URL || ScriptApp.getService().getUrl();
  const notSubmitted = getSubmissionStatus(targetMonth).filter((row) => row.Status !== 'Submitted');
  const memberMap = getMembers().reduce((map, member) => {
    map[member.Member_ID] = member;
    return map;
  }, {});
  const logRows = [];

  notSubmitted.forEach((status) => {
    const member = memberMap[status.Member_ID];
    if (!member || !member.Email) return;
    try {
      MailApp.sendEmail({
        to: member.Email,
        subject: `[진행팀 Scheduler] ${targetMonth} 가능 주일 제출 부탁드립니다`,
        htmlBody: [
          `<p>${member.Name}님 안녕하세요.</p>`,
          `<p>${targetMonth} 진행팀 가능 주일 제출을 부탁드립니다.</p>`,
          `<p><a href="${webAppUrl}">제출하러 가기</a></p>`,
        ].join(''),
      });
      logRows.push([timestamp_(), 'Reminder', targetMonth, member.Name, member.Email, 'Sent']);
    } catch (error) {
      logRows.push([timestamp_(), 'Reminder', targetMonth, member.Name, member.Email, `Failed: ${error.message}`]);
    }
  });

  if (logRows.length) {
    getSheet_(SHEETS.EMAIL_LOG).getRange(getSheet_(SHEETS.EMAIL_LOG).getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
  }
  return { sent: logRows.filter((row) => row[5] === 'Sent').length, total: notSubmitted.length, logs: logRows };
}

function getSundaysForMonth(month) {
  const parts = String(month || '').split('-');
  const year = Number(parts[0]);
  const monthIndex = Number(parts[1]) - 1;
  if (!year || monthIndex < 0) return [];
  const date = new Date(year, monthIndex, 1);
  const sundays = [];
  while (date.getMonth() === monthIndex) {
    if (date.getDay() === 0) sundays.push(Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    date.setDate(date.getDate() + 1);
  }
  return sundays;
}

function getPublicSundays_() {
  const currentMonth = getCurrentMonth_();
  const nextMonth = getNextMonth_(currentMonth);
  return getSundaysForMonth(currentMonth).concat(getSundaysForMonth(nextMonth));
}

function getVisibleSundays_(month) {
  const targetMonth = normalizeMonth_(month || getCurrentMonth_());
  return getSundaysForMonth(targetMonth).concat(getSundaysForMonth(getNextMonth_(targetMonth)));
}

function getMonthsForDates_(dates) {
  return dates
    .map((date) => normalizeMonth_(date))
    .filter((month, index, months) => month && months.indexOf(month) === index);
}

function getActiveTargetMonth_(settings) {
  return normalizeMonth_((settings && settings.Current_Target_Month) || getCurrentMonth_()) || getCurrentMonth_();
}

function getCurrentMonth_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
}

function getNextMonth_(month) {
  const parts = String(month || '').split('-');
  const year = Number(parts[0]);
  const monthIndex = Number(parts[1]) - 1;
  if (!year || monthIndex < 0) return getCurrentMonth_();
  return Utilities.formatDate(new Date(year, monthIndex + 1, 1), Session.getScriptTimeZone(), 'yyyy-MM');
}

function refreshAssignmentHeaders_() {
  const sheet = getSheet_(SHEETS.ASSIGNMENTS);
  const activePositions = getPositions().filter(isActive_).map((position) => position.Position_Name);
  const headers = ['Date'].concat(activePositions);
  ensureHeaders_(sheet, headers, true);
}

function seedSettings_() {
  const sheet = getSheet_(SHEETS.SETTINGS);
  if (sheet.getLastRow() > 1) return;
  const rows = Object.keys(DEFAULT_SETTINGS).map((key) => [key, DEFAULT_SETTINGS[key]]);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function seedPositions_() {
  const sheet = getSheet_(SHEETS.POSITIONS);
  if (sheet.getLastRow() > 1) return;
  ['자막', '조명', '음향', '카메라', '믹싱', 'PD', 'FD'].forEach((name, index) => {
    sheet.appendRow([createId_('P'), name, 1, 'Yes', index + 1]);
  });
}

function ensureMemberDefaults_() {
  const sheet = getSheet_(SHEETS.MEMBERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headers = getHeaders_(sheet);
  const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const idCol = headers.indexOf('Member_ID');
  const nameCol = headers.indexOf('Name');
  const activeCol = headers.indexOf('Active');
  const joinDateCol = headers.indexOf('Join_Date');
  let changed = false;

  rows.forEach((row) => {
    if (nameCol > -1 && idCol > -1 && !row[nameCol] && row[idCol] && !isGeneratedId_(row[idCol], 'M')) {
      row[nameCol] = row[idCol];
      row[idCol] = '';
      changed = true;
    }

    if (!row[nameCol]) return;
    if (idCol > -1 && !row[idCol]) {
      row[idCol] = createId_('M');
      changed = true;
    }
    if (activeCol > -1 && !row[activeCol]) {
      row[activeCol] = 'Yes';
      changed = true;
    }
    if (joinDateCol > -1 && !row[joinDateCol]) {
      row[joinDateCol] = today_();
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    clearPublicCache_();
  }
}

function updateMemberField_(memberId, field, value) {
  ensureSetup();
  const sheet = getSheet_(SHEETS.MEMBERS);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = findRowIndex_(rows, headers.indexOf('Member_ID'), memberId);
  if (rowIndex > -1) sheet.getRange(rowIndex + 1, headers.indexOf(field) + 1).setValue(value);
  syncSubmissionStatus();
  clearPublicCache_();
  return getAppData();
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error(`${name} sheet is missing. Run setupWorkbook first.`);
  return sheet;
}

function ensureHeaders_(sheet, headers, force) {
  const headerWidth = Math.max(sheet.getLastColumn(), headers.length, 1);
  const current = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
  const missing = headers.some((header, index) => current[index] !== header);
  if (force || sheet.getLastRow() === 0 || missing) {
    sheet.getRange(1, 1, 1, headerWidth).clearContent();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#edf2f7');
  }
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(Boolean);
}

function getRows_(sheetName) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];
  const headers = getHeaders_(sheet);
  return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
    .filter((row) => row.some((cell) => cell !== ''))
    .map((row) => {
      return headers.reduce((object, header, index) => {
        object[header] = row[index] instanceof Date
          ? formatDateCell_(row[index], header)
          : row[index];
        return object;
      }, {});
    });
}

function formatDateCell_(value, header) {
  const dateOnlyHeaders = ['Date', 'Available_Sundays'];
  return Utilities.formatDate(
    value,
    Session.getScriptTimeZone(),
    dateOnlyHeaders.indexOf(header) > -1 ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm'
  );
}

function parseDateList_(value) {
  if (value instanceof Date) return [Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd')];
  return String(value || '').split(',').map((item) => item.trim()).filter(Boolean);
}

function writeObjects_(sheetName, objects) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  if (!objects.length) return;
  const rows = objects.map((object) => headers.map((header) => object[header] || ''));
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function findRowIndex_(rows, colIndex, value) {
  return rows.findIndex((row, index) => index > 0 && row[colIndex] === value);
}

function createId_(prefix) {
  return `${prefix}-${Utilities.getUuid().slice(0, 8).toUpperCase()}`;
}

function isGeneratedId_(value, prefix) {
  return String(value || '').indexOf(`${prefix}-`) === 0;
}

function clean_(value) {
  return String(value || '').trim();
}

function isActive_(record) {
  return String(record.Active || 'Yes').trim() !== 'No';
}

function sameMemberSubmission_(submission, member) {
  return submission.Member_ID === member.Member_ID || normalizeName_(submission.Name) === normalizeName_(member.Name);
}

function normalizeName_(value) {
  return String(value || '').trim().toLowerCase();
}

function clearPublicCache_() {
  CacheService.getScriptCache().remove('publicData:v1');
}

function validateRequired_(value, message) {
  if (!String(value || '').trim()) throw new Error(message);
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function timestamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function normalizeMonth_(dateString) {
  if (dateString instanceof Date) {
    return Utilities.formatDate(dateString, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return String(dateString || '').slice(0, 7);
}
