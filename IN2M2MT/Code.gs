const SHEETS = {
  MEMBERS: 'Members',
  POSITIONS: 'Positions',
  AVAILABILITY: 'Availability',
  SUBMISSION_STATUS: 'Submission_Status',
  ASSIGNMENTS: 'Assignments',
  SETTINGS: 'Settings',
  EMAIL_LOG: 'Email_Log',
};

const TEAMS = {
  PRAISE: 'praise',
  PRODUCTION: 'production',
};

const TEAM_LABELS = {
  [TEAMS.PRAISE]: '찬양팀',
  [TEAMS.PRODUCTION]: '진행팀',
};

const TEAM_SHEET_PREFIXES = {
  [TEAMS.PRAISE]: 'PR-',
  [TEAMS.PRODUCTION]: 'OP-',
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
  Admin_Name: '',
  Team_Name: '진행팀',
  Reminder_Message: '가능한 주일 제출을 아직 완료하지 않으셨습니다. 아래 링크에서 제출 부탁드립니다.',
};

function doGet(e) {
  const page = resolvePage_(e);
  const templateName = page === 'admin' ? 'Admin' : page === 'submission' ? 'Submission' : 'Index';
  const template = HtmlService.createTemplateFromFile(templateName);
  template.appUrl = ScriptApp.getService().getUrl();
  return template
    .evaluate()
    .setTitle('IN2M2 Ministry Center')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function resolvePage_(e) {
  const rawPage = (e && e.parameter && e.parameter.page) || (e && e.pathInfo) || 'index';
  const page = String(rawPage).replace(/^\/+/, '').replace(/\.html$/i, '').toLowerCase();
  if (page === 'admin') return 'admin';
  if (page === 'submission') return 'submission';
  return 'entry';
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const baseName = baseSheetName_(sheet.getName());
  if (baseName !== SHEETS.MEMBERS && baseName !== SHEETS.POSITIONS) return;

  const editedCol = e.range.getColumn();
  const editedRow = e.range.getRow();
  if (editedRow <= 1) return;

  const value = typeof e.value === 'string' ? e.value.trim() : e.range.getValue();
  if (!value) return;

  if (baseName === SHEETS.MEMBERS) {
    const nameCol = HEADERS[SHEETS.MEMBERS].indexOf('Name') + 1;
    if (editedCol === nameCol) ensureMemberRowDefaults_(sheet, editedRow);
  }

  if (baseName === SHEETS.POSITIONS) {
    const nameCol = HEADERS[SHEETS.POSITIONS].indexOf('Position_Name') + 1;
    if (editedCol === nameCol) ensurePositionRowDefaults_(sheet, editedRow);
  }
}

function setupWorkbook() {
  ensureSetup(true);
  return getAppData();
}

function ensureSetup(forceHeaders) {
  ensureSheets_(Object.keys(HEADERS), forceHeaders, TEAMS.PRODUCTION);
  ensureSheets_(teamDataSheetNames_(), forceHeaders, TEAMS.PRAISE);
  seedSettings_();
  seedSettings_(TEAMS.PRAISE);
  seedPositions_();
  seedPositions_(TEAMS.PRAISE);
  ensureMemberDefaults_();
  ensurePositionDefaults_();
  ensureMemberDefaults_(TEAMS.PRAISE);
  ensurePositionDefaults_(TEAMS.PRAISE);
  syncSubmissionStatus();
  syncSubmissionStatus(null, TEAMS.PRAISE);
  refreshAssignmentHeaders_();
  refreshAssignmentHeaders_(TEAMS.PRAISE);
}

function ensurePublicSetup_(forceHeaders) {
  ensureSheets_([SHEETS.MEMBERS, SHEETS.SETTINGS], forceHeaders, TEAMS.PRODUCTION);
  ensureSheets_([SHEETS.MEMBERS, SHEETS.SETTINGS], forceHeaders, TEAMS.PRAISE);
  seedSettings_();
  seedSettings_(TEAMS.PRAISE);
}

function ensurePublicSubmissionSetup_(forceHeaders, team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureSheets_([SHEETS.MEMBERS, SHEETS.SETTINGS, SHEETS.AVAILABILITY, SHEETS.SUBMISSION_STATUS], forceHeaders, normalizedTeam);
  seedSettings_(normalizedTeam);
  ensureMemberDefaults_(normalizedTeam);
}

function ensureAdminPrimarySetup_(forceHeaders, team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureSheets_([SHEETS.MEMBERS, SHEETS.POSITIONS, SHEETS.AVAILABILITY, SHEETS.ASSIGNMENTS, SHEETS.SETTINGS, SHEETS.SUBMISSION_STATUS, SHEETS.EMAIL_LOG], forceHeaders, normalizedTeam);
  seedSettings_(normalizedTeam);
  seedPositions_(normalizedTeam);
  ensureMemberDefaults_(normalizedTeam);
  ensurePositionDefaults_(normalizedTeam);
}

function teamDataSheetNames_() {
  return Object.keys(HEADERS);
}

function ensureSheets_(sheetNames, forceHeaders, team) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetNames.forEach((sheetName) => {
    const actualSheetName = teamSheet_(sheetName, team);
    const sheet = ss.getSheetByName(actualSheetName) || ss.insertSheet(actualSheetName);
    ensureHeaders_(sheet, HEADERS[baseSheetName_(sheetName)], forceHeaders);
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

function getAdminPrimaryData(team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureAdminPrimarySetup_(false, normalizedTeam);
  const settings = getSettings(normalizedTeam);
  const members = getMembers(normalizedTeam);
  const positions = getPositions(normalizedTeam);
  const activeMembers = members.filter(isActive_);
  const activePositions = positions.filter(isActive_);
  const targetMonth = getActiveTargetMonth_(settings);
  const sundays = getVisibleSundays_(targetMonth);
  const submissions = getRows_(SHEETS.AVAILABILITY, normalizedTeam).filter((row) => normalizeMonth_(row.Month) === targetMonth);
  const statusRows = buildSubmissionStatusRows_(activeMembers, submissions, targetMonth);
  return {
    team: normalizedTeam,
    teamLabel: TEAM_LABELS[normalizedTeam],
    settings,
    members,
    positions,
    activeMembers,
    activePositions,
    targetMonth,
    sundays,
    status: statusRows,
    availabilityRows: statusRows,
    assignments: getAssignmentsForPositions_(targetMonth, activePositions, sundays, normalizedTeam),
    appUrl: ScriptApp.getService().getUrl(),
  };
}

function getAdminStatusData(team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureAdminPrimarySetup_(false, normalizedTeam);
  const settings = getSettings(normalizedTeam);
  const members = getMembers(normalizedTeam);
  const positions = getPositions(normalizedTeam);
  const activeMembers = members.filter(isActive_);
  const targetMonth = getActiveTargetMonth_(settings);
  const sundays = getVisibleSundays_(targetMonth);
  const submissions = getRows_(SHEETS.AVAILABILITY, normalizedTeam).filter((row) => normalizeMonth_(row.Month) === targetMonth);
  const statusRows = buildSubmissionStatusRows_(activeMembers, submissions, targetMonth);
  return {
    team: normalizedTeam,
    teamLabel: TEAM_LABELS[normalizedTeam],
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

function getAdminAssignmentsData(month, team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureSheets_([SHEETS.POSITIONS, SHEETS.ASSIGNMENTS, SHEETS.SETTINGS], false, normalizedTeam);
  const settings = getSettings(normalizedTeam);
  const positions = getPositions(normalizedTeam);
  const activePositions = positions.filter(isActive_);
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(settings));
  const sundays = getVisibleSundays_(targetMonth);
  return {
    team: normalizedTeam,
    teamLabel: TEAM_LABELS[normalizedTeam],
    positions,
    activePositions,
    targetMonth,
    sundays,
    assignments: getAssignmentsForPositions_(targetMonth, activePositions, sundays, normalizedTeam),
  };
}

function getAdminSecondaryData(team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureAdminPrimarySetup_(false, normalizedTeam);
  const settings = getSettings(normalizedTeam);
  const members = getMembers(normalizedTeam);
  const positions = getPositions(normalizedTeam);
  return {
    team: normalizedTeam,
    teamLabel: TEAM_LABELS[normalizedTeam],
    settings,
    members,
    positions,
    activeMembers: members.filter(isActive_),
    activePositions: positions.filter(isActive_),
    appUrl: ScriptApp.getService().getUrl(),
  };
}

function getPublicData(team) {
  const normalizedTeam = normalizeTeam_(team);
  const cacheKey = `publicData:v2:${normalizedTeam}`;
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (cached) return JSON.parse(cached);

  ensureSheets_([SHEETS.MEMBERS], false, normalizedTeam);
  const members = getPublicMembersFast_(normalizedTeam);
  const targetMonth = getCurrentMonth_();
  const data = {
    team: normalizedTeam,
    teamLabel: TEAM_LABELS[normalizedTeam],
    activeMembers: members.filter(isActive_).map((member) => Object.assign({}, member, {
      Member_ID: member.Member_ID || member.Name,
    })),
    targetMonth,
    sundays: getPublicSundays_(),
  };
  CacheService.getScriptCache().put(cacheKey, JSON.stringify(data), 120);
  return data;
}

function getMemberAvailability(month, memberId, team) {
  const normalizedTeam = normalizeTeam_(team);
  ensurePublicSubmissionSetup_(false, normalizedTeam);
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings(normalizedTeam)));
  const member = getMembers(normalizedTeam).find((item) => item.Member_ID === memberId && isActive_(item));
  if (!member) return null;
  const submission = getRows_(SHEETS.AVAILABILITY, normalizedTeam)
    .find((row) => normalizeMonth_(row.Month) === targetMonth && sameMemberSubmission_(row, member));
  return submission || null;
}

function getSettings(team) {
  const values = getRows_(SHEETS.SETTINGS, team);
  return values.reduce((settings, row) => {
    if (row.Key) settings[row.Key] = row.Value;
    return settings;
  }, Object.assign({}, DEFAULT_SETTINGS));
}

function saveSettings(settings, team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureAdminPrimarySetup_(false, normalizedTeam);
  const merged = Object.assign(getSettings(normalizedTeam), settings || {});
  const rows = Object.keys(DEFAULT_SETTINGS).map((key) => [key, merged[key] || '']);
  const sheet = getSheet_(SHEETS.SETTINGS, normalizedTeam);
  sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).clearContent();
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  syncSubmissionStatus(null, normalizedTeam);
  refreshAssignmentHeaders_(normalizedTeam);
  return getAdminPrimaryData(normalizedTeam);
}

function getMembers(team) {
  return getRows_(SHEETS.MEMBERS, team);
}

function getPublicMembersFast_(team) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(teamSheet_(SHEETS.MEMBERS, team));
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
  const team = normalizeTeam_(member && member.Team);
  ensureAdminPrimarySetup_(false, team);
  const sheet = getSheet_(SHEETS.MEMBERS, team);
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
  syncSubmissionStatus(null, team);
  clearPublicCache_();
  return getAdminPrimaryData(team);
}

function setMemberActive(memberId, active, team) {
  return updateMemberField_(memberId, 'Active', active ? 'Yes' : 'No', team);
}

function deleteMember(memberId, team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureAdminPrimarySetup_(false, normalizedTeam);
  const sheet = getSheet_(SHEETS.MEMBERS, normalizedTeam);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = findRowIndex_(rows, headers.indexOf('Member_ID'), memberId);
  if (rowIndex > -1) sheet.deleteRow(rowIndex + 1);
  syncSubmissionStatus(null, normalizedTeam);
  clearPublicCache_();
  return getAdminPrimaryData(normalizedTeam);
}

function getPositions(team) {
  return getRows_(SHEETS.POSITIONS, team).sort((a, b) => Number(a.Sort_Order || 0) - Number(b.Sort_Order || 0));
}

function savePosition(position) {
  const team = normalizeTeam_(position && position.Team);
  ensureAdminPrimarySetup_(false, team);
  const sheet = getSheet_(SHEETS.POSITIONS, team);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const id = position.Position_ID || createId_('P');
  const record = {
    Position_ID: id,
    Position_Name: clean_(position.Position_Name),
    Required_Per_Sunday: Number(position.Required_Per_Sunday || 1),
    Active: position.Active === 'No' ? 'No' : 'Yes',
    Sort_Order: Number(position.Sort_Order || getPositions(team).length + 1),
  };
  validateRequired_(record.Position_Name, '포지션 이름을 입력해 주세요.');

  const rowIndex = findRowIndex_(rows, headers.indexOf('Position_ID'), id);
  if (rowIndex > -1) {
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([headers.map((key) => record[key] || '')]);
  } else {
    sheet.appendRow(headers.map((key) => record[key] || ''));
  }
  refreshAssignmentHeaders_(team);
  syncSubmissionStatus(null, team);
  return getAdminPrimaryData(team);
}

function setPositionActive(positionId, active, team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureAdminPrimarySetup_(false, normalizedTeam);
  const sheet = getSheet_(SHEETS.POSITIONS, normalizedTeam);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = findRowIndex_(rows, headers.indexOf('Position_ID'), positionId);
  if (rowIndex > -1) sheet.getRange(rowIndex + 1, headers.indexOf('Active') + 1).setValue(active ? 'Yes' : 'No');
  refreshAssignmentHeaders_(normalizedTeam);
  syncSubmissionStatus(null, normalizedTeam);
  return getAdminPrimaryData(normalizedTeam);
}

function deletePosition(positionId, team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureAdminPrimarySetup_(false, normalizedTeam);
  const sheet = getSheet_(SHEETS.POSITIONS, normalizedTeam);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = findRowIndex_(rows, headers.indexOf('Position_ID'), positionId);
  if (rowIndex > -1) sheet.deleteRow(rowIndex + 1);
  refreshAssignmentHeaders_(normalizedTeam);
  syncSubmissionStatus(null, normalizedTeam);
  return getAdminPrimaryData(normalizedTeam);
}

function submitAvailability(payload) {
  const team = normalizeTeam_(payload.Team);
  ensurePublicSubmissionSetup_(false, team);
  const settings = getSettings(team);
  const month = normalizeMonth_(payload.Month || getActiveTargetMonth_(settings));
  const member = getMembers(team).find((item) => (
    item.Member_ID === payload.Member_ID || normalizeName_(item.Name) === normalizeName_(payload.Member_ID)
  ) && isActive_(item));
  if (!member) throw new Error('활성 멤버를 찾을 수 없습니다.');

  const selected = Array.isArray(payload.Available_Sundays) ? payload.Available_Sundays : [];
  const sheet = getSheet_(SHEETS.AVAILABILITY, team);
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
  syncSubmissionStatus(month, team);
  return { success: true, message: '제출이 완료되었습니다.', record };
}

function getSubmissionStatus(month, team) {
  syncSubmissionStatus(month, team);
  const normalizedTeam = normalizeTeam_(team);
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings(normalizedTeam)));
  return getRows_(SHEETS.SUBMISSION_STATUS, team).filter((row) => normalizeMonth_(row.Month) === targetMonth);
}

function syncSubmissionStatus(month, team) {
  const normalizedTeam = normalizeTeam_(team);
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings(normalizedTeam)));
  const members = getMembers(normalizedTeam).filter(isActive_);
  const submissions = getRows_(SHEETS.AVAILABILITY, normalizedTeam).filter((row) => normalizeMonth_(row.Month) === targetMonth);
  const existingOtherMonths = getRows_(SHEETS.SUBMISSION_STATUS, normalizedTeam).filter((row) => normalizeMonth_(row.Month) !== targetMonth);
  const statusRows = buildSubmissionStatusRows_(members, submissions, targetMonth);
  writeObjects_(SHEETS.SUBMISSION_STATUS, existingOtherMonths.concat(statusRows), normalizedTeam);
  return statusRows;
}

function getAvailabilityMatrixRows_(month, team) {
  const normalizedTeam = normalizeTeam_(team);
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings(normalizedTeam)));
  const members = getMembers(normalizedTeam).filter(isActive_);
  const submissions = getRows_(SHEETS.AVAILABILITY, normalizedTeam).filter((row) => normalizeMonth_(row.Month) === targetMonth);
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

function getAssignments(month, team) {
  const normalizedTeam = normalizeTeam_(team);
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings(normalizedTeam)));
  const activePositions = getPositions(normalizedTeam).filter(isActive_);
  const sundays = getVisibleSundays_(targetMonth);
  return getAssignmentsForPositions_(targetMonth, activePositions, sundays, normalizedTeam);
}

function getAssignmentsForPositions_(targetMonth, activePositions, sundays, team) {
  const visibleMonths = getMonthsForDates_(sundays);
  const rows = getRows_(SHEETS.ASSIGNMENTS, team).filter((row) => visibleMonths.indexOf(normalizeMonth_(row.Date)) > -1);
  return sundays.map((date) => {
    const existing = rows.find((row) => row.Date === date) || { Date: date };
    activePositions.forEach((position) => {
      if (existing[position.Position_Name] === undefined) existing[position.Position_Name] = '';
    });
    return existing;
  });
}

function saveAssignments(payload) {
  const team = normalizeTeam_(payload && payload.Team);
  ensureAdminPrimarySetup_(false, team);
  refreshAssignmentHeaders_(team);
  const incoming = payload.Assignments || [];
  const incomingMonths = getMonthsForDates_(incoming.map((row) => row.Date));
  const allRows = getRows_(SHEETS.ASSIGNMENTS, team);
  const keepRows = allRows.filter((row) => incomingMonths.indexOf(normalizeMonth_(row.Date)) === -1);
  writeObjects_(SHEETS.ASSIGNMENTS, keepRows.concat(incoming), team);
  return getAdminPrimaryData(team);
}

function generateLineup(month, team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureAdminPrimarySetup_(false, normalizedTeam);
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(getSettings(normalizedTeam)));
  const sundays = getVisibleSundays_(targetMonth);
  const positions = getPositions(normalizedTeam).filter(isActive_);
  const members = getMembers(normalizedTeam).filter(isActive_);
  const visibleMonths = getMonthsForDates_(sundays);
  const submissions = getRows_(SHEETS.AVAILABILITY, normalizedTeam).filter((row) => visibleMonths.indexOf(normalizeMonth_(row.Month)) > -1);
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
  saveAssignments({ Team: normalizedTeam, Month: targetMonth, Assignments: assignments });
  return getAdminPrimaryData(normalizedTeam);
}

function sendReminderEmails(month, team, memberIds) {
  const normalizedTeam = normalizeTeam_(team);
  ensureAdminPrimarySetup_(false, normalizedTeam);
  const settings = getSettings(normalizedTeam);
  const targetMonth = normalizeMonth_(month || getActiveTargetMonth_(settings));
  const webAppUrl = settings.Web_App_URL || ScriptApp.getService().getUrl();
  const adminEmail = clean_(settings.Admin_Email);
  const reminderMessage = clean_(settings.Reminder_Message) || DEFAULT_SETTINGS.Reminder_Message;
  const selectedIds = Array.isArray(memberIds) ? memberIds.map(String) : [];
  if (!selectedIds.length) throw new Error('리마인더를 보낼 멤버를 선택해 주세요.');

  const selectedMembers = getMembers(normalizedTeam)
    .filter(isActive_)
    .filter((member) => selectedIds.indexOf(String(member.Member_ID)) > -1);
  const logRows = [];

  selectedMembers.forEach((member) => {
    if (!member || !member.Email) return;
    try {
      MailApp.sendEmail({
        to: member.Email,
        subject: `[${TEAM_LABELS[normalizedTeam]}] ${targetMonth} 가능 주일 제출 부탁드립니다`,
        name: settings.Admin_Name || settings.Church_Name || TEAM_LABELS[normalizedTeam],
        replyTo: adminEmail || undefined,
        htmlBody: [
          `<p>${member.Name}님 안녕하세요.</p>`,
          `<p>${htmlEscape_(reminderMessage).replace(/\n/g, '<br>')}</p>`,
          `<p><a href="${webAppUrl}">제출하러 가기</a></p>`,
        ].join(''),
      });
      logRows.push([timestamp_(), 'Reminder', targetMonth, member.Name, member.Email, 'Sent']);
    } catch (error) {
      logRows.push([timestamp_(), 'Reminder', targetMonth, member.Name, member.Email, `Failed: ${error.message}`]);
    }
  });

  if (logRows.length) {
    const logSheet = getSheet_(SHEETS.EMAIL_LOG, normalizedTeam);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
  }
  return { sent: logRows.filter((row) => row[5] === 'Sent').length, total: selectedMembers.length, logs: logRows };
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

function refreshAssignmentHeaders_(team) {
  const normalizedTeam = normalizeTeam_(team);
  const sheet = getSheet_(SHEETS.ASSIGNMENTS, normalizedTeam);
  const activePositions = getPositions(normalizedTeam).filter(isActive_).map((position) => position.Position_Name);
  const headers = ['Date'].concat(activePositions);
  ensureHeaders_(sheet, headers, true);
}

function seedSettings_(team) {
  const normalizedTeam = normalizeTeam_(team);
  const sheet = getSheet_(SHEETS.SETTINGS, normalizedTeam);
  if (sheet.getLastRow() > 1) return;
  const defaults = Object.assign({}, DEFAULT_SETTINGS, { Team_Name: TEAM_LABELS[normalizedTeam] });
  const rows = Object.keys(defaults).map((key) => [key, defaults[key]]);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function seedPositions_(team) {
  const sheet = getSheet_(SHEETS.POSITIONS, team);
  if (sheet.getLastRow() > 1) return;
  ['자막', '조명', '음향', '카메라', '믹싱', 'PD', 'FD'].forEach((name, index) => {
    sheet.appendRow([createId_('P'), name, 1, 'Yes', index + 1]);
  });
}

function ensurePositionDefaults_(team) {
  const sheet = getSheet_(SHEETS.POSITIONS, team);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const headers = getHeaders_(sheet);
  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex += 1) {
    ensurePositionRowDefaults_(sheet, rowIndex, headers);
  }
}

function ensureMemberDefaults_(team) {
  const sheet = getSheet_(SHEETS.MEMBERS, team);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headers = getHeaders_(sheet);
  let changed = false;

  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex += 1) {
    if (ensureMemberRowDefaults_(sheet, rowIndex, headers)) changed = true;
  }

  if (changed) clearPublicCache_();
}

function ensureMemberRowDefaults_(sheet, rowIndex, headers) {
  const rowHeaders = headers || getHeaders_(sheet);
  const idCol = rowHeaders.indexOf('Member_ID') + 1;
  const nameCol = rowHeaders.indexOf('Name') + 1;
  const activeCol = rowHeaders.indexOf('Active') + 1;
  const joinDateCol = rowHeaders.indexOf('Join_Date') + 1;
  if (nameCol < 1) return false;

  const row = sheet.getRange(rowIndex, 1, 1, rowHeaders.length).getValues()[0];
  let name = row[nameCol - 1];
  let changed = false;

  if (!name && idCol > 0 && row[idCol - 1] && !isGeneratedId_(row[idCol - 1], 'M')) {
    name = row[idCol - 1];
    sheet.getRange(rowIndex, nameCol).setValue(name);
    sheet.getRange(rowIndex, idCol).clearContent();
    changed = true;
  }

  if (!name) return changed;
  if (idCol > 0 && !sheet.getRange(rowIndex, idCol).getValue()) {
    sheet.getRange(rowIndex, idCol).setValue(createId_('M'));
    changed = true;
  }
  if (activeCol > 0 && !sheet.getRange(rowIndex, activeCol).getValue()) {
    sheet.getRange(rowIndex, activeCol).setValue('Yes');
    changed = true;
  }
  if (joinDateCol > 0 && !sheet.getRange(rowIndex, joinDateCol).getValue()) {
    sheet.getRange(rowIndex, joinDateCol).setValue(today_());
    changed = true;
  }
  if (changed) clearPublicCache_();
  return changed;
}

function ensurePositionRowDefaults_(sheet, rowIndex, headers) {
  const rowHeaders = headers || getHeaders_(sheet);
  const idCol = rowHeaders.indexOf('Position_ID') + 1;
  const nameCol = rowHeaders.indexOf('Position_Name') + 1;
  const requiredCol = rowHeaders.indexOf('Required_Per_Sunday') + 1;
  const activeCol = rowHeaders.indexOf('Active') + 1;
  const sortCol = rowHeaders.indexOf('Sort_Order') + 1;
  if (nameCol < 1 || !sheet.getRange(rowIndex, nameCol).getValue()) return false;

  let changed = false;
  if (idCol > 0 && !sheet.getRange(rowIndex, idCol).getValue()) {
    sheet.getRange(rowIndex, idCol).setValue(createId_('P'));
    changed = true;
  }
  if (requiredCol > 0 && !sheet.getRange(rowIndex, requiredCol).getValue()) {
    sheet.getRange(rowIndex, requiredCol).setValue(1);
    changed = true;
  }
  if (activeCol > 0 && !sheet.getRange(rowIndex, activeCol).getValue()) {
    sheet.getRange(rowIndex, activeCol).setValue('Yes');
    changed = true;
  }
  if (sortCol > 0 && !sheet.getRange(rowIndex, sortCol).getValue()) {
    sheet.getRange(rowIndex, sortCol).setValue(getNextSortOrder_(sheet, rowHeaders));
    changed = true;
  }
  return changed;
}

function getNextSortOrder_(sheet, headers) {
  const sortCol = headers.indexOf('Sort_Order') + 1;
  if (sortCol < 1 || sheet.getLastRow() < 2) return 1;
  const values = sheet.getRange(2, sortCol, sheet.getLastRow() - 1, 1).getValues()
    .map((row) => Number(row[0] || 0))
    .filter((value) => value > 0);
  return values.length ? Math.max.apply(null, values) + 1 : 1;
}

function updateMemberField_(memberId, field, value, team) {
  const normalizedTeam = normalizeTeam_(team);
  ensureAdminPrimarySetup_(false, normalizedTeam);
  const sheet = getSheet_(SHEETS.MEMBERS, normalizedTeam);
  const headers = getHeaders_(sheet);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = findRowIndex_(rows, headers.indexOf('Member_ID'), memberId);
  if (rowIndex > -1) sheet.getRange(rowIndex + 1, headers.indexOf(field) + 1).setValue(value);
  syncSubmissionStatus(null, normalizedTeam);
  clearPublicCache_();
  return getAdminPrimaryData(normalizedTeam);
}

function getSheet_(name, team) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(teamSheet_(name, team));
  if (!sheet) throw new Error(`${name} sheet is missing. Run setupWorkbook first.`);
  return sheet;
}

function teamSheet_(sheetName, team) {
  return `${TEAM_SHEET_PREFIXES[normalizeTeam_(team)] || ''}${sheetName}`;
}

function baseSheetName_(sheetName) {
  const value = String(sheetName || '');
  const prefixes = Object.keys(TEAM_SHEET_PREFIXES).map((team) => TEAM_SHEET_PREFIXES[team]);
  const prefix = prefixes.find((item) => item && value.indexOf(item) === 0);
  return prefix ? value.slice(prefix.length) : sheetName;
}

function normalizeTeam_(team) {
  return String(team || TEAMS.PRODUCTION).toLowerCase() === TEAMS.PRAISE ? TEAMS.PRAISE : TEAMS.PRODUCTION;
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

function getRows_(sheetName, team) {
  const sheet = getSheet_(sheetName, team);
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

function writeObjects_(sheetName, objects, team) {
  const sheet = getSheet_(sheetName, team);
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

function htmlEscape_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
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
  CacheService.getScriptCache().removeAll([
    'publicData:v1',
    `publicData:v2:${TEAMS.PRAISE}`,
    `publicData:v2:${TEAMS.PRODUCTION}`,
  ]);
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
