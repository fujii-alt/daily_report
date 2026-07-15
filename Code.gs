const USERS_SHEET = 'Users';
const REQUEST_MASTER_SHEET = 'RequestMaster';
const REPORTS_SHEET = 'DailyReports';
const SUMMARIES_SHEET = 'DailySummaries';
const SETTINGS_SHEET = 'AppSettings';

const USER_COL_NAME = 1;
const USER_COL_PIN = 2;
const USER_COL_ENABLED = 3;
const USER_COL_EMAIL = 4;

const REQUEST_COL_NO = 1;
const REQUEST_COL_ENABLED = 2;

const REPORT_COL_SUBMITTED_AT = 1;
const REPORT_COL_WORK_DATE = 2;
const REPORT_COL_NAME = 3;
const REPORT_COL_REQUEST_NO = 4;
const REPORT_COL_WORK_CONTENT = 5;
const REPORT_COL_WORK_HOURS = 6;
const REPORT_COL_TOKEN = 7;
const REPORT_COL_REPORT_ID = 8;
const REPORT_COL_DIRECT_DIARY = 9;

const SUMMARY_COL_SUBMITTED_AT = 1;
const SUMMARY_COL_WORK_DATE = 2;
const SUMMARY_COL_NAME = 3;
const SUMMARY_COL_INPUT_TOTAL = 4;
const SUMMARY_COL_CUMULATIVE_TOTAL = 5;
const SUMMARY_COL_DIARY = 6;
const SUMMARY_COL_TOKEN = 7;
const SUMMARY_COL_REPORT_ID = 8;

const SESSION_PREFIX = 'session_';
const SESSION_EXPIRE_SEC = 60 * 60 * 6; // 6時間
const SETTING_MAIL_TO = 'MAIL_TO';
const DIRECT_INPUT_TOKEN = 'SPREADSHEET_DIRECT';

/**
 * スプレッドシートを開いたときに、直接入力用の管理メニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日報管理')
    .addItem('直接入力用の設定を適用', 'setupDirectInputSheet')
    .addItem('直接入力分の集計を更新', 'rebuildDirectInputSummaries')
    .addToUi();
}

/**
 * DailyReports シートへ直接入力された行を自動補完します。
 * スクリプトからの書き込みでは発火しないため、アプリ送信時の処理には影響しません。
 * この処理ではメール送信関数を呼びません。
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== REPORTS_SHEET) return;
  if (e.range.getLastRow() < 2) return;

  const inputStartCol = REPORT_COL_WORK_DATE;
  const inputEndCol = REPORT_COL_DIRECT_DIARY;
  if (e.range.getLastColumn() < inputStartCol || e.range.getColumn() > inputEndCol) {
    return;
  }

  const firstRow = Math.max(2, e.range.getRow());
  const lastRow = e.range.getLastRow();

  for (let row = firstRow; row <= lastRow; row++) {
    completeDirectInputRow_(sheet, row);
  }

  rebuildDirectInputSummaries_();
}

/**
 * 直接入力しやすいように、見出し・入力規則・説明を設定します。
 * 初回導入時とUsers／RequestMaster更新後にメニューから実行してください。
 */
function setupDirectInputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORTS_SHEET);
  if (!sheet) {
    throw new Error('DailyReportsシートが見つかりません。');
  }

  // Googleスプレッドシートの「テーブル」では、列型・ドロップダウンなどを
  // Range.setDataValidation() / clearDataValidations() / setNote() 等から
  // 上書きできない場合があります。
  // そのため、この設定処理ではテーブル内のセルを一切変更しません。
  // 既存のテーブル列型をそのまま使用し、直接入力の自動補完だけを onEdit で行います。
  const headerValues = sheet
    .getRange(1, 1, 1, Math.min(REPORT_COL_DIRECT_DIARY, sheet.getMaxColumns()))
    .getDisplayValues()[0];

  const expectedHeaders = [
    '登録日時',
    '作業日',
    '氏名',
    '依頼No',
    '作業内容',
    '作業時間',
    'セッションID',
    '日報ID',
    '一日の総括',
  ];

  const warnings = [];
  expectedHeaders.forEach((expected, index) => {
    const actual = String(headerValues[index] || '').trim();
    if (!actual) {
      warnings.push(`${columnLetter_(index + 1)}列の見出しが空欄です`);
    }
  });

  const warningText = warnings.length
    ? `\n\n確認事項：\n・${warnings.join('\n・')}`
    : '';

  SpreadsheetApp.getUi().alert(
    'テーブル対応の直接入力設定を確認しました。\n\n' +
    'この処理では、テーブルの列型・入力規則・見出し・メモを変更しません。\n' +
    'B～F列を入力すると、A列・G列・H列を自動補完します。\n' +
    'I列の「一日の総括」は任意です。\n' +
    'スプレッドシートから直接入力した場合、メールは送信されません。' +
    warningText
  );
}

function columnLetter_(columnNumber) {
  let number = Number(columnNumber || 0);
  let result = '';

  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }

  return result || '?';
}

/**
 * メニューから直接入力分の集計を再構築するための公開関数です。
 */
function rebuildDirectInputSummaries() {
  rebuildDirectInputSummaries_();
  SpreadsheetApp.getUi().alert('直接入力分の集計を更新しました。');
}

function getMasterData() {
  try {
    return {
      ok: true,
      users: getActiveUsers_(),
      requestNos: getActiveRequestNos_(),
    };
  } catch (e) {
    console.error(e);
    return {
      ok: false,
      message: e.message || 'マスタデータ取得時にエラーが発生しました。',
    };
  }
}

function verifyLogin(name) {
  name = String(name || '').trim();

  if (!name) {
    return { ok: false, message: '氏名を選択してください。' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) {
    return { ok: false, message: 'Usersシートが見つかりません。' };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return { ok: false, message: 'Usersシートに利用者が登録されていません。' };
  }

  for (let i = 1; i < values.length; i++) {
    const rowName = String(values[i][USER_COL_NAME - 1] || '').trim();
    const enabled = values[i][USER_COL_ENABLED - 1];

    if (rowName === name && isEnabledValue_(enabled)) {
      const token = Utilities.getUuid();
      const payload = {
        name: rowName,
        loginAt: new Date().toISOString(),
      };

      CacheService.getScriptCache().put(
        SESSION_PREFIX + token,
        JSON.stringify(payload),
        SESSION_EXPIRE_SEC
      );

      return {
        ok: true,
        token,
        name: rowName,
      };
    }
  }

  return { ok: false, message: '氏名が登録されていないか、有効ではありません。' };
}

function getDayTotal(token, workDate) {
  const session = getSession_(token);
  if (!session.ok) return session;

  const normalizedDate = normalizeWorkDate_(workDate);
  if (!normalizedDate) {
    return { ok: false, message: '日付が不正です。' };
  }

  return {
    ok: true,
    name: session.name,
    total: calcDayTotal_(session.name, normalizedDate),
  };
}

function saveReport(data) {
  const token = String((data && data.token) || '').trim();
  const workDate = normalizeWorkDate_(data && data.workDate);
  const rows = Array.isArray(data && data.rows) ? data.rows : [];
  const diary = String((data && data.diary) || '').trim();

  const session = getSession_(token);
  if (!session.ok) return session;

  if (!workDate) {
    return { ok: false, message: '日付を入力してください。' };
  }

  if (rows.length === 0) {
    return { ok: false, message: '入力行がありません。' };
  }

  const validRows = rows
    .map((row) => ({
      requestNo: String((row && row.requestNo) || '').trim(),
      workContent: String((row && row.workContent) || '').trim(),
      workHours: Number((row && row.workHours) || 0),
    }))
    .filter((row) => row.requestNo !== '' || row.workContent !== '' || row.workHours > 0);

  if (validRows.length === 0) {
    return { ok: false, message: '少なくとも1行は入力してください。' };
  }

  for (const row of validRows) {
    if (!row.requestNo) {
      return { ok: false, message: '依頼No.を選択してください。' };
    }
    if (!row.workContent) {
      return { ok: false, message: '作業内容を入力してください。' };
    }
    if (!(row.workHours > 0)) {
      return { ok: false, message: '作業時間は0より大きい数値を入力してください。' };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportsSheet = ss.getSheetByName(REPORTS_SHEET);
  const summariesSheet = ss.getSheetByName(SUMMARIES_SHEET);

  if (!reportsSheet) {
    return { ok: false, message: 'DailyReportsシートが見つかりません。' };
  }
  if (!summariesSheet) {
    return { ok: false, message: 'DailySummariesシートが見つかりません。' };
  }

  const now = new Date();
  const reportId = buildDailyReportId_(workDate, session.name);

  const appendValues = validRows.map((row) => [
    now,
    workDate,
    session.name,
    row.requestNo,
    row.workContent,
    row.workHours,
    token,
    reportId,
  ]);

  reportsSheet
    .getRange(
      reportsSheet.getLastRow() + 1,
      1,
      appendValues.length,
      appendValues[0].length
    )
    .setValues(appendValues);

  const inputTotal = Number(
    validRows.reduce((sum, row) => sum + row.workHours, 0).toFixed(2)
  );
  const cumulativeTotal = calcDayTotal_(session.name, workDate);

  summariesSheet.appendRow([
    now,
    workDate,
    session.name,
    inputTotal,
    cumulativeTotal,
    diary,
    token,
    reportId,
  ]);

  SpreadsheetApp.flush();

  const report = buildCombinedDayReport_(session.name, workDate, {
    fallbackSubmittedAt: now,
    fallbackReportId: reportId,
  });

  report.userEmail = getUserEmailByName_(session.name);

  const mailResult = sendReportMails_(report);

  return {
    ok: true,
    message: `送信しました。${mailResult.message}`.trim(),
    total: report.cumulativeTotal,
  };
}
function getRequestMasterList(token) {
  const session = getSession_(token);
  if (!session.ok) return session;

  try {
    return {
      ok: true,
      items: getRequestMasterItems_(),
    };
  } catch (e) {
    console.error(e);
    return {
      ok: false,
      message: e.message || '依頼No.一覧取得時にエラーが発生しました。',
    };
  }
}

function addRequestNo(token, requestNo) {
  const session = getSession_(token);
  if (!session.ok) return session;

  requestNo = String(requestNo || '').trim();
  if (!requestNo) {
    return { ok: false, message: '依頼No.を入力してください。' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(REQUEST_MASTER_SHEET);
    if (!sheet) {
      return { ok: false, message: 'RequestMasterシートが見つかりません。' };
    }

    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const rowRequestNo = String(values[i][REQUEST_COL_NO - 1] || '').trim();
      if (rowRequestNo === requestNo) {
        return { ok: false, message: 'その依頼No.は既に登録されています。' };
      }
    }

    sheet.appendRow([requestNo, true]);

    return {
      ok: true,
      message: '依頼No.を追加しました。',
      items: getRequestMasterItems_(),
    };
  } catch (e) {
    console.error(e);
    return {
      ok: false,
      message: e.message || '依頼No.追加時にエラーが発生しました。',
    };
  } finally {
    lock.releaseLock();
  }
}

function disableRequestNo(token, requestNo) {
  const session = getSession_(token);
  if (!session.ok) return session;

  requestNo = String(requestNo || '').trim();
  if (!requestNo) {
    return { ok: false, message: '依頼No.が不正です。' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(REQUEST_MASTER_SHEET);
    if (!sheet) {
      return { ok: false, message: 'RequestMasterシートが見つかりません。' };
    }

    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const rowRequestNo = String(values[i][REQUEST_COL_NO - 1] || '').trim();
      if (rowRequestNo === requestNo) {
        sheet.getRange(i + 1, REQUEST_COL_ENABLED).setValue(false);
        return {
          ok: true,
          message: '依頼No.を無効化しました。',
          items: getRequestMasterItems_(),
        };
      }
    }

    return { ok: false, message: '対象の依頼No.が見つかりません。' };
  } catch (e) {
    console.error(e);
    return {
      ok: false,
      message: e.message || '依頼No.無効化時にエラーが発生しました。',
    };
  } finally {
    lock.releaseLock();
  }
}

function getHistoryDay(token, workDate) {
  const session = getSession_(token);
  if (!session.ok) return session;

  const normalizedDate = normalizeWorkDate_(workDate);
  if (!normalizedDate) {
    return { ok: false, message: '日付が不正です。' };
  }

  try {
    return {
      ok: true,
      ...buildDayHistory_(session.name, normalizedDate),
    };
  } catch (e) {
    console.error(e);
    return {
      ok: false,
      message: e.message || '1日分履歴取得時にエラーが発生しました。',
    };
  }
}

function getHistoryWeek(token, baseDate) {
  const session = getSession_(token);
  if (!session.ok) return session;

  const normalizedDate = normalizeWorkDate_(baseDate);
  if (!normalizedDate) {
    return { ok: false, message: '基準日が不正です。' };
  }

  try {
    return {
      ok: true,
      ...buildWeekHistory_(session.name, normalizedDate),
    };
  } catch (e) {
    console.error(e);
    return {
      ok: false,
      message: e.message || '1週間分履歴取得時にエラーが発生しました。',
    };
  }
}

function getActiveUsers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) throw new Error('Usersシートが見つかりません。');

  const values = sheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][USER_COL_NAME - 1] || '').trim();
    const enabled = values[i][USER_COL_ENABLED - 1];
    if (name && isEnabledValue_(enabled)) {
      users.push(name);
    }
  }

  return users;
}

function getActiveRequestNos_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REQUEST_MASTER_SHEET);
  if (!sheet) throw new Error('RequestMasterシートが見つかりません。');

  const values = sheet.getDataRange().getValues();
  const requestNos = [];

  for (let i = 1; i < values.length; i++) {
    const requestNo = String(values[i][REQUEST_COL_NO - 1] || '').trim();
    const enabled = values[i][REQUEST_COL_ENABLED - 1];
    if (requestNo && isEnabledValue_(enabled)) {
      requestNos.push(requestNo);
    }
  }

  return requestNos;
}

function getRequestMasterItems_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REQUEST_MASTER_SHEET);
  if (!sheet) throw new Error('RequestMasterシートが見つかりません。');

  const values = sheet.getDataRange().getValues();
  const items = [];

  for (let i = 1; i < values.length; i++) {
    const requestNo = String(values[i][REQUEST_COL_NO - 1] || '').trim();
    if (!requestNo) continue;

    items.push({
      requestNo,
      enabled: isEnabledValue_(values[i][REQUEST_COL_ENABLED - 1]),
    });
  }

  return items;
}

function isEnabledValue_(value) {
  const str = String(value).trim().toUpperCase();
  return str !== '' && str !== 'FALSE';
}

function getSession_(token) {
  if (!token) {
    return { ok: false, message: 'セッションが無効です。再ログインしてください。' };
  }

  const raw = CacheService.getScriptCache().get(SESSION_PREFIX + token);
  if (!raw) {
    return { ok: false, message: 'ログイン有効期限が切れました。再ログインしてください。' };
  }

  const session = JSON.parse(raw);
  return {
    ok: true,
    name: session.name,
  };
}

function calcDayTotal_(name, workDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const values = sheet.getDataRange().getValues();
  let total = 0;

  for (let i = 1; i < values.length; i++) {
    const rowDate = normalizeSheetWorkDate_(values[i][REPORT_COL_WORK_DATE - 1]);
    const rowName = String(values[i][REPORT_COL_NAME - 1] || '').trim();
    const requestNo = String(values[i][REPORT_COL_REQUEST_NO - 1] || '').trim();
    const workContent = String(values[i][REPORT_COL_WORK_CONTENT - 1] || '').trim();
    const hours = Number(values[i][REPORT_COL_WORK_HOURS - 1] || 0);

    if (
      rowDate === workDate &&
      rowName === name &&
      requestNo &&
      workContent &&
      hours > 0
    ) {
      total += hours;
    }
  }

  return Number(total.toFixed(2));
}

function getSettingValue_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    throw new Error('AppSettingsシートが見つかりません。');
  }

  const values = sheet.getDataRange().getValues();
  const matchedValues = [];

  for (let i = 1; i < values.length; i++) {
    const rowKey = String(values[i][0] || '').trim();
    const rowValue = String(values[i][1] || '').trim();
    if (rowKey === key && rowValue) {
      matchedValues.push(rowValue);
    }
  }

  return matchedValues.join('\n');
}

function getUserEmailByName_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return '';

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const rowName = String(values[i][USER_COL_NAME - 1] || '').trim();
    if (rowName === name) {
      return String(values[i][USER_COL_EMAIL - 1] || '').trim();
    }
  }
  return '';
}

function buildDailyReportId_(workDate, name) {
  return `${String(workDate || '').replace(/-/g, '')}-${String(name || '').trim()}`;
}

function buildCombinedDayReport_(name, workDate, options) {
  const opts = options || {};
  const reportRows = getReportRowsForDay_(name, workDate);
  const summaryRows = getSummaryRowsForDay_(name, workDate);
  const reportId = String(opts.fallbackReportId || buildDailyReportId_(workDate, name)).trim();

  const totalHours = Number(
    reportRows.reduce((sum, row) => sum + Number(row.workHours || 0), 0).toFixed(2)
  );

  const diaries = summaryRows
    .map((row) => String(row.diary || '').trim())
    .filter((text) => text !== '');

  const latestSummary = summaryRows.length ? summaryRows[summaryRows.length - 1] : null;
  const latestReport = reportRows.length ? reportRows[reportRows.length - 1] : null;
  const fallbackSubmittedAt = opts.fallbackSubmittedAt || new Date();

  const submittedAt =
    (latestSummary && latestSummary.submittedAt) ||
    (latestReport && latestReport.submittedAt) ||
    toIsoString_(fallbackSubmittedAt);

  const latestInputTotal = latestSummary ? Number(latestSummary.inputTotal || 0) : totalHours;
  const cumulativeTotal = totalHours;

  return {
    submittedAt,
    workDate,
    workDateDisplay: formatJapaneseDate_(workDate),
    name,
    rows: reportRows,
    diary: diaries.join('\n\n'),
    inputTotal: Number(latestInputTotal.toFixed ? latestInputTotal.toFixed(2) : latestInputTotal),
    dailyTotal: cumulativeTotal,
    cumulativeTotal,
    reportCount: summaryRows.length || (reportRows.length ? 1 : 0),
    reportId,
    userEmail: '',
  };
}

function getReportRowsForDay_(name, workDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const rowDate = normalizeSheetWorkDate_(values[i][REPORT_COL_WORK_DATE - 1]);
    const rowName = String(values[i][REPORT_COL_NAME - 1] || '').trim();

    if (rowDate !== workDate || rowName !== name) continue;

    const requestNo = String(values[i][REPORT_COL_REQUEST_NO - 1] || '').trim();
    const workContent = String(values[i][REPORT_COL_WORK_CONTENT - 1] || '').trim();
    const hours = Number(values[i][REPORT_COL_WORK_HOURS - 1] || 0);
    if (!requestNo || !workContent || !(hours > 0)) continue;

    rows.push({
      submittedAt: toIsoString_(values[i][REPORT_COL_SUBMITTED_AT - 1]),
      requestNo,
      workContent,
      workHours: Number(hours.toFixed ? hours.toFixed(2) : hours),
      reportId: String(values[i][REPORT_COL_REPORT_ID - 1] || '').trim(),
      sheetRow: i + 1,
    });
  }

  return rows.sort((a, b) => {
    const timeCompare = String(a.submittedAt).localeCompare(String(b.submittedAt));
    return timeCompare || a.sheetRow - b.sheetRow;
  });
}

function getSummaryRowsForDay_(name, workDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SUMMARIES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const rowDate = normalizeSheetWorkDate_(values[i][SUMMARY_COL_WORK_DATE - 1]);
    const rowName = String(values[i][SUMMARY_COL_NAME - 1] || '').trim();

    if (rowDate !== workDate || rowName !== name) continue;

    rows.push({
      submittedAt: toIsoString_(values[i][SUMMARY_COL_SUBMITTED_AT - 1]),
      inputTotal: Number(values[i][SUMMARY_COL_INPUT_TOTAL - 1] || 0),
      cumulativeTotal: Number(values[i][SUMMARY_COL_CUMULATIVE_TOTAL - 1] || 0),
      diary: String(values[i][SUMMARY_COL_DIARY - 1] || '').trim(),
      reportId: String(values[i][SUMMARY_COL_REPORT_ID - 1] || '').trim(),
    });
  }

  return rows.sort((a, b) => String(a.submittedAt).localeCompare(String(b.submittedAt)));
}

function formatJapaneseDate_(value) {
  const normalized = normalizeWorkDate_(value);
  if (!normalized) return String(value || '');

  const parts = normalized.split('-');
  return `${Number(parts[0])}年${Number(parts[1])}月${Number(parts[2])}日`;
}


function buildDayHistory_(name, workDate) {
  const combinedReport = buildCombinedDayReport_(name, workDate, {
    fallbackReportId: buildDailyReportId_(workDate, name),
  });

  const entries = combinedReport.rows.length
    ? [
        {
          reportId: combinedReport.reportId,
          submittedAt: combinedReport.submittedAt,
          inputTotal: combinedReport.dailyTotal,
          cumulativeTotal: combinedReport.cumulativeTotal,
          diary: combinedReport.diary,
          rows: combinedReport.rows,
        },
      ]
    : [];

  return {
    name,
    workDate: combinedReport.workDateDisplay,
    dailyTotal: combinedReport.dailyTotal,
    entryCount: combinedReport.reportCount,
    entries,
  };
}
function buildWeekHistory_(name, baseDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summariesSheet = ss.getSheetByName(SUMMARIES_SHEET);

  const startDate = getOffsetDateString_(baseDate, -6);
  const endDate = baseDate;
  const daysMap = {};

  for (let i = 0; i < 7; i++) {
    const date = getOffsetDateString_(startDate, i);
    daysMap[date] = {
      workDate: date,
      totalHours: 0,
      reportCount: 0,
      diaries: [],
      latestSubmittedAt: '',
    };
  }

  if (summariesSheet && summariesSheet.getLastRow() >= 2) {
    const summaryValues = summariesSheet.getDataRange().getValues();

    for (let i = 1; i < summaryValues.length; i++) {
      const rowDate = normalizeSheetWorkDate_(summaryValues[i][SUMMARY_COL_WORK_DATE - 1]);
      const rowName = String(summaryValues[i][SUMMARY_COL_NAME - 1] || '').trim();

      if (rowName !== name) continue;
      if (!isDateInRange_(rowDate, startDate, endDate)) continue;

      const day = daysMap[rowDate];
      if (!day) continue;

      const submittedAt = toIsoString_(summaryValues[i][SUMMARY_COL_SUBMITTED_AT - 1]);
      const inputTotal = Number(summaryValues[i][SUMMARY_COL_INPUT_TOTAL - 1] || 0);
      const diary = String(summaryValues[i][SUMMARY_COL_DIARY - 1] || '').trim();

      day.totalHours = Number((day.totalHours + inputTotal).toFixed(2));
      day.reportCount += 1;

      if (diary) {
        day.diaries.push({
          submittedAt,
          text: diary,
        });
      }

      if (!day.latestSubmittedAt || submittedAt > day.latestSubmittedAt) {
        day.latestSubmittedAt = submittedAt;
      }
    }
  }

  const days = Object.values(daysMap)
    .sort((a, b) => b.workDate.localeCompare(a.workDate))
    .map((day) => ({
      workDate: formatJapaneseDate_(day.workDate),
      totalHours: day.totalHours,
      reportCount: day.reportCount,
      latestDiary: day.diaries
        .sort((a, b) => String(a.submittedAt).localeCompare(String(b.submittedAt)))
        .map((item) => item.text)
        .join('\n\n'),
      latestSubmittedAt: day.latestSubmittedAt,
    }));

  return {
    name,
    baseDate: formatJapaneseDate_(baseDate),
    startDate: formatJapaneseDate_(startDate),
    endDate: formatJapaneseDate_(endDate),
    days,
  };
}
/**
 * DailyReports の直接入力行に、内部管理用の値を補完します。
 */
function completeDirectInputRow_(sheet, rowNumber) {
  const rowValues = sheet
    .getRange(rowNumber, 1, 1, REPORT_COL_DIRECT_DIARY)
    .getValues()[0];

  const workDate = normalizeWorkDate_(rowValues[REPORT_COL_WORK_DATE - 1]);
  const name = String(rowValues[REPORT_COL_NAME - 1] || '').trim();
  const requestNo = String(rowValues[REPORT_COL_REQUEST_NO - 1] || '').trim();
  const workContent = String(rowValues[REPORT_COL_WORK_CONTENT - 1] || '').trim();
  const workHours = Number(rowValues[REPORT_COL_WORK_HOURS - 1] || 0);
  const token = String(rowValues[REPORT_COL_TOKEN - 1] || '').trim();

  const hasInput = Boolean(
    workDate || name || requestNo || workContent || workHours > 0 ||
    String(rowValues[REPORT_COL_DIRECT_DIARY - 1] || '').trim()
  );

  if (!hasInput) {
    if (token === DIRECT_INPUT_TOKEN) {
      sheet.getRange(rowNumber, REPORT_COL_SUBMITTED_AT).clearContent();
      sheet.getRange(rowNumber, REPORT_COL_TOKEN).clearContent();
      sheet.getRange(rowNumber, REPORT_COL_REPORT_ID).clearContent();
    }
    return;
  }

  const isComplete = Boolean(
    workDate && name && requestNo && workContent && workHours > 0
  );
  if (!isComplete) return;

  if (!rowValues[REPORT_COL_SUBMITTED_AT - 1]) {
    sheet.getRange(rowNumber, REPORT_COL_SUBMITTED_AT).setValue(new Date());
  }

  if (!token) {
    sheet.getRange(rowNumber, REPORT_COL_TOKEN).setValue(DIRECT_INPUT_TOKEN);
  }

  const currentReportId = String(rowValues[REPORT_COL_REPORT_ID - 1] || '').trim();
  if (!currentReportId || token === DIRECT_INPUT_TOKEN) {
    sheet
      .getRange(rowNumber, REPORT_COL_REPORT_ID)
      .setValue(buildDailyReportId_(workDate, name));
  }
}

/**
 * 直接入力行を日付・氏名ごとにまとめ、DailySummariesへ反映します。
 * アプリから登録されたDailySummaries行は保持します。
 */
function rebuildDirectInputSummaries_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reportsSheet = ss.getSheetByName(REPORTS_SHEET);
    const summariesSheet = ss.getSheetByName(SUMMARIES_SHEET);

    if (!reportsSheet) {
      throw new Error('DailyReportsシートが見つかりません。');
    }
    if (!summariesSheet) {
      throw new Error('DailySummariesシートが見つかりません。');
    }

    const reportValues = reportsSheet.getDataRange().getValues();
    const groups = {};
    const allTotals = {};

    for (let i = 1; i < reportValues.length; i++) {
      const row = reportValues[i];
      const workDate = normalizeSheetWorkDate_(row[REPORT_COL_WORK_DATE - 1]);
      const name = String(row[REPORT_COL_NAME - 1] || '').trim();
      const requestNo = String(row[REPORT_COL_REQUEST_NO - 1] || '').trim();
      const workContent = String(row[REPORT_COL_WORK_CONTENT - 1] || '').trim();
      const hours = Number(row[REPORT_COL_WORK_HOURS - 1] || 0);

      if (!workDate || !name || !requestNo || !workContent || !(hours > 0)) continue;

      const key = `${workDate}|||${name}`;
      allTotals[key] = Number(((allTotals[key] || 0) + hours).toFixed(2));

      const token = String(row[REPORT_COL_TOKEN - 1] || '').trim();
      if (token !== DIRECT_INPUT_TOKEN) continue;

      if (!groups[key]) {
        groups[key] = {
          workDate,
          name,
          inputTotal: 0,
          diaries: [],
          latestSubmittedAt: null,
          reportId: buildDailyReportId_(workDate, name),
        };
      }

      const group = groups[key];
      group.inputTotal = Number((group.inputTotal + hours).toFixed(2));

      const diary = String(row[REPORT_COL_DIRECT_DIARY - 1] || '').trim();
      if (diary && !group.diaries.includes(diary)) {
        group.diaries.push(diary);
      }

      const submittedAt = row[REPORT_COL_SUBMITTED_AT - 1];
      const submittedDate = submittedAt instanceof Date ? submittedAt : new Date(submittedAt);
      if (!isNaN(submittedDate.getTime())) {
        if (!group.latestSubmittedAt || submittedDate > group.latestSubmittedAt) {
          group.latestSubmittedAt = submittedDate;
        }
      }
    }

    const currentValues = summariesSheet.getDataRange().getValues();
    const defaultHeaders = [
      '登録日時',
      '作業日',
      '氏名',
      '入力合計',
      '累計',
      '一日の総括',
      '入力元／トークン',
      '日報ID',
    ];
    const currentHeader = currentValues.length ? currentValues[0] : [];
    const header = defaultHeaders.map((value, index) => currentHeader[index] || value);

    const preservedRows = currentValues
      .slice(1)
      .filter((row) => String(row[SUMMARY_COL_TOKEN - 1] || '').trim() !== DIRECT_INPUT_TOKEN);

    const directRows = Object.keys(groups)
      .sort()
      .map((key) => {
        const group = groups[key];
        return [
          group.latestSubmittedAt || new Date(),
          group.workDate,
          group.name,
          group.inputTotal,
          Number(allTotals[key] || group.inputTotal),
          group.diaries.join('\n\n'),
          DIRECT_INPUT_TOKEN,
          group.reportId,
        ];
      });

    const outputRows = preservedRows.concat(directRows);
    const requiredRows = Math.max(outputRows.length + 1, 1);
    if (summariesSheet.getMaxRows() < requiredRows) {
      summariesSheet.insertRowsAfter(
        summariesSheet.getMaxRows(),
        requiredRows - summariesSheet.getMaxRows()
      );
    }

    const clearRows = Math.max(summariesSheet.getLastRow(), requiredRows);
    summariesSheet.getRange(1, 1, clearRows, 8).clearContent();
    summariesSheet.getRange(1, 1, 1, 8).setValues([header.slice(0, 8)]);

    if (outputRows.length) {
      summariesSheet.getRange(2, 1, outputRows.length, 8).setValues(outputRows);
    }
  } finally {
    lock.releaseLock();
  }
}

function normalizeWorkDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const str = String(value || '').trim();
  if (!str) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str;
  }

  const slashMatch = str.match(/^(\d{4})[\/年](\d{1,2})[\/月](\d{1,2})日?$/);
  if (slashMatch) {
    return `${slashMatch[1]}-${String(slashMatch[2]).padStart(2, '0')}-${String(slashMatch[3]).padStart(2, '0')}`;
  }

  const date = new Date(str);
  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return '';
}

function normalizeSheetWorkDate_(value) {
  return normalizeWorkDate_(value);
}

function isDateInRange_(target, start, end) {
  return target >= start && target <= end;
}

function getOffsetDateString_(baseDate, offsetDays) {
  const date = new Date(`${baseDate}T00:00:00`);
  date.setDate(date.getDate() + offsetDays);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function toIsoString_(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }
  const date = new Date(value);
  if (!isNaN(date.getTime())) {
    return date.toISOString();
  }
  return String(value || '');
}
