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

const ANALYTICS_SHEET = '集計ダッシュボード';
const ANALYTICS_MODE_CELL = 'B2';
const ANALYTICS_BASE_DATE_CELL = 'B3';
const ANALYTICS_CUSTOM_START_CELL = 'B4';
const ANALYTICS_CUSTOM_END_CELL = 'B5';
const ANALYTICS_DATE_BASIS_CELL = 'B6';
const ANALYTICS_PERIOD_LABEL_CELL = 'B7';
const ANALYTICS_USER_FILTER_CELL = 'E2';


const PAYROLL_SUMMARY_SHEET = '給与集計';
const HOURLY_RATE_MASTER_SHEET = '時給マスタ';
const PAYROLL_ADJUSTMENT_SHEET = '支給時間調整';
const PAYROLL_MISC_ADJUSTMENT_SHEET = '給与加減算';
const PAYROLL_FINALIZATION_SHEET = '給与確定履歴';
const PAYROLL_MONTH_CELL = 'B2';
const PAYROLL_USER_FILTER_CELL = 'E2';
const PAYROLL_CONFIRM_LABEL_CELL = 'G2';
const PAYROLL_CONFIRM_CELL = 'H2';
const DEFAULT_HOURLY_RATE = 1180;

const RATE_COL_NAME = 1;
const RATE_COL_HOURLY_RATE = 2;
const RATE_COL_START_DATE = 3;
const RATE_COL_END_DATE = 4;
const RATE_COL_ENABLED = 5;
const RATE_COL_NOTE = 6;

const ADJUST_COL_PAYROLL_MONTH = 1;
const ADJUST_COL_NAME = 2;
const ADJUST_COL_REQUEST_NO = 3;
const ADJUST_COL_PAID_HOURS = 4;
const ADJUST_COL_NOTE = 5;
const ADJUST_COL_INCENTIVE_HOURS = 6;
const ADJUST_COL_ADJUSTED_PAY = 7;

const MISC_COL_PAYROLL_MONTH = 1;
const MISC_COL_NAME = 2;
const MISC_COL_AMOUNT = 3;
const MISC_COL_CATEGORY = 4;
const MISC_COL_NOTE = 5;
const MISC_COL_REGISTERED_AT = 6;

const FINAL_COL_PAYROLL_MONTH = 1;
const FINAL_COL_NAME = 2;
const FINAL_COL_CALCULATED_PAY = 3;
const FINAL_COL_FINAL_PAY = 4;
const FINAL_COL_DIFFERENCE = 5;
const FINAL_COL_METHOD = 6;
const FINAL_COL_NOTE = 7;
const FINAL_COL_CONFIRMED_AT = 8;
const FINAL_COL_CONFIRMED_BY = 9;

/**
 * スプレッドシートを開いたときに、直接入力用の管理メニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日報管理')
    .addItem('直接入力用の設定を適用', 'setupDirectInputSheet')
    .addItem('直接入力分の集計を更新', 'rebuildDirectInputSummaries')
    .addSeparator()
    .addItem('集計ダッシュボードを作成・開く', 'setupAnalyticsDashboard')
    .addItem('集計ダッシュボードを更新', 'refreshAnalyticsDashboard')
    .addSeparator()
    .addItem('給与関連シートを作成・開く', 'setupPayrollSheets')
    .addItem('給与確定履歴を作成・開く', 'openPayrollFinalizationSheet')
    .addItem('給与集計を履歴へ確定保存', 'confirmPayrollSummaryToHistory')
    .addItem('給与集計を更新', 'refreshPayrollSummary')
    .addToUi();
}

/**
 * シンプルトリガー側では、DailyReports の A・G・H列補完だけを即時実行します。
 * インストール型トリガーが未設定・旧スプレッドシート向けになっている場合でも、
 * 登録日時などの内部列が止まらないための安全策です。
 * メール送信処理は呼びません。
 */
function onEdit(e) {
  if (!e || !e.range) return;

  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== REPORTS_SHEET) return;
    if (!isDirectInputEditRange_(e.range)) return;

    completeDirectInputRowsForRange_(sheet, e.range);
    SpreadsheetApp.flush();
  } catch (error) {
    console.error('直接入力行の即時補完エラー', error);
  }
}

/**
 * 権限付きのインストール型編集トリガーです。
 * DailyReports の内部列補完に加え、DailySummaries と集計ダッシュボードを更新します。
 * この処理ではメール送信関数を呼びません。
 */
function handleSpreadsheetEdit(e) {
  if (!e || !e.range) return;

  try {
    const sheet = e.range.getSheet();
    const spreadsheet = e.source || sheet.getParent();
    const sheetName = sheet.getName();

    if (sheetName === ANALYTICS_SHEET) {
      if (
        rangeContainsCell_(e.range, ANALYTICS_MODE_CELL) ||
        rangeContainsCell_(e.range, ANALYTICS_BASE_DATE_CELL) ||
        rangeContainsCell_(e.range, ANALYTICS_CUSTOM_START_CELL) ||
        rangeContainsCell_(e.range, ANALYTICS_CUSTOM_END_CELL) ||
        rangeContainsCell_(e.range, ANALYTICS_DATE_BASIS_CELL) ||
        rangeContainsCell_(e.range, ANALYTICS_USER_FILTER_CELL)
      ) {
        refreshAnalyticsDashboard_(spreadsheet);
      }
      return;
    }

    if (sheetName === PAYROLL_SUMMARY_SHEET) {
      if (rangeContainsCell_(e.range, PAYROLL_CONFIRM_CELL)) {
        const isChecked = sheet.getRange(PAYROLL_CONFIRM_CELL).getValue() === true;
        if (isChecked) {
          try {
            const result = confirmPayrollSummaryToHistory_(spreadsheet, sheet);
            spreadsheet.toast(
              `${result.savedCount}名分を給与確定履歴へ保存しました。`,
              '給与確定',
              5
            );
          } catch (error) {
            spreadsheet.toast(
              String((error && error.message) || error || '給与確定に失敗しました。'),
              '給与確定エラー',
              8
            );
            throw error;
          } finally {
            // スクリプトによる値変更では onEdit は再発火しません。
            sheet.getRange(PAYROLL_CONFIRM_CELL).setValue(false);
          }
        }
        return;
      }

      if (
        rangeContainsCell_(e.range, PAYROLL_MONTH_CELL) ||
        rangeContainsCell_(e.range, PAYROLL_USER_FILTER_CELL)
      ) {
        refreshPayrollSummary_(spreadsheet, { skipCapture: true });
        return;
      }

      if (savePayrollDetailAdjustedPayFromRange_(spreadsheet, sheet, e.range)) {
        refreshPayrollSummary_(spreadsheet, { skipCapture: true });
      }
      return;
    }

    if (sheetName === PAYROLL_FINALIZATION_SHEET) {
      processPayrollFinalizationRowsForRange_(spreadsheet, sheet, e.range);
      refreshPayrollSummaryIfExists_(spreadsheet);
      return;
    }

    if (sheetName === PAYROLL_MISC_ADJUSTMENT_SHEET) {
      completePayrollMiscAdjustmentRowsForRange_(sheet, e.range);
      refreshPayrollSummaryIfExists_(spreadsheet);
      return;
    }

    if (
      sheetName === HOURLY_RATE_MASTER_SHEET ||
      sheetName === PAYROLL_ADJUSTMENT_SHEET
    ) {
      refreshPayrollSummaryIfExists_(spreadsheet);
      return;
    }

    if (sheetName !== REPORTS_SHEET) return;
    if (!isDirectInputEditRange_(e.range)) return;

    completeDirectInputRowsForRange_(sheet, e.range);
    SpreadsheetApp.flush();
    rebuildDirectInputSummaries_(spreadsheet);
    refreshAnalyticsDashboardIfExists_(spreadsheet);
    refreshPayrollSummaryIfExists_(spreadsheet);
  } catch (error) {
    console.error('スプレッドシート直接入力の自動更新エラー', error);
  }
}

function isDirectInputEditRange_(range) {
  if (!range || range.getLastRow() < 2) return false;

  const inputStartCol = REPORT_COL_WORK_DATE;
  const inputEndCol = REPORT_COL_DIRECT_DIARY;
  return !(
    range.getLastColumn() < inputStartCol ||
    range.getColumn() > inputEndCol
  );
}

function completeDirectInputRowsForRange_(sheet, range) {
  const firstRow = Math.max(2, range.getRow());
  const lastRow = range.getLastRow();

  for (let row = firstRow; row <= lastRow; row++) {
    completeDirectInputRow_(sheet, row);
  }
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

  const triggerResult = ensureDirectInputEditTrigger_(ss);
  const syncResult = syncDirectInputAndAnalytics_(ss);
  const triggerText =
    '\n\n自動更新トリガーを再設定しました。' +
    (triggerResult.deletedCount
      ? `（旧トリガー ${triggerResult.deletedCount}件を削除）`
      : '');
  const syncText = buildDirectInputSyncResultText_(syncResult);

  SpreadsheetApp.getUi().alert(
    'テーブル対応の直接入力設定を確認しました。\n\n' +
    'この処理では、テーブルの列型・入力規則・見出し・メモを変更しません。\n' +
    'B～F列を入力すると、A列・G列・H列を自動補完します。\n' +
    'I列の「一日の総括」は任意です。\n' +
    'スプレッドシートから直接入力した場合、メールは送信されません。' +
    triggerText +
    '\n\n' + syncText +
    warningText
  );
}

/**
 * 権限付きのインストール型編集トリガーを現在のスプレッドシート向けに再設定します。
 * シンプルトリガーは内部列の即時補完、こちらは別シート更新とダッシュボード更新を担当します。
 */
function ensureDirectInputEditTrigger_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('自動更新対象のスプレッドシートを取得できませんでした。');
  }

  const handlerName = 'handleSpreadsheetEdit';
  let deletedCount = 0;

  // コピー前のスプレッドシートを参照している旧トリガーや、重複トリガーを残さないため、
  // このハンドラーの編集トリガーをすべて削除し、現在のスプレッドシートへ作り直します。
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (
      trigger.getHandlerFunction() === handlerName &&
      trigger.getEventType() === ScriptApp.EventType.ON_EDIT
    ) {
      ScriptApp.deleteTrigger(trigger);
      deletedCount += 1;
    }
  });

  ScriptApp.newTrigger(handlerName)
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  return {
    created: true,
    deletedCount,
    spreadsheetId: ss.getId(),
  };
}

/**
 * 既に入力済みの直接入力行も含め、内部列・日別集計・ダッシュボードを同期します。
 */
function syncDirectInputAndAnalytics_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('同期対象のスプレッドシートを取得できませんでした。');
  }
  const reportsSheet = ss.getSheetByName(REPORTS_SHEET);
  if (!reportsSheet) {
    throw new Error('DailyReportsシートが見つかりません。');
  }

  const result = createDirectInputSyncResult_();
  const lastRow = reportsSheet.getLastRow();
  for (let row = 2; row <= lastRow; row++) {
    try {
      const rowResult = completeDirectInputRow_(reportsSheet, row, { forceRepair: true });
      addDirectInputRowResult_(result, rowResult);
    } catch (error) {
      result.checked += 1;
      result.errors += 1;
      result.errorRows.push(row);
      console.error(`DailyReports ${row}行目の補完エラー`, error);
    }
  }

  SpreadsheetApp.flush();
  rebuildDirectInputSummaries_(ss);
  refreshAnalyticsDashboardIfExists_(ss);
  refreshPayrollSummaryIfExists_(ss);
  return result;
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
  const result = syncDirectInputAndAnalytics_();
  SpreadsheetApp.getUi().alert(
    '直接入力行の補完、日別集計、集計ダッシュボードの更新が完了しました。\n\n' +
    buildDirectInputSyncResultText_(result)
  );
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
  refreshAnalyticsDashboardIfExists_();
  refreshPayrollSummaryIfExists_();

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


function setRequestNoEnabled(token, requestNo, enabled) {
  const session = getSession_(token);
  if (!session.ok) return session;

  requestNo = String(requestNo || '').trim();
  if (!requestNo) {
    return { ok: false, message: '依頼No.が不正です。' };
  }

  const targetEnabled = enabled === true || String(enabled).toLowerCase() === 'true';
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
        sheet.getRange(i + 1, REQUEST_COL_ENABLED).setValue(targetEnabled);
        return {
          ok: true,
          message: targetEnabled
            ? '依頼No.を有効化しました。'
            : '依頼No.を無効化しました。',
          items: getRequestMasterItems_(),
        };
      }
    }

    return { ok: false, message: '対象の依頼No.が見つかりません。' };
  } catch (e) {
    console.error(e);
    return {
      ok: false,
      message: e.message || '依頼No.の状態変更時にエラーが発生しました。',
    };
  } finally {
    lock.releaseLock();
  }
}

function updateRequestNo(token, oldRequestNo, newRequestNo) {
  const session = getSession_(token);
  if (!session.ok) return session;

  oldRequestNo = String(oldRequestNo || '').trim();
  newRequestNo = String(newRequestNo || '').trim();

  if (!oldRequestNo || !newRequestNo) {
    return { ok: false, message: '依頼No.が不正です。' };
  }

  if (oldRequestNo === newRequestNo) {
    return {
      ok: true,
      message: '依頼No.は変更されていません。',
      items: getRequestMasterItems_(),
    };
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
    let targetRow = 0;

    for (let i = 1; i < values.length; i++) {
      const rowRequestNo = String(values[i][REQUEST_COL_NO - 1] || '').trim();

      if (rowRequestNo === newRequestNo) {
        return { ok: false, message: '変更後の依頼No.は既に登録されています。' };
      }

      if (rowRequestNo === oldRequestNo) {
        targetRow = i + 1;
      }
    }

    if (!targetRow) {
      return { ok: false, message: '変更対象の依頼No.が見つかりません。' };
    }

    sheet.getRange(targetRow, REQUEST_COL_NO).setValue(newRequestNo);

    return {
      ok: true,
      message: '依頼No.を変更しました。',
      items: getRequestMasterItems_(),
    };
  } catch (e) {
    console.error(e);
    return {
      ok: false,
      message: e.message || '依頼No.変更時にエラーが発生しました。',
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


function buildDayHistory_(name, historyDate) {
  const reportRows = getHistoryReportRowsForDate_(name, historyDate);
  const appSummaryRows = getAppSummaryRowsForHistoryDate_(name, historyDate);

  const dailyTotal = Number(
    reportRows.reduce((sum, row) => sum + Number(row.workHours || 0), 0).toFixed(2)
  );

  const diaryItems = [];

  appSummaryRows.forEach((row) => {
    const text = String(row.diary || '').trim();
    if (text) {
      diaryItems.push({
        submittedAt: row.submittedAt,
        text,
      });
    }
  });

  reportRows.forEach((row) => {
    if (row.inputSource !== 'direct') return;
    const text = String(row.diary || '').trim();
    if (!text) return;

    const duplicated = diaryItems.some((item) => item.text === text);
    if (!duplicated) {
      diaryItems.push({
        submittedAt: row.submittedAt,
        text,
      });
    }
  });

  diaryItems.sort((a, b) => String(a.submittedAt).localeCompare(String(b.submittedAt)));
  const diary = diaryItems.map((item) => item.text).join('\n\n');

  const submittedCandidates = reportRows
    .map((row) => row.submittedAt)
    .concat(appSummaryRows.map((row) => row.submittedAt))
    .filter(Boolean)
    .sort();

  const submittedAt = submittedCandidates.length
    ? submittedCandidates[submittedCandidates.length - 1]
    : toIsoString_(new Date());

  const hasAppRows = reportRows.some((row) => row.inputSource === 'app');
  const hasDirectRows = reportRows.some((row) => row.inputSource === 'direct');
  const appReportCount = appSummaryRows.length || (hasAppRows ? 1 : 0);
  const reportCount = appReportCount + (hasDirectRows ? 1 : 0);

  const entries = reportRows.length
    ? [
        {
          reportId: buildDailyReportId_(historyDate, name),
          submittedAt,
          inputTotal: dailyTotal,
          cumulativeTotal: dailyTotal,
          diary,
          rows: reportRows.map((row) => ({
            submittedAt: row.submittedAt,
            requestNo: row.requestNo,
            workContent: row.workContent,
            workHours: row.workHours,
            reportId: row.reportId,
            sheetRow: row.sheetRow,
          })),
        },
      ]
    : [];

  return {
    name,
    workDate: formatJapaneseDate_(historyDate),
    dailyTotal,
    entryCount: reportCount,
    entries,
  };
}

function buildWeekHistory_(name, baseDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportsSheet = ss.getSheetByName(REPORTS_SHEET);
  const summariesSheet = ss.getSheetByName(SUMMARIES_SHEET);

  const startDate = getOffsetDateString_(baseDate, -6);
  const endDate = baseDate;
  const daysMap = {};

  for (let i = 0; i < 7; i++) {
    const date = getOffsetDateString_(startDate, i);
    daysMap[date] = {
      workDate: date,
      totalHours: 0,
      appSummaryCount: 0,
      hasAppRows: false,
      hasDirectRows: false,
      diaries: [],
      latestSubmittedAt: '',
    };
  }

  // 作業時間はDailyReports本体から集計します。
  // これにより、直接入力行がDailySummariesへ未反映でも履歴に表示されます。
  if (reportsSheet && reportsSheet.getLastRow() >= 2) {
    const reportValues = reportsSheet.getDataRange().getValues();

    for (let i = 1; i < reportValues.length; i++) {
      const row = reportValues[i];
      const rowName = String(row[REPORT_COL_NAME - 1] || '').trim();
      if (rowName !== name) continue;

      const historyInfo = getHistoryDateInfoForReportRow_(row);
      const historyDate = historyInfo.historyDate;
      if (!historyDate || !isDateInRange_(historyDate, startDate, endDate)) continue;

      const requestNo = String(row[REPORT_COL_REQUEST_NO - 1] || '').trim();
      const workContent = String(row[REPORT_COL_WORK_CONTENT - 1] || '').trim();
      const hours = Number(row[REPORT_COL_WORK_HOURS - 1] || 0);
      if (!requestNo || !workContent || !(hours > 0)) continue;

      const day = daysMap[historyDate];
      if (!day) continue;

      day.totalHours = Number((day.totalHours + hours).toFixed(2));
      if (historyInfo.inputSource === 'app') {
        day.hasAppRows = true;
      } else {
        day.hasDirectRows = true;

        const diary = String(row[REPORT_COL_DIRECT_DIARY - 1] || '').trim();
        if (diary && !day.diaries.some((item) => item.text === diary)) {
          day.diaries.push({
            submittedAt: historyInfo.submittedAt,
            text: diary,
          });
        }
      }

      if (
        historyInfo.submittedAt &&
        (!day.latestSubmittedAt || historyInfo.submittedAt > day.latestSubmittedAt)
      ) {
        day.latestSubmittedAt = historyInfo.submittedAt;
      }
    }
  }

  // アプリ入力の総括と送信回数はDailySummariesから取得します。
  // 直接入力分のDailySummariesは、登録日キーとずれる可能性があるため使用しません。
  if (summariesSheet && summariesSheet.getLastRow() >= 2) {
    const summaryValues = summariesSheet.getDataRange().getValues();

    for (let i = 1; i < summaryValues.length; i++) {
      const row = summaryValues[i];
      const rowName = String(row[SUMMARY_COL_NAME - 1] || '').trim();
      const token = String(row[SUMMARY_COL_TOKEN - 1] || '').trim();
      if (rowName !== name || token === DIRECT_INPUT_TOKEN) continue;

      const rowDate = normalizeSheetWorkDate_(row[SUMMARY_COL_WORK_DATE - 1]);
      if (!rowDate || !isDateInRange_(rowDate, startDate, endDate)) continue;

      const day = daysMap[rowDate];
      if (!day) continue;

      const submittedAt = toIsoString_(row[SUMMARY_COL_SUBMITTED_AT - 1]);
      const diary = String(row[SUMMARY_COL_DIARY - 1] || '').trim();

      day.appSummaryCount += 1;
      if (diary) {
        day.diaries.push({ submittedAt, text: diary });
      }

      if (!day.latestSubmittedAt || submittedAt > day.latestSubmittedAt) {
        day.latestSubmittedAt = submittedAt;
      }
    }
  }

  const days = Object.values(daysMap)
    .sort((a, b) => b.workDate.localeCompare(a.workDate))
    .map((day) => {
      const reportCount =
        (day.appSummaryCount || (day.hasAppRows ? 1 : 0)) +
        (day.hasDirectRows ? 1 : 0);

      return {
        workDate: formatJapaneseDate_(day.workDate),
        workDateValue: day.workDate,
        totalHours: day.totalHours,
        reportCount,
        latestDiary: day.diaries
          .sort((a, b) => String(a.submittedAt).localeCompare(String(b.submittedAt)))
          .map((item) => item.text)
          .filter((text, index, array) => array.indexOf(text) === index)
          .join('\n\n'),
        latestSubmittedAt: day.latestSubmittedAt,
      };
    });

  return {
    name,
    baseDate: formatJapaneseDate_(baseDate),
    startDate: formatJapaneseDate_(startDate),
    endDate: formatJapaneseDate_(endDate),
    days,
  };
}

/**
 * 履歴表示用の日付と入力元を判定します。
 * アプリ入力は従来どおりB列の作業日を使用します。
 * スプレッドシート直接入力は、A列の登録日時があればその日付を優先し、
 * A列が空欄・不正な場合のみB列の作業日へフォールバックします。
 */
function getHistoryDateInfoForReportRow_(row) {
  const workDate = normalizeSheetWorkDate_(row[REPORT_COL_WORK_DATE - 1]);
  const name = String(row[REPORT_COL_NAME - 1] || '').trim();
  const submittedAtValue = row[REPORT_COL_SUBMITTED_AT - 1];
  const submittedAt = toIsoString_(submittedAtValue);
  const token = String(row[REPORT_COL_TOKEN - 1] || '').trim();
  const reportId = String(row[REPORT_COL_REPORT_ID - 1] || '').trim();
  const expectedReportId = workDate && name ? buildDailyReportId_(workDate, name) : '';

  const isAppRow = isConfirmedAppReportRow_(
    submittedAtValue,
    token,
    reportId,
    expectedReportId
  );

  const registeredDate = hasValidSubmittedAt_(submittedAtValue)
    ? normalizeSheetWorkDate_(submittedAtValue)
    : '';

  return {
    inputSource: isAppRow ? 'app' : 'direct',
    historyDate: isAppRow ? workDate : (registeredDate || workDate),
    submittedAt,
    workDate,
  };
}

function getHistoryReportRowsForDate_(name, historyDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowName = String(row[REPORT_COL_NAME - 1] || '').trim();
    if (rowName !== name) continue;

    const historyInfo = getHistoryDateInfoForReportRow_(row);
    if (historyInfo.historyDate !== historyDate) continue;

    const requestNo = String(row[REPORT_COL_REQUEST_NO - 1] || '').trim();
    const workContent = String(row[REPORT_COL_WORK_CONTENT - 1] || '').trim();
    const hours = Number(row[REPORT_COL_WORK_HOURS - 1] || 0);
    if (!requestNo || !workContent || !(hours > 0)) continue;

    rows.push({
      submittedAt: historyInfo.submittedAt,
      requestNo,
      workContent,
      workHours: Number(hours.toFixed ? hours.toFixed(2) : hours),
      reportId: String(row[REPORT_COL_REPORT_ID - 1] || '').trim(),
      diary: String(row[REPORT_COL_DIRECT_DIARY - 1] || '').trim(),
      inputSource: historyInfo.inputSource,
      sourceWorkDate: historyInfo.workDate,
      sheetRow: i + 1,
    });
  }

  return rows.sort((a, b) => {
    const timeCompare = String(a.submittedAt).localeCompare(String(b.submittedAt));
    return timeCompare || a.sheetRow - b.sheetRow;
  });
}

function getAppSummaryRowsForHistoryDate_(name, historyDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SUMMARIES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const rowDate = normalizeSheetWorkDate_(values[i][SUMMARY_COL_WORK_DATE - 1]);
    const rowName = String(values[i][SUMMARY_COL_NAME - 1] || '').trim();
    const token = String(values[i][SUMMARY_COL_TOKEN - 1] || '').trim();

    if (rowDate !== historyDate || rowName !== name || token === DIRECT_INPUT_TOKEN) continue;

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

/**
 * DailyReports の直接入力行に、内部管理用の値を補完します。
 */
function completeDirectInputRow_(sheet, rowNumber, options) {
  const opts = options || {};
  const rowValues = sheet
    .getRange(rowNumber, 1, 1, REPORT_COL_DIRECT_DIARY)
    .getValues()[0];

  const submittedAt = rowValues[REPORT_COL_SUBMITTED_AT - 1];
  const workDate = normalizeSheetWorkDate_(rowValues[REPORT_COL_WORK_DATE - 1]);
  const name = String(rowValues[REPORT_COL_NAME - 1] || '').trim();
  const requestNo = String(rowValues[REPORT_COL_REQUEST_NO - 1] || '').trim();
  const workContent = String(rowValues[REPORT_COL_WORK_CONTENT - 1] || '').trim();
  const rawHours = rowValues[REPORT_COL_WORK_HOURS - 1];
  const workHours = Number(rawHours || 0);
  const token = String(rowValues[REPORT_COL_TOKEN - 1] || '').trim();
  const reportId = String(rowValues[REPORT_COL_REPORT_ID - 1] || '').trim();
  const diary = String(rowValues[REPORT_COL_DIRECT_DIARY - 1] || '').trim();

  const hasInput = Boolean(
    workDate || name || requestNo || workContent || hasMeaningfulHours_(rawHours) || diary
  );

  const expectedReportId = workDate && name
    ? buildDailyReportId_(workDate, name)
    : '';

  // アプリ登録済み行として除外するのは、A・G・H列がすべて揃い、
  // G列がアプリで発行するUUID形式、かつH列が現在の作業日・氏名と一致する行だけです。
  // G列に古い値・コピー値だけが残っている行は、直接入力として修復します。
  const isConfirmedAppRow = isConfirmedAppReportRow_(
    submittedAt,
    token,
    reportId,
    expectedReportId
  );
  if (isConfirmedAppRow) {
    return { status: 'app', rowNumber };
  }

  if (!hasInput) {
    if (token === DIRECT_INPUT_TOKEN || opts.forceRepair) {
      clearDirectInputInternalCells_(sheet, rowNumber);
    }
    return { status: 'empty', rowNumber };
  }

  const isComplete = Boolean(
    workDate && name && requestNo && workContent && Number.isFinite(workHours) && workHours > 0
  );
  if (!isComplete) {
    return { status: 'incomplete', rowNumber };
  }

  const needsSubmittedAt = !hasValidSubmittedAt_(submittedAt);
  const hasStaleInternalValue = Boolean(
    (token && token !== DIRECT_INPUT_TOKEN) ||
    (reportId && reportId !== expectedReportId)
  );
  const needsTokenOrId = token !== DIRECT_INPUT_TOKEN || reportId !== expectedReportId;

  // テーブルの型付き列に合わせ、A列はDate、G・H列は文字列として書き込みます。
  if (needsSubmittedAt) {
    sheet.getRange(rowNumber, REPORT_COL_SUBMITTED_AT).setValue(new Date());
  }

  if (needsTokenOrId) {
    sheet
      .getRange(rowNumber, REPORT_COL_TOKEN, 1, 2)
      .setValues([[DIRECT_INPUT_TOKEN, expectedReportId]]);
  }

  return {
    status: needsSubmittedAt || needsTokenOrId ? 'completed' : 'alreadyComplete',
    rowNumber,
    repairedStaleInternalValue: hasStaleInternalValue,
  };
}

function hasMeaningfulHours_(value) {
  if (value === '' || value === null || value === undefined) return false;
  const number = Number(value);
  return Number.isFinite(number) && number !== 0;
}

function hasValidSubmittedAt_(value) {
  if (value instanceof Date) return !isNaN(value.getTime());
  if (!value) return false;
  const date = new Date(value);
  return !isNaN(date.getTime());
}

function isAppSessionToken_(token) {
  const value = String(token || '').trim();
  if (!value || value === DIRECT_INPUT_TOKEN) return false;
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}

function isConfirmedAppReportRow_(submittedAt, token, reportId, expectedReportId) {
  return Boolean(
    hasValidSubmittedAt_(submittedAt) &&
    isAppSessionToken_(token) &&
    expectedReportId &&
    String(reportId || '').trim() === expectedReportId
  );
}

function clearDirectInputInternalCells_(sheet, rowNumber) {
  sheet.getRange(rowNumber, REPORT_COL_SUBMITTED_AT).clearContent();
  sheet.getRange(rowNumber, REPORT_COL_TOKEN, 1, 2).clearContent();
}

function createDirectInputSyncResult_() {
  return {
    checked: 0,
    completed: 0,
    alreadyComplete: 0,
    repairedStaleInternalValue: 0,
    app: 0,
    incomplete: 0,
    empty: 0,
    errors: 0,
    errorRows: [],
  };
}

function addDirectInputRowResult_(result, rowResult) {
  if (!result || !rowResult) return;
  result.checked += 1;
  const status = String(rowResult.status || '');
  if (Object.prototype.hasOwnProperty.call(result, status)) {
    result[status] += 1;
  }
  if (rowResult.repairedStaleInternalValue) {
    result.repairedStaleInternalValue += 1;
  }
}

function buildDirectInputSyncResultText_(result) {
  const data = result || createDirectInputSyncResult_();
  const lines = [
    `確認行数：${data.checked}行`,
    `今回補完：${data.completed}行`,
    `補完済み：${data.alreadyComplete}行`,
    `古い内部値を修復：${data.repairedStaleInternalValue}行`,
    `アプリ登録行として保持：${data.app}行`,
    `入力途中のため保留：${data.incomplete}行`,
    `補完エラー：${data.errors || 0}行`,
  ];
  if (data.errorRows && data.errorRows.length) {
    lines.push(`エラー行：${data.errorRows.join(', ')}`);
  }
  return lines.join('\n');
}

/**
 * 直接入力行を日付・氏名ごとにまとめ、DailySummariesへ反映します。
 * アプリから登録されたDailySummaries行は保持します。
 */
function rebuildDirectInputSummaries_(spreadsheet) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('集計対象のスプレッドシートを取得できませんでした。');
    }
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

      const submittedAt = row[REPORT_COL_SUBMITTED_AT - 1];
      const token = String(row[REPORT_COL_TOKEN - 1] || '').trim();
      const reportId = String(row[REPORT_COL_REPORT_ID - 1] || '').trim();
      const expectedReportId = buildDailyReportId_(workDate, name);

      // G列が空欄・古い値でも、A/G/Hが正しく揃ったアプリ行でなければ
      // スプレッドシート直接入力として日別集計へ含めます。
      if (isConfirmedAppReportRow_(submittedAt, token, reportId, expectedReportId)) {
        continue;
      }

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


/**
 * 月・年・期間指定・Totalを切り替えて確認できる集計ダッシュボードを作成します。
 */
function setupAnalyticsDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ANALYTICS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(ANALYTICS_SHEET);
    const today = new Date();
    sheet.getRange(ANALYTICS_MODE_CELL).setValue('月');
    sheet.getRange(ANALYTICS_BASE_DATE_CELL).setValue(today);
    sheet.getRange(ANALYTICS_CUSTOM_START_CELL).setValue(new Date(today.getFullYear(), today.getMonth(), 1));
    sheet.getRange(ANALYTICS_CUSTOM_END_CELL).setValue(today);
    sheet.getRange(ANALYTICS_DATE_BASIS_CELL).setValue('登録日時');
    sheet.getRange(ANALYTICS_USER_FILTER_CELL).setValue('全員');
  }

  refreshAnalyticsDashboard_(ss);
  ss.setActiveSheet(sheet);

  SpreadsheetApp.getUi().alert(
    '集計ダッシュボードを作成・更新しました。\n\n' +
    'B2で「月・年・期間指定・Total」を選択できます。\n' +
    '月・年はB3の基準日、期間指定はB4の開始日とB5の終了日を使用します。\n' +
    'B6で集計日基準を「登録日時・作業日」から選択できます。\n' +
    'E2で表示対象を「全員」または個人名から選択できます。\n' +
    '選択内容を変更すると自動で集計が切り替わります。'
  );
}

/**
 * メニューから手動更新するための公開関数です。
 */
function refreshAnalyticsDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(ANALYTICS_SHEET)) {
    setupAnalyticsDashboard();
    return;
  }

  refreshAnalyticsDashboard_(ss);
  SpreadsheetApp.getUi().alert('集計ダッシュボードを更新しました。');
}

/**
 * ダッシュボードが作成済みの場合のみ、安全に更新します。
 * 集計更新の失敗によって日報保存自体が止まらないようにします。
 */
function refreshAnalyticsDashboardIfExists_(spreadsheet) {
  try {
    const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getSheetByName(ANALYTICS_SHEET)) {
      refreshAnalyticsDashboard_(ss);
    }
  } catch (e) {
    console.error('集計ダッシュボード更新エラー', e);
  }
}

function refreshAnalyticsDashboard_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('集計対象のスプレッドシートを取得できませんでした。');
  }
  const sheet = ss.getSheetByName(ANALYTICS_SHEET);
  if (!sheet) return;

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const currentMode = normalizeAnalyticsMode_(sheet.getRange(ANALYTICS_MODE_CELL).getDisplayValue());
  const currentBaseValue = sheet.getRange(ANALYTICS_BASE_DATE_CELL).getValue();
  const currentCustomStartValue = sheet.getRange(ANALYTICS_CUSTOM_START_CELL).getValue();
  const currentCustomEndValue = sheet.getRange(ANALYTICS_CUSTOM_END_CELL).getValue();
  const currentDateBasis = normalizeAnalyticsDateBasis_(
    sheet.getRange(ANALYTICS_DATE_BASIS_CELL).getDisplayValue()
  );
  const currentUserFilter = String(
    sheet.getRange(ANALYTICS_USER_FILTER_CELL).getDisplayValue() || ''
  ).trim();

  const mode = currentMode || '月';
  const baseDate = normalizeWorkDate_(currentBaseValue) || today;
  const defaultCustomStart = `${baseDate.slice(0, 7)}-01`;
  const customStartDate = normalizeWorkDate_(currentCustomStartValue) || defaultCustomStart;
  const customEndDate = normalizeWorkDate_(currentCustomEndValue) || baseDate;
  const dateBasis = currentDateBasis || '登録日時';
  const userOptions = getAnalyticsUserOptions_(ss);
  const userFilter = userOptions.includes(currentUserFilter) ? currentUserFilter : '全員';

  const period = buildAnalyticsPeriod_(mode, baseDate, customStartDate, customEndDate);
  const data = buildAnalyticsData_(period, ss, dateBasis, userFilter);

  // 既存の選択値を保持しながら、集計表示部分を作り直します。
  sheet.clear();
  ensureSheetSize_(sheet, Math.max(84, data.detailRows.length + 39), 11);

  sheet.getRange('A1:K1')
    .setBackground('#6ea86b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16);
  sheet.getRange('A1').setValue('日報 集計ダッシュボード');

  sheet.getRange('A2').setValue('表示単位');
  sheet.getRange(ANALYTICS_MODE_CELL).setValue(mode);
  const modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['月', '年', '期間指定', 'Total'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(ANALYTICS_MODE_CELL).setDataValidation(modeRule);

  sheet.getRange('A3').setValue('基準日');
  sheet.getRange(ANALYTICS_BASE_DATE_CELL).setValue(new Date(`${baseDate}T00:00:00`));
  sheet.getRange(ANALYTICS_BASE_DATE_CELL).setNumberFormat('yyyy年m月d日');

  sheet.getRange('A4').setValue('開始日');
  sheet.getRange(ANALYTICS_CUSTOM_START_CELL).setValue(new Date(`${period.customStartDate}T00:00:00`));
  sheet.getRange(ANALYTICS_CUSTOM_START_CELL).setNumberFormat('yyyy年m月d日');

  sheet.getRange('A5').setValue('終了日');
  sheet.getRange(ANALYTICS_CUSTOM_END_CELL).setValue(new Date(`${period.customEndDate}T00:00:00`));
  sheet.getRange(ANALYTICS_CUSTOM_END_CELL).setNumberFormat('yyyy年m月d日');

  sheet.getRange('A6').setValue('集計日基準');
  sheet.getRange(ANALYTICS_DATE_BASIS_CELL).setValue(dateBasis);
  const dateBasisRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['登録日時', '作業日'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(ANALYTICS_DATE_BASIS_CELL).setDataValidation(dateBasisRule);

  sheet.getRange('A7').setValue('対象期間');
  sheet.getRange(ANALYTICS_PERIOD_LABEL_CELL).setValue(data.periodLabel);

  sheet.getRange('D2').setValue('表示対象');
  sheet.getRange(ANALYTICS_USER_FILTER_CELL).setValue(userFilter);
  const userFilterRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(userOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(ANALYTICS_USER_FILTER_CELL).setDataValidation(userFilterRule);

  styleControlArea_(sheet);
  writeAnalyticsSummary_(sheet, data, 9);

  const requestHeaderRow = 13;
  const userHeaderRow = 13;
  writeRequestSummaryTable_(sheet, requestHeaderRow, data);
  writeUserSummaryTable_(sheet, userHeaderRow, data);

  const tableEndRow = Math.max(
    requestHeaderRow + 2 + Math.max(data.requestRows.length, 1),
    userHeaderRow + 2 + Math.max(data.userRows.length, 1)
  );
  const detailStartRow = tableEndRow + 3;

  writeUserDetailTable_(sheet, detailStartRow, data);
  writeRequestUserDetailTable_(sheet, detailStartRow, data);

  sheet.setFrozenRows(7);
  sheet.setColumnWidth(1, 54);
  sheet.setColumnWidth(2, 185);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 145);
  sheet.setColumnWidth(7, 54);
  sheet.setColumnWidth(8, 145);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 90);

  // 集計ダッシュボードのA列は、見出し・項目名・値をすべて左詰めで表示します。
  sheet.getRange(1, 1, sheet.getMaxRows(), 1).setHorizontalAlignment('left');
}

function buildAnalyticsPeriod_(mode, baseDate, customStartDate, customEndDate) {
  const year = baseDate.slice(0, 4);
  const month = baseDate.slice(5, 7);
  let normalizedCustomStart = normalizeWorkDate_(customStartDate) || `${year}-${month}-01`;
  let normalizedCustomEnd = normalizeWorkDate_(customEndDate) || baseDate;

  if (normalizedCustomStart > normalizedCustomEnd) {
    const temporary = normalizedCustomStart;
    normalizedCustomStart = normalizedCustomEnd;
    normalizedCustomEnd = temporary;
  }

  if (mode === 'Total') {
    return {
      mode,
      startDate: '',
      endDate: '',
      customStartDate: normalizedCustomStart,
      customEndDate: normalizedCustomEnd,
      label: '全期間',
    };
  }

  if (mode === '期間指定') {
    return {
      mode,
      startDate: normalizedCustomStart,
      endDate: normalizedCustomEnd,
      customStartDate: normalizedCustomStart,
      customEndDate: normalizedCustomEnd,
      label: `${formatJapaneseDate_(normalizedCustomStart)} ～ ${formatJapaneseDate_(normalizedCustomEnd)}`,
    };
  }

  if (mode === '年') {
    return {
      mode,
      startDate: `${year}-01-01`,
      endDate: `${year}-12-31`,
      customStartDate: normalizedCustomStart,
      customEndDate: normalizedCustomEnd,
      label: `${Number(year)}年`,
    };
  }

  const lastDate = new Date(Number(year), Number(month), 0);
  const endDate = Utilities.formatDate(lastDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return {
    mode: '月',
    startDate: `${year}-${month}-01`,
    endDate,
    customStartDate: normalizedCustomStart,
    customEndDate: normalizedCustomEnd,
    label: `${Number(year)}年${Number(month)}月`,
  };
}

function buildAnalyticsData_(period, spreadsheet, dateBasisValue, userFilterValue) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('集計対象のスプレッドシートを取得できませんでした。');
  }
  const reportsSheet = ss.getSheetByName(REPORTS_SHEET);
  if (!reportsSheet) {
    throw new Error('DailyReportsシートが見つかりません。');
  }

  const values = reportsSheet.getDataRange().getValues();
  const requestMap = {};
  const userMap = {};
  const userDetailMap = {};
  const requestUserMap = {};
  const dateBasis = normalizeAnalyticsDateBasis_(dateBasisValue) || '登録日時';
  const userFilter = String(userFilterValue || '全員').trim() || '全員';
  let totalHours = 0;
  let recordCount = 0;
  let minDate = '';
  let maxDate = '';

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const submittedDate = normalizeWorkDate_(row[REPORT_COL_SUBMITTED_AT - 1]);
    const workDate = normalizeSheetWorkDate_(row[REPORT_COL_WORK_DATE - 1]);
    const name = String(row[REPORT_COL_NAME - 1] || '').trim();
    const requestNo = String(row[REPORT_COL_REQUEST_NO - 1] || '').trim();
    const workContent = String(row[REPORT_COL_WORK_CONTENT - 1] || '').trim();
    const hours = Number(row[REPORT_COL_WORK_HOURS - 1] || 0);
    const aggregateDate = dateBasis === '登録日時' ? submittedDate : workDate;

    if (!aggregateDate || !name || !requestNo || !workContent || !(hours > 0)) continue;
    if (userFilter !== '全員' && name !== userFilter) continue;
    if (period.mode !== 'Total' && !isDateInRange_(aggregateDate, period.startDate, period.endDate)) {
      continue;
    }

    recordCount += 1;
    totalHours = Number((totalHours + hours).toFixed(2));
    if (!minDate || aggregateDate < minDate) minDate = aggregateDate;
    if (!maxDate || aggregateDate > maxDate) maxDate = aggregateDate;

    if (!requestMap[requestNo]) {
      requestMap[requestNo] = { requestNo, hours: 0, users: {} };
    }
    requestMap[requestNo].hours = Number((requestMap[requestNo].hours + hours).toFixed(2));
    requestMap[requestNo].users[name] = true;

    if (!userMap[name]) {
      userMap[name] = { name, hours: 0, requests: {} };
    }
    userMap[name].hours = Number((userMap[name].hours + hours).toFixed(2));
    userMap[name].requests[requestNo] = true;

    const userDetailKey = `${name}|||${requestNo}|||${workContent}`;
    if (!userDetailMap[userDetailKey]) {
      userDetailMap[userDetailKey] = { name, requestNo, workContent, hours: 0 };
    }
    userDetailMap[userDetailKey].hours = Number(
      (userDetailMap[userDetailKey].hours + hours).toFixed(2)
    );

    const requestUserKey = `${requestNo}|||${name}`;
    if (!requestUserMap[requestUserKey]) {
      requestUserMap[requestUserKey] = { requestNo, name, hours: 0 };
    }
    requestUserMap[requestUserKey].hours = Number(
      (requestUserMap[requestUserKey].hours + hours).toFixed(2)
    );
  }

  const requestRows = Object.values(requestMap)
    .sort((a, b) => b.hours - a.hours || a.requestNo.localeCompare(b.requestNo, 'ja'))
    .map((item, index) => [
      index + 1,
      item.requestNo,
      item.hours,
      Object.keys(item.users).length,
      totalHours > 0 ? item.hours / totalHours : 0,
    ]);

  const userRows = Object.values(userMap)
    .sort((a, b) => b.hours - a.hours || a.name.localeCompare(b.name, 'ja'))
    .map((item, index) => [
      index + 1,
      item.name,
      item.hours,
      Object.keys(item.requests).length,
      totalHours > 0 ? item.hours / totalHours : 0,
    ]);

  const detailRows = Object.values(userDetailMap)
    .sort((a, b) =>
      a.name.localeCompare(b.name, 'ja') ||
      b.hours - a.hours ||
      a.requestNo.localeCompare(b.requestNo, 'ja') ||
      a.workContent.localeCompare(b.workContent, 'ja')
    )
    .map((item) => [item.name, item.requestNo, item.workContent, item.hours]);

  const requestUserRows = Object.values(requestUserMap)
    .sort((a, b) =>
      a.requestNo.localeCompare(b.requestNo, 'ja') ||
      b.hours - a.hours ||
      a.name.localeCompare(b.name, 'ja')
    )
    .map((item) => [item.requestNo, item.name, item.hours]);

  let periodLabel = period.label;
  if (period.mode === 'Total' && minDate && maxDate) {
    periodLabel = `全期間（${formatJapaneseDate_(minDate)} ～ ${formatJapaneseDate_(maxDate)}）`;
  }

  return {
    totalHours,
    recordCount,
    requestCount: requestRows.length,
    userCount: userRows.length,
    requestRows,
    userRows,
    detailRows,
    requestUserRows,
    periodLabel,
    dateBasis,
    userFilter,
  };
}

function writeAnalyticsSummary_(sheet, data, startRow) {
  const row = Number(startRow || 6);
  const labels = [
    ['合計作業時間', '', '案件数', '', '人数', '', '明細件数', ''],
    [`${data.totalHours} h`, '', data.requestCount, '', data.userCount, '', data.recordCount, ''],
  ];
  sheet.getRange(row, 1, 2, 8).setValues(labels);

  const labelRanges = [
    `A${row}:B${row}`,
    `C${row}:D${row}`,
    `E${row}:F${row}`,
    `G${row}:H${row}`,
  ];
  const valueRow = row + 1;
  const valueRanges = [
    `A${valueRow}:B${valueRow}`,
    `C${valueRow}:D${valueRow}`,
    `E${valueRow}:F${valueRow}`,
    `G${valueRow}:H${valueRow}`,
  ];
  labelRanges.forEach((a1) => {
    sheet.getRange(a1)
      .setBackground('#edf7eb')
      .setFontColor('#6a846d')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#d7e8d4', SpreadsheetApp.BorderStyle.SOLID);
  });
  valueRanges.forEach((a1) => {
    sheet.getRange(a1)
      .setBackground('#ffffff')
      .setFontColor('#314c35')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#d7e8d4', SpreadsheetApp.BorderStyle.SOLID);
  });
}

function writeRequestSummaryTable_(sheet, headerRow, data) {
  sheet.getRange(headerRow, 1, 1, 5)
    .setValues([['案件別集計', '', '', '', '']])
    .setBackground('#6ea86b')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.getRange(headerRow + 1, 1, 1, 5)
    .setValues([['順位', '依頼No.', '合計時間', '担当人数', '構成比']]);
  styleTableHeader_(sheet.getRange(headerRow + 1, 1, 1, 5));

  if (!data.requestRows.length) {
    sheet.getRange(headerRow + 2, 1, 1, 5).setValues([['-', '対象期間の記録はありません', '', '', '']]);
    return;
  }

  sheet.getRange(headerRow + 2, 1, data.requestRows.length, 5).setValues(data.requestRows);
  styleTableBody_(sheet.getRange(headerRow + 2, 1, data.requestRows.length, 5));
  sheet.getRange(headerRow + 2, 3, data.requestRows.length, 1).setNumberFormat('0.00" h"');
  sheet.getRange(headerRow + 2, 5, data.requestRows.length, 1).setNumberFormat('0.0%');
}

function writeUserSummaryTable_(sheet, headerRow, data) {
  sheet.getRange(headerRow, 7, 1, 5)
    .setValues([['個人別集計', '', '', '', '']])
    .setBackground('#6ea86b')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.getRange(headerRow + 1, 7, 1, 5)
    .setValues([['順位', '氏名', '合計時間', '案件数', '構成比']]);
  styleTableHeader_(sheet.getRange(headerRow + 1, 7, 1, 5));

  if (!data.userRows.length) {
    sheet.getRange(headerRow + 2, 7, 1, 5).setValues([['-', '対象期間の記録はありません', '', '', '']]);
    return;
  }

  sheet.getRange(headerRow + 2, 7, data.userRows.length, 5).setValues(data.userRows);
  styleTableBody_(sheet.getRange(headerRow + 2, 7, data.userRows.length, 5));
  sheet.getRange(headerRow + 2, 9, data.userRows.length, 1).setNumberFormat('0.00" h"');
  sheet.getRange(headerRow + 2, 11, data.userRows.length, 1).setNumberFormat('0.0%');
}

function writeUserDetailTable_(sheet, startRow, data) {
  sheet.getRange(startRow, 1, 1, 4)
    .setValues([['個人別 作業内訳', '', '', '']])
    .setBackground('#8cc68a')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.getRange(startRow + 1, 1, 1, 4)
    .setValues([['氏名', '依頼No.', '作業内容', '作業時間']]);
  styleTableHeader_(sheet.getRange(startRow + 1, 1, 1, 4));

  if (!data.detailRows.length) {
    sheet.getRange(startRow + 2, 1, 1, 4).setValues([['-', '対象期間の記録はありません', '', '']]);
    return;
  }

  sheet.getRange(startRow + 2, 1, data.detailRows.length, 4).setValues(data.detailRows);
  styleTableBody_(sheet.getRange(startRow + 2, 1, data.detailRows.length, 4));
  sheet.getRange(startRow + 2, 4, data.detailRows.length, 1).setNumberFormat('0.00" h"');
}

function writeRequestUserDetailTable_(sheet, startRow, data) {
  sheet.getRange(startRow, 6, 1, 3)
    .setValues([['案件別 担当者内訳', '', '']])
    .setBackground('#8cc68a')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.getRange(startRow + 1, 6, 1, 3)
    .setValues([['依頼No.', '氏名', '作業時間']]);
  styleTableHeader_(sheet.getRange(startRow + 1, 6, 1, 3));

  if (!data.requestUserRows.length) {
    sheet.getRange(startRow + 2, 6, 1, 3).setValues([['対象期間の記録はありません', '', '']]);
    return;
  }

  sheet.getRange(startRow + 2, 6, data.requestUserRows.length, 3).setValues(data.requestUserRows);
  styleTableBody_(sheet.getRange(startRow + 2, 6, data.requestUserRows.length, 3));
  sheet.getRange(startRow + 2, 8, data.requestUserRows.length, 1).setNumberFormat('0.00" h"');
}

function styleControlArea_(sheet) {
  sheet.getRange('A2:A7')
    .setBackground('#edf7eb')
    .setFontWeight('bold')
    .setFontColor('#314c35');
  sheet.getRange('B2:B7')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, true, true, '#d7e8d4', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('D2')
    .setBackground('#edf7eb')
    .setFontWeight('bold')
    .setFontColor('#314c35');
  sheet.getRange('E2')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, true, true, '#d7e8d4', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('B7').setFontWeight('bold').setFontColor('#314c35');
}

function styleTableHeader_(range) {
  range
    .setBackground('#edf7eb')
    .setFontColor('#314c35')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#d7e8d4', SpreadsheetApp.BorderStyle.SOLID);
}

function styleTableBody_(range) {
  range
    .setBackground('#ffffff')
    .setFontColor('#314c35')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#d7e8d4', SpreadsheetApp.BorderStyle.SOLID);
}


function getAnalyticsUserOptions_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const names = {};

  if (ss) {
    const usersSheet = ss.getSheetByName(USERS_SHEET);
    if (usersSheet && usersSheet.getLastRow() >= 2) {
      const userValues = usersSheet.getRange(2, USER_COL_NAME, usersSheet.getLastRow() - 1, 1).getValues();
      userValues.forEach((row) => {
        const name = String(row[0] || '').trim();
        if (name) names[name] = true;
      });
    }

    const reportsSheet = ss.getSheetByName(REPORTS_SHEET);
    if (reportsSheet && reportsSheet.getLastRow() >= 2) {
      const reportValues = reportsSheet
        .getRange(2, REPORT_COL_NAME, reportsSheet.getLastRow() - 1, 1)
        .getValues();
      reportValues.forEach((row) => {
        const name = String(row[0] || '').trim();
        if (name) names[name] = true;
      });
    }
  }

  return ['全員'].concat(Object.keys(names).sort((a, b) => a.localeCompare(b, 'ja')));
}

function normalizeAnalyticsDateBasis_(value) {
  const basis = String(value || '').trim();
  if (basis === '登録日時' || basis === '作業日') return basis;
  return '';
}

function normalizeAnalyticsMode_(value) {
  const mode = String(value || '').trim();
  if (mode === '月' || mode === '年' || mode === '期間指定' || mode === 'Total') return mode;
  return '';
}

function rangeContainsCell_(range, a1Notation) {
  const target = range.getSheet().getRange(a1Notation);
  const row = target.getRow();
  const column = target.getColumn();
  return (
    row >= range.getRow() &&
    row <= range.getLastRow() &&
    column >= range.getColumn() &&
    column <= range.getLastColumn()
  );
}

function ensureSheetSize_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
}


/**
 * 給与集計・時給マスタ・支給時間調整・給与加減算・給与確定履歴を作成します。
 * 既存シートがある場合、入力済みのマスタ・調整値は保持します。
 */
function setupPayrollSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('スプレッドシートを取得できませんでした。');
  }

  setupHourlyRateMaster_(ss);
  setupPayrollAdjustmentSheet_(ss);
  setupPayrollMiscAdjustmentSheet_(ss);
  setupPayrollFinalizationSheet_(ss);

  let summarySheet = ss.getSheetByName(PAYROLL_SUMMARY_SHEET);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(PAYROLL_SUMMARY_SHEET);
    const today = new Date();
    summarySheet.getRange(PAYROLL_MONTH_CELL).setValue(
      new Date(today.getFullYear(), today.getMonth(), 1)
    );
    summarySheet.getRange(PAYROLL_USER_FILTER_CELL).setValue('全員');
  }

  ensureDirectInputEditTrigger_(ss);
  refreshPayrollSummary_(ss, { skipCapture: true });
  ss.setActiveSheet(summarySheet);

  SpreadsheetApp.getUi().alert(
    '給与関連シートを作成・更新しました。\n\n' +
    '・給与集計：前月21日～当月20日の支給額を確認（DailyReports A列の登録日時基準）\n' +
    '・時給マスタ：スタッフ別の時給と適用期間を管理\n' +
    '・支給時間調整：案件ごとの支給対象時間またはインセンティブ時間を入力\n' +
    '・給与加減算：交通費などの天引きはマイナス、追加支給はプラスで記録\n' +
    '・給与確定履歴：最終決定支払額と変更理由を履歴として保存\n\n' +
    `初期時給は${DEFAULT_HOURLY_RATE.toLocaleString()}円で登録しています。`
  );
}

/**
 * メニューから給与集計を更新します。
 */
function refreshPayrollSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(PAYROLL_SUMMARY_SHEET)) {
    setupPayrollSheets();
    return;
  }

  setupHourlyRateMaster_(ss);
  setupPayrollAdjustmentSheet_(ss);
  setupPayrollMiscAdjustmentSheet_(ss);
  setupPayrollFinalizationSheet_(ss);
  refreshPayrollSummary_(ss);
  SpreadsheetApp.getUi().alert('給与集計を更新しました。');
}

function refreshPayrollSummaryIfExists_(spreadsheet) {
  try {
    const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getSheetByName(PAYROLL_SUMMARY_SHEET)) {
      refreshPayrollSummary_(ss, { skipCapture: true });
    }
  } catch (error) {
    console.error('給与集計更新エラー', error);
  }
}

function setupHourlyRateMaster_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HOURLY_RATE_MASTER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(HOURLY_RATE_MASTER_SHEET);
  }

  const headers = ['氏名', '時給', '適用開始日', '適用終了日', '有効', '備考'];
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  const isHeaderEmpty = currentHeaders.every((value) => !String(value || '').trim());
  if (isHeaderEmpty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const existingNames = {};
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, RATE_COL_NAME, sheet.getLastRow() - 1, 1)
      .getValues()
      .forEach((row) => {
        const name = String(row[0] || '').trim();
        if (name) existingNames[name] = true;
      });
  }

  const names = getPayrollUserOptions_(ss).filter((name) => name !== '全員');
  const appendRows = names
    .filter((name) => !existingNames[name])
    .map((name) => [
      name,
      DEFAULT_HOURLY_RATE,
      new Date(2000, 0, 1),
      '',
      true,
      '初期設定',
    ]);

  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, headers.length)
      .setValues(appendRows);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#6ea86b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  if (sheet.getMaxRows() >= 2) {
    sheet.getRange(2, RATE_COL_HOURLY_RATE, sheet.getMaxRows() - 1, 1)
      .setNumberFormat('¥#,##0');
    sheet.getRange(2, RATE_COL_START_DATE, sheet.getMaxRows() - 1, 2)
      .setNumberFormat('yyyy年m月d日');
  }
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 70);
  sheet.setColumnWidth(6, 180);
}

function setupPayrollAdjustmentSheet_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PAYROLL_ADJUSTMENT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PAYROLL_ADJUSTMENT_SHEET);
  }

  const headers = [
    '支給対象月', 'スタッフ', '依頼No.', '支給対象時間', '理由・備考',
    'インセンティブ時間（任意）', '調整支給額（任意）'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const validationRows = Math.max(1, sheet.getMaxRows() - 1);
  const userOptions = getPayrollUserOptions_(ss).filter((name) => name !== '全員');
  const requestOptions = getPayrollRequestOptions_(ss);

  if (userOptions.length) {
    const userRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(userOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, ADJUST_COL_NAME, validationRows, 1).setDataValidation(userRule);
  }

  if (requestOptions.length) {
    const requestRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(requestOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, ADJUST_COL_REQUEST_NO, validationRows, 1).setDataValidation(requestRule);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#6ea86b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(2, ADJUST_COL_PAYROLL_MONTH, validationRows, 1)
    .setNumberFormat('yyyy年m月分');
  sheet.getRange(2, ADJUST_COL_PAID_HOURS, validationRows, 1)
    .setNumberFormat('0.00" h"');
  sheet.getRange(2, ADJUST_COL_INCENTIVE_HOURS, validationRows, 1)
    .setNumberFormat('0.00" h"');
  sheet.getRange(2, ADJUST_COL_ADJUSTED_PAY, validationRows, 1)
    .setNumberFormat('¥#,##0');
  const adjustedPayRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, ADJUST_COL_ADJUSTED_PAY, validationRows, 1)
    .setDataValidation(adjustedPayRule);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 240);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 140);
}

function setupPayrollMiscAdjustmentSheet_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PAYROLL_MISC_ADJUSTMENT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PAYROLL_MISC_ADJUSTMENT_SHEET);
  }

  const headers = [
    '支給対象月', 'スタッフ', '加減算額（±）', '区分', '理由・備考', '登録日時'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const validationRows = Math.max(1, sheet.getMaxRows() - 1);
  const userOptions = getPayrollUserOptions_(ss).filter((name) => name !== '全員');
  if (userOptions.length) {
    const userRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(userOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, MISC_COL_NAME, validationRows, 1).setDataValidation(userRule);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#6ea86b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(2, MISC_COL_PAYROLL_MONTH, validationRows, 1)
    .setNumberFormat('yyyy年m月分');
  sheet.getRange(2, MISC_COL_AMOUNT, validationRows, 1)
    .setNumberFormat('¥#,##0;[Red]-¥#,##0');
  sheet.getRange(2, MISC_COL_REGISTERED_AT, validationRows, 1)
    .setNumberFormat('yyyy年m月d日 hh:mm');
  sheet.getRange('C1').setNote(
    '給与から差し引く場合はマイナスで入力します。例：交通費の天引き -1500\n' +
    '追加支給する場合はプラスで入力できます。複数件ある場合は1件につき1行追加してください。'
  );
  sheet.getRange('A1').setNote(
    '同じ支給対象月・スタッフで複数行を登録できます。給与集計ではすべて合算します。\n' +
    '履歴を残すため、既存行の上書きより新しい行の追加を推奨します。'
  );
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 260);
  sheet.setColumnWidth(6, 150);
}

function completePayrollMiscAdjustmentRowsForRange_(sheet, range) {
  if (!sheet || !range || range.getLastRow() < 2) return;

  const startRow = Math.max(2, range.getRow());
  const endRow = Math.max(startRow, range.getLastRow());
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber++) {
    const values = sheet.getRange(rowNumber, 1, 1, 6).getValues()[0];
    const payrollMonth = normalizePayrollMonth_(values[MISC_COL_PAYROLL_MONTH - 1]);
    const name = String(values[MISC_COL_NAME - 1] || '').trim();
    const amountRaw = values[MISC_COL_AMOUNT - 1];
    const hasAmount = amountRaw !== '' && amountRaw !== null && Number.isFinite(Number(amountRaw));
    const registeredAt = values[MISC_COL_REGISTERED_AT - 1];

    if (payrollMonth && name && hasAmount && !registeredAt) {
      sheet.getRange(rowNumber, MISC_COL_REGISTERED_AT).setValue(new Date());
    }
  }
}

function setupPayrollFinalizationSheet_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PAYROLL_FINALIZATION_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PAYROLL_FINALIZATION_SHEET);
  }

  const headers = [
    '支給対象月', 'スタッフ', '自動計算額', '最終決定支払額', '差額',
    '確定方法', '理由・備考', '確定日時', '確定者'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const validationRows = Math.max(1, sheet.getMaxRows() - 1);
  const userOptions = getPayrollUserOptions_(ss).filter((name) => name !== '全員');
  if (userOptions.length) {
    const userRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(userOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, FINAL_COL_NAME, validationRows, 1).setDataValidation(userRule);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#6ea86b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(2, FINAL_COL_PAYROLL_MONTH, validationRows, 1)
    .setNumberFormat('yyyy年m月分');
  sheet.getRange(2, FINAL_COL_CALCULATED_PAY, validationRows, 3)
    .setNumberFormat('¥#,##0');
  sheet.getRange(2, FINAL_COL_CONFIRMED_AT, validationRows, 1)
    .setNumberFormat('yyyy年m月d日 hh:mm');
  sheet.getRange('A1').setNote(
    '最終決定額を変更・再確定する場合は、既存行を上書きせず新しい行を追加してください。\n' +
    '同じ支給対象月・スタッフでは、一番下の完成行を最新の確定値として使用します。'
  );
  sheet.getRange('D1').setNote(
    '自動計算額のまま確定する場合も、その金額を入力してください。0円の確定も可能です。'
  );
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 260);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 190);
}

function openPayrollFinalizationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupPayrollFinalizationSheet_(ss);
  ensureDirectInputEditTrigger_(ss);
  const sheet = ss.getSheetByName(PAYROLL_FINALIZATION_SHEET);
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert(
    '給与確定履歴を開きました。\n\n' +
    '支給対象月・スタッフ・最終決定支払額・理由を新しい行へ入力してください。\n' +
    '自動計算額、差額、確定日時、確定者は自動補完されます。'
  );
}


/**
 * メニューから、現在の給与集計を給与確定履歴へ保存します。
 * 表示対象が「全員」の場合は全スタッフ、個人名の場合はそのスタッフのみ保存します。
 */
function confirmPayrollSummaryToHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PAYROLL_SUMMARY_SHEET);
  if (!sheet) {
    throw new Error('給与集計シートが見つかりません。');
  }

  const result = confirmPayrollSummaryToHistory_(ss, sheet);
  SpreadsheetApp.getUi().alert(
    `${result.payrollMonthLabel}の給与を確定しました。\n\n` +
    `給与確定履歴へ ${result.savedCount}名分を追加しました。`
  );
}

function confirmPayrollSummaryToHistory_(spreadsheet, summarySheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = summarySheet || ss.getSheetByName(PAYROLL_SUMMARY_SHEET);
  if (!sheet) {
    throw new Error('給与集計シートが見つかりません。');
  }

  // 内訳へ直接入力した調整支給額を先に支給時間調整へ保存します。
  capturePayrollDetailAdjustedPays_(ss, sheet);
  SpreadsheetApp.flush();

  setupPayrollAdjustmentSheet_(ss);
  setupPayrollMiscAdjustmentSheet_(ss);
  setupPayrollFinalizationSheet_(ss);

  const payrollMonth = normalizePayrollMonth_(sheet.getRange(PAYROLL_MONTH_CELL).getValue());
  if (!payrollMonth) {
    throw new Error('支給対象月を確認してください。');
  }

  const userOptions = getPayrollUserOptions_(ss);
  const currentUser = String(sheet.getRange(PAYROLL_USER_FILTER_CELL).getDisplayValue() || '').trim();
  const userFilter = userOptions.includes(currentUser) ? currentUser : '全員';
  const period = buildPayrollPeriod_(payrollMonth);
  const data = buildPayrollData_(ss, period, userFilter, { applyFinalization: false });

  if (!data.staffRows.length) {
    throw new Error('確定対象となる給与データがありません。');
  }

  const historySheet = ss.getSheetByName(PAYROLL_FINALIZATION_SHEET);
  const confirmedAt = new Date();
  const confirmedBy = getPayrollEditorName_();
  const rows = data.staffRows.map((staffRow) => {
    const name = String(staffRow[0] || '').trim();
    const calculatedPay = Math.round(Number(staffRow[5] || 0));
    const note = buildPayrollConfirmationNote_(data, name);
    return [
      new Date(`${payrollMonth}T00:00:00`),
      name,
      calculatedPay,
      calculatedPay,
      0,
      '給与集計から確定',
      note,
      confirmedAt,
      confirmedBy,
    ];
  });

  const startRow = historySheet.getLastRow() + 1;
  historySheet.getRange(startRow, 1, rows.length, 9).setValues(rows);
  historySheet.getRange(startRow, FINAL_COL_PAYROLL_MONTH, rows.length, 1)
    .setNumberFormat('yyyy年m月分');
  historySheet.getRange(startRow, FINAL_COL_CALCULATED_PAY, rows.length, 3)
    .setNumberFormat('¥#,##0');
  historySheet.getRange(startRow, FINAL_COL_CONFIRMED_AT, rows.length, 1)
    .setNumberFormat('yyyy年m月d日 hh:mm');

  SpreadsheetApp.flush();
  refreshPayrollSummary_(ss, { skipCapture: true });

  return {
    savedCount: rows.length,
    payrollMonth,
    payrollMonthLabel: formatPayrollMonthLabel_(payrollMonth),
  };
}

function buildPayrollConfirmationNote_(data, name) {
  const details = (data.detailRows || []).filter((row) => String(row[0] || '').trim() === name);
  const miscEntries = (data.miscAdjustmentEntries || []).filter(
    (item) => String(item.name || '').trim() === name
  );
  const adjustedDetails = details.filter((row) => {
    const incentiveHours = Number(row[4] || 0);
    const adjustedPayRaw = row[7];
    const hasAdjustedPay =
      adjustedPayRaw !== '' &&
      adjustedPayRaw !== null &&
      Number.isFinite(Number(adjustedPayRaw));
    return incentiveHours > 0 || hasAdjustedPay;
  });

  if (!adjustedDetails.length && !miscEntries.length) {
    return '給与集計から確定（案件別の時間・金額調整、給与加減算なし）';
  }

  const summaries = adjustedDetails.map((row) => {
    const requestNo = String(row[1] || '').trim();
    const actualHours = Number(row[2] || 0);
    const paidHours = Number(row[3] || 0);
    const incentiveHours = Number(row[4] || 0);
    const adjustedPayRaw = row[7];
    const hasAdjustedPay =
      adjustedPayRaw !== '' &&
      adjustedPayRaw !== null &&
      Number.isFinite(Number(adjustedPayRaw));

    const parts = [`${requestNo}: 実${formatPayrollHours_(actualHours)}h→支給${formatPayrollHours_(paidHours)}h`];
    if (incentiveHours > 0) {
      parts.push(`インセン${formatPayrollHours_(incentiveHours)}h`);
    }
    if (hasAdjustedPay) {
      parts.push(`調整支給額${Math.round(Number(adjustedPayRaw)).toLocaleString()}円`);
    }
    return parts.join(' / ');
  });

  const miscSummaries = miscEntries.map((item) => {
    const signAmount = `${item.amount >= 0 ? '+' : ''}${Number(item.amount).toLocaleString()}円`;
    const label = item.category || '加減算';
    const note = item.note ? ` (${item.note})` : '';
    return `${label} ${signAmount}${note}`;
  });

  return `給与集計から確定：${summaries.concat(miscSummaries).join(' ; ')}`;
}

function formatPayrollHours_(value) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return '0';
  return Number(number.toFixed(2)).toString();
}

function formatPayrollMonthLabel_(payrollMonth) {
  const normalized = normalizePayrollMonth_(payrollMonth);
  if (!normalized) return '';
  return `${Number(normalized.slice(0, 4))}年${Number(normalized.slice(5, 7))}月分`;
}

function processPayrollFinalizationRowsForRange_(spreadsheet, sheet, range) {
  const ss = spreadsheet || sheet.getParent();
  const startRow = Math.max(2, range.getRow());
  const endRow = Math.max(startRow, range.getLastRow());

  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber++) {
    const values = sheet.getRange(rowNumber, 1, 1, 9).getValues()[0];
    const payrollMonth = normalizePayrollMonth_(values[FINAL_COL_PAYROLL_MONTH - 1]);
    const name = String(values[FINAL_COL_NAME - 1] || '').trim();
    const finalRaw = values[FINAL_COL_FINAL_PAY - 1];
    const hasFinalValue = finalRaw !== '' && finalRaw !== null && Number.isFinite(Number(finalRaw));
    if (!payrollMonth || !name || !hasFinalValue || Number(finalRaw) < 0) continue;

    // 確定履歴は確定時点のスナップショットとして保持します。
    // 自動計算額がすでに入っている行は、後日の勤怠・時給変更で上書きしません。
    const storedCalculatedRaw = values[FINAL_COL_CALCULATED_PAY - 1];
    const hasStoredCalculated =
      storedCalculatedRaw !== '' &&
      storedCalculatedRaw !== null &&
      Number.isFinite(Number(storedCalculatedRaw));

    let calculatedPay;
    if (hasStoredCalculated) {
      calculatedPay = Math.round(Number(storedCalculatedRaw));
    } else {
      const period = buildPayrollPeriod_(payrollMonth);
      const baseData = buildPayrollData_(ss, period, name, { applyFinalization: false });
      calculatedPay = Math.round(Number(baseData.totalCalculatedPay || 0));
      sheet.getRange(rowNumber, FINAL_COL_CALCULATED_PAY).setValue(calculatedPay);
    }

    const finalPay = Math.round(Number(finalRaw));
    const difference = finalPay - calculatedPay;
    const existingMethod = String(values[FINAL_COL_METHOD - 1] || '').trim();
    const method = existingMethod === '給与集計から確定'
      ? existingMethod
      : difference === 0
        ? '自動計算額で確定'
        : '最終支給額を手入力';
    const confirmedAt = values[FINAL_COL_CONFIRMED_AT - 1] || new Date();
    const confirmedBy = String(values[FINAL_COL_CONFIRMED_BY - 1] || '').trim() || getPayrollEditorName_();

    sheet.getRange(rowNumber, FINAL_COL_DIFFERENCE).setValue(difference);
    sheet.getRange(rowNumber, FINAL_COL_METHOD).setValue(method);
    sheet.getRange(rowNumber, FINAL_COL_CONFIRMED_AT).setValue(confirmedAt);
    sheet.getRange(rowNumber, FINAL_COL_CONFIRMED_BY).setValue(confirmedBy);
  }
  SpreadsheetApp.flush();
}

function processPayrollFinalizationAllRows_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PAYROLL_FINALIZATION_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  processPayrollFinalizationRowsForRange_(
    ss,
    sheet,
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 9)
  );
}

function getPayrollEditorName_() {
  try {
    const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    return String(email || '').trim() || '手入力';
  } catch (error) {
    return '手入力';
  }
}

function refreshPayrollSummary_(spreadsheet, options) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('給与集計対象のスプレッドシートを取得できませんでした。');
  }

  const sheet = ss.getSheetByName(PAYROLL_SUMMARY_SHEET);
  if (!sheet) return;

  // メニューからの手動更新時は、案件別内訳へ直接入力された調整支給額を
  // 再描画前に「支給時間調整」へ保存します。
  // 対象月変更などの自動更新時は、旧表示を新しい月へ誤保存しないよう省略します。
  if (!(options && options.skipCapture)) {
    capturePayrollDetailAdjustedPays_(ss, sheet);
  }
  processPayrollFinalizationAllRows_(ss);

  const today = new Date();
  const currentMonthValue = sheet.getRange(PAYROLL_MONTH_CELL).getValue();
  const currentUserValue = String(
    sheet.getRange(PAYROLL_USER_FILTER_CELL).getDisplayValue() || ''
  ).trim();

  const payrollMonth = normalizePayrollMonth_(currentMonthValue) ||
    Utilities.formatDate(
      new Date(today.getFullYear(), today.getMonth(), 1),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  const userOptions = getPayrollUserOptions_(ss);
  const userFilter = userOptions.includes(currentUserValue) ? currentUserValue : '全員';
  const period = buildPayrollPeriod_(payrollMonth);
  const data = buildPayrollData_(ss, period, userFilter);

  sheet.clear();
  const requiredRows = Math.max(60, data.staffRows.length + data.detailRows.length + 20);
  ensureSheetSize_(sheet, requiredRows, 10);

  sheet.getRange('A1:J1')
    .setBackground('#6ea86b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16);
  sheet.getRange('A1').setValue('給与集計');

  sheet.getRange('A2').setValue('支給対象月');
  sheet.getRange(PAYROLL_MONTH_CELL)
    .setValue(new Date(`${payrollMonth}T00:00:00`))
    .setNumberFormat('yyyy年m月分');

  sheet.getRange('A3').setValue('対象期間（登録日時）');
  sheet.getRange('B3').setValue(period.label);

  sheet.getRange('D2').setValue('表示対象');
  sheet.getRange(PAYROLL_USER_FILTER_CELL).setValue(userFilter);
  const userRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(userOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(PAYROLL_USER_FILTER_CELL).setDataValidation(userRule);

  sheet.getRange(PAYROLL_CONFIRM_LABEL_CELL).setValue('履歴へ確定保存');
  sheet.getRange(PAYROLL_CONFIRM_CELL)
    .insertCheckboxes()
    .setValue(false)
    .setNote(
      'チェックすると、現在表示している支給対象月の給与を給与確定履歴へ保存します。\n' +
      '表示対象が「全員」なら全スタッフ、個人名ならその人だけを確定します。'
    );

  stylePayrollControlArea_(sheet);
  writePayrollSummaryCards_(sheet, data, 5);

  const staffHeaderRow = 9;
  writePayrollStaffTable_(sheet, staffHeaderRow, data);
  const staffEndRow = staffHeaderRow + 2 + Math.max(data.staffRows.length, 1);
  const detailStartRow = staffEndRow + 3;
  writePayrollDetailTable_(sheet, detailStartRow, data);

  sheet.setFrozenRows(3);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 140);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 140);
  sheet.setColumnWidth(10, 130);
  sheet.getRange(1, 1, sheet.getMaxRows(), 1).setHorizontalAlignment('left');
}

function buildPayrollPeriod_(payrollMonth) {
  const normalized = normalizePayrollMonth_(payrollMonth);
  if (!normalized) {
    throw new Error('支給対象月が不正です。');
  }

  const year = Number(normalized.slice(0, 4));
  const month = Number(normalized.slice(5, 7));
  const startDate = Utilities.formatDate(
    new Date(year, month - 2, 21),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
  const endDate = Utilities.formatDate(
    new Date(year, month - 1, 20),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );

  return {
    payrollMonth: normalized,
    startDate,
    endDate,
    label: `${formatJapaneseDate_(startDate)} ～ ${formatJapaneseDate_(endDate)}`,
  };
}

function buildPayrollData_(spreadsheet, period, userFilterValue, options) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const reportsSheet = ss.getSheetByName(REPORTS_SHEET);
  if (!reportsSheet) {
    throw new Error('DailyReportsシートが見つかりません。');
  }

  const applyFinalization = !(options && options.applyFinalization === false);
  const userFilter = String(userFilterValue || '全員').trim() || '全員';
  const rateRows = getHourlyRateRows_(ss);
  const adjustmentMap = getPayrollAdjustmentMap_(ss, period.payrollMonth);
  const miscAdjustmentMap = getPayrollMiscAdjustmentMap_(ss, period.payrollMonth);
  const finalizationMap = applyFinalization
    ? getPayrollFinalizationMap_(ss, period.payrollMonth)
    : {};
  const detailMap = {};

  const reportValues = reportsSheet.getDataRange().getValues();
  for (let i = 1; i < reportValues.length; i++) {
    const row = reportValues[i];
    // 給与集計では、A列「登録日時」が入力されている行だけを有効な日報として扱います。
    // スプレッドシート直接入力分も、A列を入力すればG・H列の補完状況に関係なく集計します。
    const submittedDate = normalizeWorkDate_(row[REPORT_COL_SUBMITTED_AT - 1]);
    const name = String(row[REPORT_COL_NAME - 1] || '').trim();
    const requestNo = String(row[REPORT_COL_REQUEST_NO - 1] || '').trim();
    const workContent = String(row[REPORT_COL_WORK_CONTENT - 1] || '').trim();
    const hours = Number(row[REPORT_COL_WORK_HOURS - 1] || 0);

    if (!submittedDate || !name || !requestNo || !workContent || !(hours > 0)) continue;
    if (!isDateInRange_(submittedDate, period.startDate, period.endDate)) continue;
    if (userFilter !== '全員' && name !== userFilter) continue;

    const key = `${name}|||${requestNo}`;
    if (!detailMap[key]) {
      detailMap[key] = {
        name,
        requestNo,
        actualHours: 0,
        actualPay: 0,
        rateSet: {},
      };
    }

    const rate = getHourlyRateForDate_(rateRows, name, submittedDate);
    detailMap[key].actualHours = Number((detailMap[key].actualHours + hours).toFixed(2));
    detailMap[key].actualPay = Math.round(detailMap[key].actualPay + hours * rate);
    detailMap[key].rateSet[String(rate)] = true;
  }

  Object.keys(adjustmentMap).forEach((key) => {
    const adjustment = adjustmentMap[key];
    if (userFilter !== '全員' && adjustment.name !== userFilter) return;
    if (!detailMap[key]) {
      detailMap[key] = {
        name: adjustment.name,
        requestNo: adjustment.requestNo,
        actualHours: 0,
        actualPay: 0,
        rateSet: {},
      };
    }
  });

  const staffMap = {};
  const detailItems = Object.keys(detailMap).map((key) => {
    const item = detailMap[key];
    const adjustment = adjustmentMap[key] || null;
    const requestedPaidHours = adjustment && adjustment.paidHours > 0
      ? Number(adjustment.paidHours)
      : adjustment
        ? Number((item.actualHours + Math.max(0, adjustment.incentiveHours || 0)).toFixed(2))
        : 0;
    const paidHours = adjustment
      ? Math.max(item.actualHours, requestedPaidHours)
      : item.actualHours;
    const incentiveHours = Number(Math.max(0, paidHours - item.actualHours).toFixed(2));
    const incentiveRate = getHourlyRateForDate_(rateRows, item.name, period.endDate);
    const incentivePay = Math.round(incentiveHours * incentiveRate);
    const automaticPay = Math.round(item.actualPay + incentivePay);
    const adjustedPay = adjustment && adjustment.hasAdjustedPay
      ? Math.round(Number(adjustment.adjustedPay))
      : null;
    const effectivePay = adjustedPay !== null ? adjustedPay : automaticPay;
    const rateValues = Object.keys(item.rateSet).map(Number).filter((value) => value > 0);
    if (incentiveHours > 0 && !rateValues.includes(incentiveRate)) {
      rateValues.push(incentiveRate);
    }
    const rateDisplay = rateValues.length <= 1
      ? (rateValues[0] || incentiveRate)
      : '複数';

    let note = adjustment ? adjustment.note : '';
    if (adjustment && adjustment.paidHours > 0 && requestedPaidHours < item.actualHours) {
      note = [note, '支給対象時間が実稼働未満のため、実稼働時間を使用'].filter(Boolean).join(' / ');
    }
    if (adjustment && adjustment.incentiveHours > 0 && adjustment.paidHours > 0) {
      note = [note, '支給対象時間を優先'].filter(Boolean).join(' / ');
    }
    if (adjustment && item.actualHours === 0) {
      note = [note, '実稼働なし・調整のみ'].filter(Boolean).join(' / ');
    }
    if (adjustedPay !== null) {
      note = [note, '調整支給額を優先'].filter(Boolean).join(' / ');
    }

    if (!staffMap[item.name]) {
      staffMap[item.name] = {
        name: item.name,
        actualHours: 0,
        incentiveHours: 0,
        paidHours: 0,
        baseCalculatedPay: 0,
        rateSet: {},
      };
    }
    const staff = staffMap[item.name];
    staff.actualHours = Number((staff.actualHours + item.actualHours).toFixed(2));
    staff.incentiveHours = Number((staff.incentiveHours + incentiveHours).toFixed(2));
    staff.paidHours = Number((staff.paidHours + paidHours).toFixed(2));
    staff.baseCalculatedPay += effectivePay;
    rateValues.forEach((rate) => {
      if (rate > 0) staff.rateSet[String(rate)] = true;
    });

    return {
      name: item.name,
      requestNo: item.requestNo,
      actualHours: item.actualHours,
      paidHours,
      incentiveHours,
      rateDisplay,
      automaticPay,
      adjustedPay,
      effectivePay,
      note,
    };
  });

  Object.keys(miscAdjustmentMap).forEach((name) => {
    if (userFilter !== '全員' && name !== userFilter) return;
    if (!staffMap[name]) {
      staffMap[name] = {
        name,
        actualHours: 0,
        incentiveHours: 0,
        paidHours: 0,
        baseCalculatedPay: 0,
        rateSet: {},
      };
    }
  });

  const staffItems = Object.values(staffMap)
    .sort((a, b) => a.name.localeCompare(b.name, 'ja'))
    .map((item) => {
      const miscAdjustment = miscAdjustmentMap[item.name]
        ? Number(miscAdjustmentMap[item.name].total || 0)
        : 0;
      const calculatedPay = Math.round(Number(item.baseCalculatedPay || 0) + miscAdjustment);
      const finalization = finalizationMap[item.name] || null;
      const finalPay = finalization ? finalization.finalPay : calculatedPay;
      return {
        ...item,
        miscAdjustment,
        calculatedPay,
        finalPay,
        difference: finalPay - calculatedPay,
        finalized: Boolean(finalization),
      };
    });

  const staffRows = staffItems.map((item) => {
    const rates = Object.keys(item.rateSet).map(Number).filter((value) => value > 0);
    return [
      item.name,
      item.actualHours,
      item.incentiveHours,
      item.paidHours,
      rates.length <= 1 ? (rates[0] || getHourlyRateForDate_(rateRows, item.name, period.endDate)) : '複数',
      item.calculatedPay,
      item.finalPay,
      item.difference,
      item.finalized ? '確定済' : '未確定',
      item.miscAdjustment,
    ];
  });

  const detailRows = detailItems
    .sort((a, b) =>
      a.name.localeCompare(b.name, 'ja') ||
      a.requestNo.localeCompare(b.requestNo, 'ja')
    )
    .map((item) => [
      item.name,
      item.requestNo,
      item.actualHours,
      item.paidHours,
      item.incentiveHours,
      item.rateDisplay,
      item.automaticPay,
      item.adjustedPay === null ? '' : item.adjustedPay,
      item.note,
    ]);

  const totalActualHours = Number(
    staffItems.reduce((sum, item) => sum + item.actualHours, 0).toFixed(2)
  );
  const totalIncentiveHours = Number(
    staffItems.reduce((sum, item) => sum + item.incentiveHours, 0).toFixed(2)
  );
  const totalPaidHours = Number(
    staffItems.reduce((sum, item) => sum + item.paidHours, 0).toFixed(2)
  );
  const totalMiscAdjustment = staffItems.reduce((sum, item) => sum + item.miscAdjustment, 0);
  const totalCalculatedPay = staffItems.reduce((sum, item) => sum + item.calculatedPay, 0);
  const totalFinalPay = staffItems.reduce((sum, item) => sum + item.finalPay, 0);
  const totalFinalDifference = totalFinalPay - totalCalculatedPay;

  return {
    userFilter,
    staffRows,
    detailRows,
    miscAdjustmentEntries: Object.values(miscAdjustmentMap).reduce((all, item) => all.concat(item.entries || []), []),
    totalActualHours,
    totalIncentiveHours,
    totalPaidHours,
    totalMiscAdjustment,
    totalCalculatedPay,
    totalFinalPay,
    totalFinalDifference,
    totalPay: totalFinalPay,
  };
}

function getHourlyRateRows_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(HOURLY_RATE_MASTER_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
    .map((row, index) => ({
      rowNumber: index + 2,
      name: String(row[RATE_COL_NAME - 1] || '').trim(),
      rate: Number(row[RATE_COL_HOURLY_RATE - 1] || 0),
      startDate: normalizeWorkDate_(row[RATE_COL_START_DATE - 1]),
      endDate: normalizeWorkDate_(row[RATE_COL_END_DATE - 1]),
      enabled: isEnabledValue_(row[RATE_COL_ENABLED - 1]),
      note: String(row[RATE_COL_NOTE - 1] || '').trim(),
    }))
    .filter((item) => item.name && item.rate > 0 && item.startDate && item.enabled)
    .sort((a, b) => a.startDate.localeCompare(b.startDate) || a.rowNumber - b.rowNumber);
}

function getHourlyRateForDate_(rateRows, name, targetDate) {
  const normalizedDate = normalizeWorkDate_(targetDate);
  const normalizedName = String(name || '').trim();
  if (!normalizedDate || !normalizedName) return DEFAULT_HOURLY_RATE;

  const matches = (rateRows || []).filter((item) =>
    item.name === normalizedName &&
    item.startDate <= normalizedDate &&
    (!item.endDate || normalizedDate <= item.endDate)
  );
  if (!matches.length) return DEFAULT_HOURLY_RATE;

  matches.sort((a, b) =>
    b.startDate.localeCompare(a.startDate) || b.rowNumber - a.rowNumber
  );
  return Number(matches[0].rate || DEFAULT_HOURLY_RATE);
}

function getPayrollAdjustmentMap_(spreadsheet, payrollMonth) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PAYROLL_ADJUSTMENT_SHEET);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;

  const targetMonth = normalizePayrollMonth_(payrollMonth);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  values.forEach((row, index) => {
    const rowMonth = normalizePayrollMonth_(row[ADJUST_COL_PAYROLL_MONTH - 1]);
    const name = String(row[ADJUST_COL_NAME - 1] || '').trim();
    const requestNo = String(row[ADJUST_COL_REQUEST_NO - 1] || '').trim();
    const paidHours = Number(row[ADJUST_COL_PAID_HOURS - 1] || 0);
    const note = String(row[ADJUST_COL_NOTE - 1] || '').trim();
    const incentiveHours = Number(row[ADJUST_COL_INCENTIVE_HOURS - 1] || 0);
    const adjustedPayRaw = row[ADJUST_COL_ADJUSTED_PAY - 1];
    const hasAdjustedPay =
      adjustedPayRaw !== '' &&
      adjustedPayRaw !== null &&
      Number.isFinite(Number(adjustedPayRaw)) &&
      Number(adjustedPayRaw) >= 0;
    if (rowMonth !== targetMonth || !name || !requestNo) return;

    const key = `${name}|||${requestNo}`;
    if (!(paidHours > 0) && !(incentiveHours > 0) && !hasAdjustedPay) {
      // 同じキーの下側の行が空欄なら、以前の調整値を解除したものとして扱います。
      delete map[key];
      return;
    }

    // 同じ対象月・スタッフ・依頼No.が複数ある場合は、下にある行を最終値として採用します。
    map[key] = {
      rowNumber: index + 2,
      name,
      requestNo,
      paidHours,
      incentiveHours,
      adjustedPay: hasAdjustedPay ? Math.round(Number(adjustedPayRaw)) : null,
      hasAdjustedPay,
      note,
    };
  });
  return map;
}

function getPayrollMiscAdjustmentMap_(spreadsheet, payrollMonth) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PAYROLL_MISC_ADJUSTMENT_SHEET);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;

  const targetMonth = normalizePayrollMonth_(payrollMonth);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  values.forEach((row, index) => {
    const rowMonth = normalizePayrollMonth_(row[MISC_COL_PAYROLL_MONTH - 1]);
    const name = String(row[MISC_COL_NAME - 1] || '').trim();
    const amountRaw = row[MISC_COL_AMOUNT - 1];
    const hasAmount = amountRaw !== '' && amountRaw !== null && Number.isFinite(Number(amountRaw));
    if (rowMonth !== targetMonth || !name || !hasAmount) return;

    const amount = Math.round(Number(amountRaw));
    const category = String(row[MISC_COL_CATEGORY - 1] || '').trim();
    const note = String(row[MISC_COL_NOTE - 1] || '').trim();
    const registeredAt = row[MISC_COL_REGISTERED_AT - 1];

    if (!map[name]) {
      map[name] = { name, total: 0, entries: [] };
    }
    map[name].total += amount;
    map[name].entries.push({
      rowNumber: index + 2,
      name,
      amount,
      category,
      note,
      registeredAt,
    });
  });

  Object.keys(map).forEach((name) => {
    map[name].total = Math.round(map[name].total);
  });
  return map;
}

function getPayrollFinalizationMap_(spreadsheet, payrollMonth) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PAYROLL_FINALIZATION_SHEET);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;

  const targetMonth = normalizePayrollMonth_(payrollMonth);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  values.forEach((row, index) => {
    const rowMonth = normalizePayrollMonth_(row[FINAL_COL_PAYROLL_MONTH - 1]);
    const name = String(row[FINAL_COL_NAME - 1] || '').trim();
    const finalRaw = row[FINAL_COL_FINAL_PAY - 1];
    const hasFinalValue = finalRaw !== '' && finalRaw !== null && Number.isFinite(Number(finalRaw));
    if (rowMonth !== targetMonth || !name || !hasFinalValue || Number(finalRaw) < 0) return;

    // 同じ対象月・スタッフが複数ある場合は、一番下の完成行を最新の確定値として採用します。
    map[name] = {
      rowNumber: index + 2,
      finalPay: Math.round(Number(finalRaw)),
      calculatedPay: Number(row[FINAL_COL_CALCULATED_PAY - 1] || 0),
      difference: Number(row[FINAL_COL_DIFFERENCE - 1] || 0),
      method: String(row[FINAL_COL_METHOD - 1] || '').trim(),
      note: String(row[FINAL_COL_NOTE - 1] || '').trim(),
      confirmedAt: row[FINAL_COL_CONFIRMED_AT - 1],
      confirmedBy: String(row[FINAL_COL_CONFIRMED_BY - 1] || '').trim(),
    };
  });
  return map;
}

function writePayrollSummaryCards_(sheet, data, startRow) {
  const row = Number(startRow || 5);
  sheet.getRange(row, 1, 1, 10).setValues([[
    '実稼働時間', '', '支給対象時間', '', '加減算額', '', '自動計算額', '', '最終決定支払額', ''
  ]]);
  sheet.getRange(row + 1, 1, 1, 10).setValues([[
    data.totalActualHours, '', data.totalPaidHours, '', data.totalMiscAdjustment, '', data.totalCalculatedPay, '', data.totalFinalPay, ''
  ]]);

  ['A', 'C', 'E', 'G', 'I'].forEach((column) => {
    sheet.getRange(`${column}${row}:${String.fromCharCode(column.charCodeAt(0) + 1)}${row}`)
      .setBackground('#edf7eb')
      .setFontColor('#6a846d')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#d7e8d4', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(`${column}${row + 1}:${String.fromCharCode(column.charCodeAt(0) + 1)}${row + 1}`)
      .setBackground('#ffffff')
      .setFontColor('#314c35')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#d7e8d4', SpreadsheetApp.BorderStyle.SOLID);
  });
  sheet.getRange(row + 1, 1, 1, 4).setNumberFormat('0.00" h"');
  sheet.getRange(row + 1, 5, 1, 6).setNumberFormat('¥#,##0;[Red]-¥#,##0');
}

function writePayrollStaffTable_(sheet, headerRow, data) {
  sheet.getRange(headerRow, 1, 1, 10)
    .setValues([['スタッフ別給与', '', '', '', '', '', '', '', '', '']])
    .setBackground('#6ea86b')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.getRange(headerRow + 1, 1, 1, 10)
    .setValues([[
      '氏名', '実稼働時間', 'インセンティブ時間', '支給対象時間', '適用時給',
      '自動計算額', '最終決定支払額', '差額', '確定状況', '加減算額'
    ]]);
  styleTableHeader_(sheet.getRange(headerRow + 1, 1, 1, 10));
  sheet.getRange(headerRow + 1, 10).setNote(
    '「給与加減算」シートの対象月・スタッフ別合計です。天引きはマイナス、追加支給はプラスで記録します。'
  );

  if (!data.staffRows.length) {
    sheet.getRange(headerRow + 2, 1, 1, 10)
      .setValues([['対象期間の記録はありません', '', '', '', '', '', '', '', '', '']]);
    return;
  }

  sheet.getRange(headerRow + 2, 1, data.staffRows.length, 10).setValues(data.staffRows);
  styleTableBody_(sheet.getRange(headerRow + 2, 1, data.staffRows.length, 10));
  sheet.getRange(headerRow + 2, 2, data.staffRows.length, 3).setNumberFormat('0.00" h"');
  sheet.getRange(headerRow + 2, 5, data.staffRows.length, 4).setNumberFormat('¥#,##0;[Red]-¥#,##0');
  sheet.getRange(headerRow + 2, 10, data.staffRows.length, 1).setNumberFormat('¥#,##0;[Red]-¥#,##0');
}

function writePayrollDetailTable_(sheet, startRow, data) {
  sheet.getRange(startRow, 1, 1, 9)
    .setValues([['案件別 支給内訳', '', '', '', '', '', '', '', '']])
    .setBackground('#8cc68a')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.getRange(startRow + 1, 1, 1, 9)
    .setValues([[
      '氏名', '依頼No.', '実稼働時間', '支給対象時間',
      'インセンティブ時間', '適用時給', '給与支給額', '調整支給額', '理由・備考'
    ]]);
  styleTableHeader_(sheet.getRange(startRow + 1, 1, 1, 9));
  sheet.getRange(startRow + 1, 8).setNote(
    '入力がある場合、この案件の給与計算では「給与支給額」より調整支給額を優先します。\n' +
    '入力内容は「支給時間調整」シートへ保存されます。'
  );

  if (!data.detailRows.length) {
    sheet.getRange(startRow + 2, 1, 1, 9)
      .setValues([['対象期間の記録はありません', '', '', '', '', '', '', '', '']]);
    return;
  }

  sheet.getRange(startRow + 2, 1, data.detailRows.length, 9).setValues(data.detailRows);
  styleTableBody_(sheet.getRange(startRow + 2, 1, data.detailRows.length, 9));
  sheet.getRange(startRow + 2, 3, data.detailRows.length, 3).setNumberFormat('0.00" h"');
  sheet.getRange(startRow + 2, 6, data.detailRows.length, 3).setNumberFormat('¥#,##0');

  const adjustedPayRange = sheet.getRange(startRow + 2, 8, data.detailRows.length, 1);
  adjustedPayRange
    .setBackground('#fff8dc')
    .setNumberFormat('¥#,##0');
  const adjustedPayRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(true)
    .build();
  adjustedPayRange.setDataValidation(adjustedPayRule);
}

function findPayrollDetailTable_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return null;

  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === '案件別 支給内訳') {
      return {
        titleRow: i + 1,
        headerRow: i + 2,
        firstDataRow: i + 3,
      };
    }
  }
  return null;
}

function savePayrollDetailAdjustedPayFromRange_(spreadsheet, sheet, range) {
  const table = findPayrollDetailTable_(sheet);
  if (!table || range.getLastRow() < table.firstDataRow) return false;

  const adjustedPayColumn = 8;
  if (
    range.getLastColumn() < adjustedPayColumn ||
    range.getColumn() > adjustedPayColumn
  ) {
    return false;
  }

  const payrollMonth = normalizePayrollMonth_(sheet.getRange(PAYROLL_MONTH_CELL).getValue());
  if (!payrollMonth) return false;

  let changed = false;
  const startRow = Math.max(table.firstDataRow, range.getRow());
  const endRow = range.getLastRow();
  for (let row = startRow; row <= endRow; row++) {
    const name = String(sheet.getRange(row, 1).getDisplayValue() || '').trim();
    const requestNo = String(sheet.getRange(row, 2).getDisplayValue() || '').trim();
    if (!name || !requestNo) continue;

    const raw = sheet.getRange(row, adjustedPayColumn).getValue();
    upsertPayrollDetailAdjustedPay_(spreadsheet, payrollMonth, name, requestNo, raw);
    changed = true;
  }
  return changed;
}

function capturePayrollDetailAdjustedPays_(spreadsheet, sheet) {
  const table = findPayrollDetailTable_(sheet);
  if (!table || sheet.getLastRow() < table.firstDataRow) return;

  const payrollMonth = normalizePayrollMonth_(sheet.getRange(PAYROLL_MONTH_CELL).getValue());
  if (!payrollMonth) return;

  // 支給対象月を変更した直後など、上部の月と現在表示中の内訳が一致しない場合は保存しません。
  const displayedPeriodLabel = String(sheet.getRange('B3').getDisplayValue() || '').trim();
  const expectedPeriodLabel = buildPayrollPeriod_(payrollMonth).label;
  if (displayedPeriodLabel && displayedPeriodLabel !== expectedPeriodLabel) return;

  const rowCount = sheet.getLastRow() - table.firstDataRow + 1;
  const values = sheet.getRange(table.firstDataRow, 1, rowCount, 8).getValues();
  values.forEach((row) => {
    const name = String(row[0] || '').trim();
    const requestNo = String(row[1] || '').trim();
    if (!name || !requestNo) return;
    upsertPayrollDetailAdjustedPay_(spreadsheet, payrollMonth, name, requestNo, row[7]);
  });
}

function upsertPayrollDetailAdjustedPay_(spreadsheet, payrollMonth, name, requestNo, rawValue) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PAYROLL_ADJUSTMENT_SHEET);
  if (!sheet) {
    setupPayrollAdjustmentSheet_(ss);
    sheet = ss.getSheetByName(PAYROLL_ADJUSTMENT_SHEET);
  }
  if (!sheet) return;

  const targetMonth = normalizePayrollMonth_(payrollMonth);
  const normalizedName = String(name || '').trim();
  const normalizedRequestNo = String(requestNo || '').trim();
  if (!targetMonth || !normalizedName || !normalizedRequestNo) return;

  const hasValue =
    rawValue !== '' &&
    rawValue !== null &&
    Number.isFinite(Number(rawValue)) &&
    Number(rawValue) >= 0;
  const adjustedPay = hasValue ? Math.round(Number(rawValue)) : '';

  let targetRow = 0;
  if (sheet.getLastRow() >= 2) {
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    values.forEach((row, index) => {
      const rowMonth = normalizePayrollMonth_(row[0]);
      const rowName = String(row[1] || '').trim();
      const rowRequestNo = String(row[2] || '').trim();
      if (
        rowMonth === targetMonth &&
        rowName === normalizedName &&
        rowRequestNo === normalizedRequestNo
      ) {
        targetRow = index + 2;
      }
    });
  }

  if (!targetRow) {
    if (!hasValue) return;
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1, 1, 7).setValues([[
      new Date(`${targetMonth}T00:00:00`),
      normalizedName,
      normalizedRequestNo,
      '',
      '',
      '',
      adjustedPay,
    ]]);
  } else {
    sheet.getRange(targetRow, ADJUST_COL_ADJUSTED_PAY).setValue(adjustedPay);
  }
}

function stylePayrollControlArea_(sheet) {
  sheet.getRange('A2:A3')
    .setBackground('#edf7eb')
    .setFontWeight('bold')
    .setFontColor('#314c35');
  sheet.getRange('B2:B3')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, true, true, '#d7e8d4', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('D2')
    .setBackground('#edf7eb')
    .setFontWeight('bold')
    .setFontColor('#314c35');
  sheet.getRange('E2')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, true, true, '#d7e8d4', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(PAYROLL_CONFIRM_LABEL_CELL)
    .setBackground('#6ea86b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#6ea86b', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(PAYROLL_CONFIRM_CELL)
    .setBackground('#edf7eb')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#6ea86b', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('B3').setFontWeight('bold');
}

function getPayrollUserOptions_(spreadsheet) {
  return getAnalyticsUserOptions_(spreadsheet);
}

function getPayrollRequestOptions_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const requests = {};

  const masterSheet = ss.getSheetByName(REQUEST_MASTER_SHEET);
  if (masterSheet && masterSheet.getLastRow() >= 2) {
    masterSheet.getRange(2, REQUEST_COL_NO, masterSheet.getLastRow() - 1, 1)
      .getValues()
      .forEach((row) => {
        const requestNo = String(row[0] || '').trim();
        if (requestNo) requests[requestNo] = true;
      });
  }

  const reportsSheet = ss.getSheetByName(REPORTS_SHEET);
  if (reportsSheet && reportsSheet.getLastRow() >= 2) {
    reportsSheet.getRange(2, REPORT_COL_REQUEST_NO, reportsSheet.getLastRow() - 1, 1)
      .getValues()
      .forEach((row) => {
        const requestNo = String(row[0] || '').trim();
        if (requestNo) requests[requestNo] = true;
      });
  }

  return Object.keys(requests).sort((a, b) => a.localeCompare(b, 'ja'));
}

function normalizePayrollMonth_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      new Date(value.getFullYear(), value.getMonth(), 1),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  const text = String(value || '').trim();
  if (!text) return '';
  const match = text.match(/^(\d{4})[\/\-年](\d{1,2})(?:月|[\/\-]\d{1,2})?/);
  if (!match) return '';
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (!(year >= 1900) || month < 1 || month > 12) return '';
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-01`;
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
