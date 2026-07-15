const state = {
  sessionToken: '',
  loginUserName: '',
  registeredTotal: 0,
  requestOptions: [],
  requestMasterItems: [],
  currentTab: 'report',
  historyMode: 'day',
  selectedRequestNo: '',
};

const STORAGE_KEY =
  (window.APP_CONFIG && window.APP_CONFIG.STORAGE_KEY) || 'daily-report-app-session';

document.addEventListener('DOMContentLoaded', init);

function init() {
  const today = getTodayString();

  el('workDate').value = today;
  el('historyBaseDate').value = today;

  clearSession();
  showLogin();
  setLoginMasterLoading(true);

  bindEvents();
  loadMasterData();

  if (!document.querySelector('.report-row')) {
    addRow();
  }
}

function bindEvents() {
  el('loginBtn').addEventListener('click', login);
  el('logoutBtn').addEventListener('click', logout);
  el('workDate').addEventListener('change', handleDateChange);
  el('addRowBtn').addEventListener('click', () => addRow());
  el('submitBtn').addEventListener('click', submitReport);

  el('tabReportBtn').addEventListener('click', () => switchTab('report'));
  el('tabHistoryBtn').addEventListener('click', () => switchTab('history'));
  el('tabRequestBtn').addEventListener('click', () => switchTab('request'));

  el('historyModeDayBtn').addEventListener('click', () => switchHistoryMode('day'));
  el('historyModeWeekBtn').addEventListener('click', () => switchHistoryMode('week'));
  el('historyLoadBtn').addEventListener('click', loadHistory);

  el('addRequestBtn').addEventListener('click', addRequestNoFromApp);

  el('loginName').addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      login();
    }
  });
}

function el(id) {
  return document.getElementById(id);
}

function getTodayString() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function setMessage(id, text, isError = true) {
  const target = el(id);
  if (!target) return;
  target.textContent = text || '';
  target.style.color = isError ? '#d96c6c' : '#4b7f50';
}

function setButtonLoading(id, loading, loadingText) {
  const btn = el(id);
  if (!btn) return;

  if (!btn.dataset.defaultText) {
    btn.dataset.defaultText = btn.textContent;
  }

  btn.disabled = loading;
  btn.textContent = loading ? loadingText : btn.dataset.defaultText;
}

function setLoginMasterLoading(loading) {
  const select = el('loginName');
  const btn = el('loginBtn');

  if (!select || !btn) return;

  if (loading) {
    select.innerHTML = '<option value="">読み込み中...</option>';
    select.disabled = true;
    btn.disabled = true;
    return;
  }

  select.disabled = false;
  btn.disabled = false;
}

function setLoginUnavailable(message) {
  const select = el('loginName');
  const btn = el('loginBtn');

  if (select) {
    select.innerHTML = '<option value="">取得できませんでした</option>';
    select.disabled = true;
  }

  if (btn) {
    btn.disabled = true;
  }

  setMessage('loginMessage', message || '氏名リストを取得できませんでした。');
}

function clearSession() {
  state.sessionToken = '';
  state.loginUserName = '';

  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch (e) {
    console.error(e);
  }
}

function showApp() {
  el('loginCard').classList.add('hidden');
  el('appCard').classList.remove('hidden');
}

function showLogin() {
  el('appCard').classList.add('hidden');
  el('loginCard').classList.remove('hidden');
}

function logout(message = '') {
  clearSession();
  state.registeredTotal = 0;
  state.currentTab = 'report';

  el('displayName').textContent = '-';
  el('totalHours').textContent = '0 h';
  el('loginName').value = '';

  showLogin();
  switchTab('report');

  setMessage('loginMessage', message, true);
}

async function loadMasterData() {
  setMessage('loginMessage', '');
  setLoginMasterLoading(true);

  try {
    const res = await apiGet('master');
    if (!res.ok) {
      setLoginUnavailable(res.message || 'マスタデータ取得に失敗しました。');
      return;
    }

    const users = res.users || [];
    loadUserOptions(users);

    state.requestOptions = res.requestNos || [];
    syncRequestSelectOptions();

    if (users.length === 0) {
      setLoginUnavailable('有効な利用者が登録されていません。');
      return;
    }

    setLoginMasterLoading(false);

    if (!document.querySelector('.report-row')) {
      addRow();
    }
  } catch (e) {
    console.error(e);
    setLoginUnavailable(buildNetworkErrorMessage_(e));
  }
}

function loadUserOptions(users) {
  const select = el('loginName');
  select.innerHTML = '<option value="">選択してください</option>';

  users.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    select.appendChild(option);
  });
}

async function login() {
  const name = el('loginName').value.trim();

  setMessage('loginMessage', '');

  if (!name) {
    setMessage('loginMessage', '氏名を選択してください。');
    return;
  }

  setButtonLoading('loginBtn', true, 'ログイン中...');

  try {
    const res = await apiPost('login', { name });
    setButtonLoading('loginBtn', false, 'ログイン中...');

    if (!res.ok) {
      setMessage('loginMessage', res.message || 'ログインに失敗しました。');
      return;
    }

    state.sessionToken = res.token;
    state.loginUserName = res.name;

    el('displayName').textContent = state.loginUserName;
    showApp();

    switchTab('report');
    await refreshTotal();
  } catch (e) {
    console.error(e);
    setButtonLoading('loginBtn', false, 'ログイン中...');
    setMessage('loginMessage', buildNetworkErrorMessage_(e));
  }
}

function switchTab(tabName) {
  state.currentTab = tabName;

  el('reportTab').classList.toggle('hidden', tabName !== 'report');
  el('historyTab').classList.toggle('hidden', tabName !== 'history');
  el('requestTab').classList.toggle('hidden', tabName !== 'request');

  el('tabReportBtn').classList.toggle('active', tabName === 'report');
  el('tabHistoryBtn').classList.toggle('active', tabName === 'history');
  el('tabRequestBtn').classList.toggle('active', tabName === 'request');

  if (tabName === 'request') {
    loadRequestMasterList();
  } else if (tabName === 'history') {
    loadHistory();
  }
}

function switchHistoryMode(mode) {
  state.historyMode = mode;
  el('historyModeDayBtn').classList.toggle('active', mode === 'day');
  el('historyModeWeekBtn').classList.toggle('active', mode === 'week');
  loadHistory();
}

function buildRequestOptionsHtml(selectedValue = '') {
  let html = '<option value="">選択してください</option>';

  state.requestOptions.forEach((item) => {
    const selected = item === selectedValue ? 'selected' : '';
    html += `<option value="${escapeHtml(item)}" ${selected}>${escapeHtml(item)}</option>`;
  });

  return html;
}

function addRow(data = {}) {
  const rowsArea = el('rowsArea');
  const row = document.createElement('div');
  row.className = 'report-row';

  row.innerHTML = `
    <div class="report-row-toolbar">
      <div class="report-row-index">項目</div>
      <div class="row-order-controls" aria-label="項目の並べ替え">
        <button type="button" class="order-btn move-up-btn" aria-label="この項目を上へ移動">↑ 上へ</button>
        <button type="button" class="order-btn move-down-btn" aria-label="この項目を下へ移動">↓ 下へ</button>
      </div>
    </div>

    <div>
      <label class="mini-label">依頼No.</label>
      <select class="requestNo">
        ${buildRequestOptionsHtml(data.requestNo || '')}
      </select>
    </div>

    <div>
      <label class="mini-label">作業内容</label>
      <input
        type="text"
        class="workContent"
        placeholder="作業内容"
        value="${escapeHtml(data.workContent || '')}"
      >
    </div>

    <div>
      <label class="mini-label">作業時間</label>
      <div class="hour-wrap">
        <input
          type="number"
          class="workHours"
          min="0"
          step="0.5"
          inputmode="decimal"
          placeholder="0"
          value="${escapeHtml(data.workHours || '')}"
        >
        <span class="hour-unit">h</span>
      </div>
    </div>

    <div>
      <button type="button" class="delete-btn">削除</button>
    </div>
  `;

  rowsArea.appendChild(row);

  row.querySelector('.workHours').addEventListener('input', updateTotalPreview);
  row.querySelector('.delete-btn').addEventListener('click', () => removeRow(row));
  row.querySelector('.move-up-btn').addEventListener('click', () => moveRow(row, -1));
  row.querySelector('.move-down-btn').addEventListener('click', () => moveRow(row, 1));

  refreshRowOrderButtons();
}

function moveRow(row, direction) {
  const rowsArea = el('rowsArea');
  if (!rowsArea || !row) return;

  if (direction < 0) {
    const previous = row.previousElementSibling;
    if (previous) {
      rowsArea.insertBefore(row, previous);
    }
  } else {
    const next = row.nextElementSibling;
    if (next) {
      rowsArea.insertBefore(next, row);
    }
  }

  refreshRowOrderButtons();
}

function refreshRowOrderButtons() {
  const rows = [...document.querySelectorAll('.report-row')];

  rows.forEach((row, index) => {
    const indexLabel = row.querySelector('.report-row-index');
    const upBtn = row.querySelector('.move-up-btn');
    const downBtn = row.querySelector('.move-down-btn');

    if (indexLabel) {
      indexLabel.textContent = `項目 ${index + 1}`;
    }
    if (upBtn) {
      upBtn.disabled = index === 0;
    }
    if (downBtn) {
      downBtn.disabled = index === rows.length - 1;
    }
  });
}

function removeRow(row) {
  const rows = document.querySelectorAll('.report-row');

  if (rows.length <= 1) {
    clearRow(row);
    updateTotalPreview();
    refreshRowOrderButtons();
    return;
  }

  row.remove();
  updateTotalPreview();
  refreshRowOrderButtons();
}

function clearRow(row) {
  row.querySelector('.requestNo').value = '';
  row.querySelector('.workContent').value = '';
  row.querySelector('.workHours').value = '';
}

function collectRows() {
  const rows = [...document.querySelectorAll('.report-row')];
  return rows.map((row) => ({
    requestNo: row.querySelector('.requestNo').value.trim(),
    workContent: row.querySelector('.workContent').value.trim(),
    workHours: row.querySelector('.workHours').value.trim(),
  }));
}

function calcInputTotal() {
  const rows = collectRows();
  let sum = 0;

  rows.forEach((row) => {
    const h = Number(row.workHours || 0);
    if (!isNaN(h)) sum += h;
  });

  return Number(sum.toFixed(2));
}

function updateTotalPreview() {
  const inputTotal = calcInputTotal();
  const total = Number((state.registeredTotal + inputTotal).toFixed(2));
  el('totalHours').textContent = `${total} h`;
}

async function handleDateChange() {
  await refreshTotal();
}

async function refreshTotal() {
  if (!state.sessionToken) return;

  const workDate = el('workDate').value;
  setMessage('appMessage', '');

  try {
    const res = await apiGet('total', {
      token: state.sessionToken,
      workDate,
    });

    if (!res.ok) {
      handleSessionErrorIfNeeded_(res, 'appMessage');
      setMessage('appMessage', res.message || '合計取得に失敗しました。');
      return;
    }

    state.registeredTotal = Number(res.total || 0);
    updateTotalPreview();
  } catch (e) {
    console.error(e);
    setMessage('appMessage', buildNetworkErrorMessage_(e));
  }
}

async function submitReport() {
  const workDate = el('workDate').value;
  const rows = collectRows();
  const diary = el('diary').value.trim();

  setMessage('appMessage', '');
  setButtonLoading('submitBtn', true, '送信中...');

  try {
    const res = await apiPost('saveReport', {
      token: state.sessionToken,
      workDate,
      rows,
      diary,
    });

    setButtonLoading('submitBtn', false, '送信中...');

    if (!res.ok) {
      handleSessionErrorIfNeeded_(res, 'appMessage');
      setMessage('appMessage', res.message || '送信に失敗しました。');
      return;
    }

    setMessage('appMessage', res.message || '送信しました。', false);
    state.registeredTotal = Number(res.total || 0);

    el('rowsArea').innerHTML = '';
    addRow();
    el('diary').value = '';
    updateTotalPreview();

    if (state.currentTab === 'history') {
      await loadHistory();
    }

    window.scrollTo({ top: 0, behavior: 'smooth' });
  } catch (e) {
    console.error(e);
    setButtonLoading('submitBtn', false, '送信中...');
    setMessage('appMessage', buildNetworkErrorMessage_(e));
  }
}

async function loadHistory() {
  if (!state.sessionToken) return;

  const baseDate = el('historyBaseDate').value;
  if (!baseDate) {
    setMessage('historyMessage', '基準日を選択してください。');
    return;
  }

  setMessage('historyMessage', '');
  setButtonLoading('historyLoadBtn', true, '読み込み中...');

  try {
    if (state.historyMode === 'day') {
      const res = await apiGet('historyDay', {
        token: state.sessionToken,
        workDate: baseDate,
      });

      setButtonLoading('historyLoadBtn', false, '読み込み中...');

      if (!res.ok) {
        handleSessionErrorIfNeeded_(res, 'historyMessage');
        setMessage('historyMessage', res.message || '1日分取得に失敗しました。');
        return;
      }

      renderHistoryDay(res);
      return;
    }

    const res = await apiGet('historyWeek', {
      token: state.sessionToken,
      baseDate,
    });

    setButtonLoading('historyLoadBtn', false, '読み込み中...');

    if (!res.ok) {
      handleSessionErrorIfNeeded_(res, 'historyMessage');
      setMessage('historyMessage', res.message || '1週間分取得に失敗しました。');
      return;
    }

    renderHistoryWeek(res);
  } catch (e) {
    console.error(e);
    setButtonLoading('historyLoadBtn', false, '読み込み中...');
    setMessage('historyMessage', buildNetworkErrorMessage_(e));
  }
}

function renderHistoryDay(data) {
  const summary = el('historySummary');
  const list = el('historyList');

  summary.innerHTML = `
    <div class="history-summary-grid">
      <div class="history-summary-box">
        <div class="history-summary-label">日付</div>
        <div class="history-summary-value">${escapeHtml(data.workDate || '-')}</div>
      </div>
      <div class="history-summary-box">
        <div class="history-summary-label">1日合計</div>
        <div class="history-summary-value">${escapeHtml(String(data.dailyTotal || 0))} h</div>
      </div>
    </div>
  `;

  if (!data.entries || data.entries.length === 0) {
    list.innerHTML = '<div class="history-empty">この日の記録はありません。</div>';
    return;
  }

  let html = '';
  data.entries.forEach((entry) => {
    html += `
      <div class="history-card">
        <div class="history-card-header">
          <div class="history-title">送信: ${escapeHtml(formatDateTime(entry.submittedAt))}</div>
          <div class="history-meta">
            入力合計 ${escapeHtml(String(entry.inputTotal || 0))} h /
            累計 ${escapeHtml(String(entry.cumulativeTotal || 0))} h
          </div>
        </div>

        <div class="history-diary">${escapeHtml(entry.diary || '記載なし')}</div>

        <div class="history-row-list">
          ${(entry.rows || [])
            .map(
              (row) => `
                <div class="history-row-item">
                  <div class="history-row-line"><strong>依頼No.</strong> ${escapeHtml(row.requestNo || '')}</div>
                  <div class="history-row-line"><strong>作業内容</strong> ${escapeHtml(row.workContent || '')}</div>
                  <div class="history-row-line"><strong>作業時間</strong> ${escapeHtml(String(row.workHours || 0))} h</div>
                </div>
              `
            )
            .join('')}
        </div>
      </div>
    `;
  });

  list.innerHTML = html;
}

function renderHistoryWeek(data) {
  const summary = el('historySummary');
  const list = el('historyList');

  summary.innerHTML = `
    <div class="history-summary-grid">
      <div class="history-summary-box">
        <div class="history-summary-label">期間</div>
        <div class="history-summary-value">${escapeHtml(data.startDate)} 〜 ${escapeHtml(data.endDate)}</div>
      </div>
      <div class="history-summary-box">
        <div class="history-summary-label">対象</div>
        <div class="history-summary-value">${escapeHtml(data.name || '-')}</div>
      </div>
    </div>
  `;

  if (!data.days || data.days.length === 0) {
    list.innerHTML = '<div class="history-empty">この期間の記録はありません。</div>';
    return;
  }

  let html = '';
  data.days.forEach((day) => {
    const historyDate = day.workDateValue || parseHistoryDateValue_(day.workDate);

    html += `
      <div
        class="history-card history-week-card"
        data-history-date="${escapeHtml(historyDate)}"
        role="button"
        tabindex="0"
        aria-label="${escapeHtml(day.workDate || '')}の1日分を表示"
      >
        <div class="history-card-header">
          <div class="history-title">${escapeHtml(day.workDate || '-')}</div>
          <div class="history-meta">
            合計 ${escapeHtml(String(day.totalHours || 0))} h /
            送信回数 ${escapeHtml(String(day.reportCount || 0))} 回
          </div>
        </div>
        <div class="history-diary">${escapeHtml(day.latestDiary || '記載なし')}</div>
        <div class="history-card-action-note">タップしてこの日の一覧を表示</div>
      </div>
    `;
  });

  list.innerHTML = html;

  list.querySelectorAll('[data-history-date]').forEach((card) => {
    const openDay = () => openHistoryDayFromWeek_(card.dataset.historyDate);

    card.addEventListener('click', openDay);
    card.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        openDay();
      }
    });
  });
}

function openHistoryDayFromWeek_(workDate) {
  if (!workDate) return;

  el('historyBaseDate').value = workDate;
  switchHistoryMode('day');
}

function parseHistoryDateValue_(value) {
  const text = String(value || '').trim();
  const matched = text.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
  if (!matched) return '';

  return `${matched[1]}-${matched[2].padStart(2, '0')}-${matched[3].padStart(2, '0')}`;
}

async function loadRequestMasterList() {
  if (!state.sessionToken) return;

  setMessage('requestMessage', '');

  try {
    const res = await apiGet('requestList', {
      token: state.sessionToken,
    });

    if (!res.ok) {
      handleSessionErrorIfNeeded_(res, 'requestMessage');
      setMessage('requestMessage', res.message || '依頼No.一覧の取得に失敗しました。');
      return;
    }

    applyRequestMasterItems(res.items || []);
  } catch (e) {
    console.error(e);
    setMessage('requestMessage', buildNetworkErrorMessage_(e));
  }
}

function applyRequestMasterItems(items, preferredRequestNo = '') {
  state.requestMasterItems = items || [];
  state.requestOptions = state.requestMasterItems
    .filter((item) => item.enabled)
    .map((item) => item.requestNo);

  if (preferredRequestNo) {
    state.selectedRequestNo = preferredRequestNo;
  }

  syncRequestSelectOptions();
  renderRequestMasterList();
}

function syncRequestSelectOptions() {
  const selects = document.querySelectorAll('.requestNo');
  selects.forEach((select) => {
    const currentValue = select.value;
    select.innerHTML = buildRequestOptionsHtml(currentValue);

    if (currentValue && !state.requestOptions.includes(currentValue)) {
      select.value = '';
    }
  });
}

function renderRequestMasterList() {
  const list = el('requestList');

  if (!state.requestMasterItems.length) {
    state.selectedRequestNo = '';
    list.innerHTML = '<div class="empty-box">依頼No.はまだ登録されていません。</div>';
    return;
  }

  let selectedItem = state.requestMasterItems.find(
    (item) => item.requestNo === state.selectedRequestNo
  );

  if (!selectedItem) {
    selectedItem = state.requestMasterItems[0];
    state.selectedRequestNo = selectedItem.requestNo;
  }

  const optionsHtml = state.requestMasterItems
    .map((item) => {
      const selected = item.requestNo === selectedItem.requestNo ? 'selected' : '';
      const status = item.enabled ? '有効' : '無効';
      return `<option value="${escapeHtml(item.requestNo)}" ${selected}>${escapeHtml(item.requestNo)}（${status}）</option>`;
    })
    .join('');

  list.innerHTML = `
    <div class="request-manage-box">
      <div class="field">
        <label for="requestManageSelect">登録済み依頼No.</label>
        <select id="requestManageSelect">${optionsHtml}</select>
      </div>

      <div class="request-selected-panel">
        <div class="request-status-row">
          <span class="request-status-label">現在の状態</span>
          <span class="request-status-value ${selectedItem.enabled ? 'is-enabled' : 'is-disabled'}">
            ${selectedItem.enabled ? '有効' : '無効'}
          </span>
        </div>

        <div class="field request-edit-field">
          <label for="requestEditNo">依頼No.を編集</label>
          <input
            type="text"
            id="requestEditNo"
            value="${escapeHtml(selectedItem.requestNo)}"
          >
        </div>

        <div class="request-manage-actions">
          <button id="requestUpdateBtn" class="mini-btn" type="button">変更を保存</button>
          <button id="requestToggleBtn" class="mini-btn" type="button">
            ${selectedItem.enabled ? '無効化' : '有効化'}
          </button>
        </div>
      </div>
    </div>
  `;

  el('requestManageSelect').addEventListener('change', (event) => {
    state.selectedRequestNo = event.target.value;
    setMessage('requestMessage', '');
    renderRequestMasterList();
  });

  el('requestUpdateBtn').addEventListener('click', updateRequestNoFromApp);
  el('requestToggleBtn').addEventListener('click', () => {
    setRequestNoEnabledFromApp(selectedItem.requestNo, !selectedItem.enabled);
  });
}

async function addRequestNoFromApp() {
  if (!state.sessionToken) return;

  const input = el('newRequestNo');
  const requestNo = input.value.trim();

  setMessage('requestMessage', '');
  setButtonLoading('addRequestBtn', true, '追加中...');

  try {
    const res = await apiPost('requestAdd', {
      token: state.sessionToken,
      requestNo,
    });

    setButtonLoading('addRequestBtn', false, '追加中...');

    if (!res.ok) {
      handleSessionErrorIfNeeded_(res, 'requestMessage');
      setMessage('requestMessage', res.message || '依頼No.の追加に失敗しました。');
      return;
    }

    input.value = '';
    setMessage('requestMessage', res.message || '依頼No.を追加しました。', false);
    applyRequestMasterItems(res.items || [], requestNo);
  } catch (e) {
    console.error(e);
    setButtonLoading('addRequestBtn', false, '追加中...');
    setMessage('requestMessage', buildNetworkErrorMessage_(e));
  }
}

async function setRequestNoEnabledFromApp(requestNo, enabled) {
  if (!state.sessionToken || !requestNo) return;

  const actionText = enabled ? '有効化' : '無効化';
  const ok = window.confirm(`「${requestNo}」を${actionText}しますか？`);
  if (!ok) return;

  setMessage('requestMessage', '');
  setButtonLoading('requestToggleBtn', true, `${actionText}中...`);

  try {
    const res = await apiPost('requestSetEnabled', {
      token: state.sessionToken,
      requestNo,
      enabled,
    });

    if (!res.ok) {
      handleSessionErrorIfNeeded_(res, 'requestMessage');
      setMessage('requestMessage', res.message || `依頼No.の${actionText}に失敗しました。`);
      renderRequestMasterList();
      return;
    }

    setMessage('requestMessage', res.message || `依頼No.を${actionText}しました。`, false);
    applyRequestMasterItems(res.items || [], requestNo);
  } catch (e) {
    console.error(e);
    setMessage('requestMessage', buildNetworkErrorMessage_(e));
    renderRequestMasterList();
  }
}

async function updateRequestNoFromApp() {
  if (!state.sessionToken || !state.selectedRequestNo) return;

  const input = el('requestEditNo');
  const oldRequestNo = state.selectedRequestNo;
  const newRequestNo = input ? input.value.trim() : '';

  setMessage('requestMessage', '');

  if (!newRequestNo) {
    setMessage('requestMessage', '変更後の依頼No.を入力してください。');
    return;
  }

  if (newRequestNo === oldRequestNo) {
    setMessage('requestMessage', '依頼No.は変更されていません。');
    return;
  }

  setButtonLoading('requestUpdateBtn', true, '変更中...');

  try {
    const res = await apiPost('requestUpdate', {
      token: state.sessionToken,
      oldRequestNo,
      newRequestNo,
    });

    if (!res.ok) {
      handleSessionErrorIfNeeded_(res, 'requestMessage');
      setMessage('requestMessage', res.message || '依頼No.の変更に失敗しました。');
      renderRequestMasterList();
      return;
    }

    state.selectedRequestNo = newRequestNo;
    setMessage('requestMessage', res.message || '依頼No.を変更しました。', false);
    applyRequestMasterItems(res.items || [], newRequestNo);
  } catch (e) {
    console.error(e);
    setMessage('requestMessage', buildNetworkErrorMessage_(e));
    renderRequestMasterList();
  }
}

async function apiGet(action, params = {}) {
  const baseUrl = getApiBaseUrl_();
  const url = new URL(baseUrl);
  url.searchParams.set('action', action);

  Object.entries(params).forEach(([key, value]) => {
    if (value !== undefined && value !== null) {
      url.searchParams.set(key, String(value));
    }
  });

  const response = await fetch(url.toString(), {
    method: 'GET',
    cache: 'no-store',
  });

  return parseApiResponse_(response);
}

async function apiPost(action, payload = {}) {
  const baseUrl = getApiBaseUrl_();
  const url = new URL(baseUrl);
  url.searchParams.set('action', action);

  const response = await fetch(url.toString(), {
    method: 'POST',
    headers: {
      'Content-Type': 'text/plain;charset=utf-8',
    },
    body: JSON.stringify(payload),
  });

  return parseApiResponse_(response);
}

async function parseApiResponse_(response) {
  const text = await response.text();

  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error('APIのレスポンスを解析できませんでした。');
  }

  return json;
}

function getApiBaseUrl_() {
  const url = window.APP_CONFIG && window.APP_CONFIG.API_BASE_URL;
  if (!url || url.includes('ここに貼って')) {
    throw new Error('config.js に GAS の exec URL を設定してください。');
  }
  return url;
}

function handleSessionErrorIfNeeded_(res, messageId) {
  const message = String((res && res.message) || '');
  if (
    message.includes('再ログインしてください') ||
    message.includes('有効期限が切れました')
  ) {
    logout(message);
    if (messageId) {
      setMessage(messageId, '');
    }
  }
}

function buildNetworkErrorMessage_(error) {
  const message = String((error && error.message) || '');
  if (message) return message;
  return '通信時にエラーが発生しました。';
}

function formatDateTime(value) {
  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return String(value || '');
  }

  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  const hh = String(date.getHours()).padStart(2, '0');
  const mm = String(date.getMinutes()).padStart(2, '0');

  return `${y}/${m}/${d} ${hh}:${mm}`;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
