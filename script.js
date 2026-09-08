/* =====================================================
 * 勤怠管理（経理）5.0.0
 *
 * 構成：
 *   1. 設定（クラウド・モード）
 *   2. テーマ（ダークモード）
 *   3. 通知・ダイアログ
 *   4. 期間ユーティリティ（10日締め）
 *   5. データアクセス（localStorage）
 *   6. モード切替（通常／スイミング）
 *   7. クラウド連携（取込・元データ保存）
 *   8. 入力・編集・分割・削除
 *   9. 一覧表示
 *  10. Excelエクスポート
 *  11. バックアップ・リストア
 *  12. 初期化
 * ===================================================== */

/* ============ 1. 設定 ============ */

const CLOUD_URL = 'https://fdvibgxcuviftanwrjnn.supabase.co';
const CLOUD_KEY = 'sb_publishable_DA0blazZouSWkNMIylFKHg_aRPE6Cab';
const ADMIN_PASSWORD = '4564';

// 修正申請スプレッドシート（📋ボタンで画面右半分に開く。モードに応じて切り替わる）
const FIX_SHEET_URLS = {
  normal: 'https://docs.google.com/spreadsheets/d/1QEfzqy7FgpKEc6yq9XZemM5FUSNaPoEft2MmhMwRt8k/edit?usp=sharing',
  swim:   'https://docs.google.com/spreadsheets/d/1nWvfvl5mUJGPZE622L5EA1swR_zpDFyUiiA9UW1JptU/edit?usp=sharing'
};

// Excel・JSONの保存先Googleドライブフォルダ（📁ボタンで開く）
const DRIVE_FOLDER_URL = 'https://drive.google.com/drive/folders/1r0EapIF85w6xQnsPWxo7l7NCc9wsy9cN?usp=sharing';

// 通常／スイミングそれぞれの設定
const MODES = {
  normal: { label: '通常',       device: 'ipad-1',    storageKey: 'timeCards',      nameKey: 'nameUsageOrder' },
  swim:   { label: 'スイミング', device: 'ipad-swim', storageKey: 'timeCards_swim', nameKey: 'nameUsageOrder_swim' }
};

function getMode()  { return localStorage.getItem('tc_mode') || 'normal'; }
function modeCfg()  { return MODES[getMode()]; }

let editingRecord = null;   // 編集中レコード {name, date, index}

/* ============ 2. テーマ ============ */

function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  document.getElementById('themeToggle').textContent = theme === 'dark' ? '☀️' : '🌙';
  localStorage.setItem('theme', theme);
}

function toggleTheme() {
  const next = (localStorage.getItem('theme') || 'light') === 'dark' ? 'light' : 'dark';
  applyTheme(next);
}

/* ============ 3. 通知・ダイアログ ============ */

function showToast(message, type = 'info', duration = 3000) {
  const toast = document.createElement('div');
  toast.className = `toast-notification ${type}`;
  const icons = { success: '✓', error: '✕', info: 'i', warning: '!' };
  toast.innerHTML = `
    <div class="toast-content">
      <div class="toast-icon">${icons[type] || 'i'}</div>
      <div class="toast-message">${esc(message)}</div>
    </div>`;
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.style.animation = 'slideOutToast 0.3s ease-in forwards';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

function showConfirmDialog(title, message, confirmText = '実行', cancelText = 'キャンセル') {
  return new Promise(resolve => {
    const dialog = document.createElement('div');
    dialog.className = 'ios-dialog';
    dialog.innerHTML = `
      <div class="ios-dialog-content">
        <div class="ios-dialog-header">
          <div class="ios-dialog-title">${esc(title)}</div>
          <div class="ios-dialog-message">${esc(message)}</div>
        </div>
        <div class="ios-dialog-buttons">
          <button class="ios-dialog-button">${esc(cancelText)}</button>
          <button class="ios-dialog-button primary">${esc(confirmText)}</button>
        </div>
      </div>`;
    document.body.appendChild(dialog);
    const [cancelBtn, confirmBtn] = dialog.querySelectorAll('.ios-dialog-button');
    const close = result => { dialog.style.opacity = '0'; setTimeout(() => dialog.remove(), 250); resolve(result); };
    cancelBtn.addEventListener('click', () => close(false));
    confirmBtn.addEventListener('click', () => close(true));
    dialog.addEventListener('click', e => { if (e.target === dialog) close(false); });
  });
}

function showPromptDialog(title, message, defaultValue = '', inputType = 'text') {
  return new Promise(resolve => {
    const dialog = document.createElement('div');
    dialog.className = 'ios-dialog';
    dialog.innerHTML = `
      <div class="ios-dialog-content">
        <div class="ios-dialog-header">
          <div class="ios-dialog-title">${esc(title)}</div>
          <div class="ios-dialog-message">${esc(message)}</div>
          <input type="${inputType}" class="ios-dialog-input" value="${esc(defaultValue)}">
        </div>
        <div class="ios-dialog-buttons">
          <button class="ios-dialog-button">キャンセル</button>
          <button class="ios-dialog-button primary">OK</button>
        </div>
      </div>`;
    document.body.appendChild(dialog);
    const input = dialog.querySelector('input');
    const [cancelBtn, confirmBtn] = dialog.querySelectorAll('.ios-dialog-button');
    setTimeout(() => input.focus(), 100);
    const close = result => { dialog.style.opacity = '0'; setTimeout(() => dialog.remove(), 250); resolve(result); };
    cancelBtn.addEventListener('click', () => close(null));
    confirmBtn.addEventListener('click', () => close(input.value.trim()));
    input.addEventListener('keydown', e => { if (e.key === 'Enter') close(input.value.trim()); });
    dialog.addEventListener('click', e => { if (e.target === dialog) close(null); });
  });
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/* ============ 4. 期間ユーティリティ（10日締め：11日〜翌月10日） ============ */

// 指定日が属する期間の開始日・終了日
function getMonthPeriod(dateString) {
  const date = new Date(dateString);
  let y = date.getFullYear();
  let m = date.getMonth();
  if (date.getDate() <= 10) { m--; if (m < 0) { m = 11; y--; } }
  const start = new Date(y, m, 11);
  const end   = new Date(y, m + 1, 10);
  return { start, end };
}

// 期間キー "YYYY-MM"（開始月）を返す
function periodKeyOf(dateStr) {
  const y = +dateStr.slice(0, 4), m = +dateStr.slice(5, 7), d = +dateStr.slice(8, 10);
  if (d >= 11) return dateStr.slice(0, 7);
  const py = m === 1 ? y - 1 : y;
  const pm = m === 1 ? 12 : m - 1;
  return py + '-' + String(pm).padStart(2, '0');
}

function periodLabelOf(key) {
  const y = +key.slice(0, 4), m = +key.slice(5, 7);
  const nm = m === 12 ? 1 : m + 1;
  return `${y}年${m}月11日〜${nm}月10日`;
}

// 期間キー "YYYY-MM" → {start, end}
function periodOfKey(key) {
  const y = +key.slice(0, 4), m = +key.slice(5, 7);
  return { start: new Date(y, m - 1, 11), end: new Date(y, m, 10) };
}

// 「最新の月」＝データが存在する期間のうち、締め（10日）を過ぎた最新のもの。
// 例：今日が6/11でデータが5/10〜6/10までなら 5/11〜6/10（R8.5-6）。
// 有給の前倒し入力などで次の期間（6/11以降）のデータがあっても、締め前の期間は選ばない。
// まだ締めていない期間にしかデータが無い場合はその期間を返す。
function getLatestDataPeriod() {
  const data = loadData();
  const keys = new Set();
  for (const name in data) {
    for (const date in data[name]) {
      const d = new Date(date + 'T00:00:00');
      if (!isNaN(d.getTime())) keys.add(periodKeyOf(date));
    }
  }
  if (keys.size === 0) return getMonthPeriod(toDateStr(new Date()));

  const sorted = [...keys].sort();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  for (let i = sorted.length - 1; i >= 0; i--) {
    const p = periodOfKey(sorted[i]);
    if (p.end <= today) return p;
  }
  return periodOfKey(sorted[sorted.length - 1]);
}

function toDateStr(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function formatDate(dateString) {
  const parts = dateString.split('-');
  return parts.length === 3 ? `${parts[1]}-${parts[2]}` : dateString;
}

/* ============ 5. データアクセス ============ */

function loadDataFor(mode) {
  try {
    const obj = JSON.parse(localStorage.getItem(MODES[mode].storageKey) || '{}');
    return (obj && typeof obj === 'object' && !Array.isArray(obj)) ? obj : {};
  } catch { return {}; }
}

function loadData() { return loadDataFor(getMode()); }

function saveData(data) {
  localStorage.setItem(modeCfg().storageKey, JSON.stringify(data));
}

function updateNameUsageOrder(name) {
  const key = modeCfg().nameKey;
  const usage = JSON.parse(localStorage.getItem(key) || '{}');
  usage[name] = Date.now();
  localStorage.setItem(key, JSON.stringify(usage));
}

// 重複を除外しながら2つのデータを統合
function mergeTimeCards(existing, incoming) {
  const merged = JSON.parse(JSON.stringify(existing));
  let added = 0;
  for (const name in incoming) {
    if (!merged[name]) merged[name] = {};
    for (const date in incoming[name]) {
      if (!merged[name][date]) merged[name][date] = [];
      incoming[name][date].forEach(card => {
        const dup = merged[name][date].some(c =>
          c.checkIn === card.checkIn &&
          (c.checkOut || null) === (card.checkOut || null) &&
          !!c.isPaidLeave === !!card.isPaidLeave
        );
        if (!dup) { merged[name][date].push(card); added++; }
      });
    }
  }
  return { merged, added };
}

/* ============ 6. モード切替 ============ */

function setMode(mode) {
  localStorage.setItem('tc_mode', mode);
  document.querySelectorAll('#modeSwitch button').forEach(b =>
    b.classList.toggle('active', b.dataset.mode === mode));
  document.getElementById('listTitle').textContent = `勤怠一覧（${modeCfg().label}）`;
  cancelEdit();
  refreshNameList();
  displayTimeCards();
  updateRawSaveInfo();
  // クラウドログイン済みなら接続先デバイスを切り替えて再読込
  if (cloudToken && !document.getElementById('cloudMain').classList.contains('hidden')) {
    cloudLoadSnapshots();
  }
}

/* ============ 7. クラウド連携 ============ */

let cloudToken = sessionStorage.getItem('tc_cloud_token') || '';
let cloudSnapshots = [];
let cloudCurrent = null;   // 読込中のスナップショット {day, updated_at, record_count, data}

function cloudHeaders() {
  return { 'apikey': CLOUD_KEY, 'Authorization': 'Bearer ' + cloudToken, 'Content-Type': 'application/json' };
}

async function cloudLogin() {
  const email = document.getElementById('cloudEmail').value.trim();
  const password = document.getElementById('cloudPassword').value;
  const errBox = document.getElementById('cloudLoginError');
  errBox.innerHTML = '';
  try {
    const res = await fetch(CLOUD_URL + '/auth/v1/token?grant_type=password', {
      method: 'POST',
      headers: { 'apikey': CLOUD_KEY, 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password })
    });
    if (!res.ok) {
      let detail = 'ログインに失敗しました (HTTP ' + res.status + ')';
      try {
        const err = await res.json();
        const code = err.error_code || err.error || '';
        if (code === 'invalid_credentials' || code === 'invalid_grant') detail = 'メールアドレスまたはパスワードが違います。';
      } catch {}
      errBox.innerHTML = `<div class="error">${esc(detail)}</div>`;
      return;
    }
    const json = await res.json();
    cloudToken = json.access_token;
    sessionStorage.setItem('tc_cloud_token', cloudToken);
    showCloudMain();
  } catch {
    errBox.innerHTML = '<div class="error">通信エラーです。ネット接続を確認してください。</div>';
  }
}

function cloudLogout() {
  cloudToken = '';
  sessionStorage.removeItem('tc_cloud_token');
  cloudCurrent = null;
  document.getElementById('cloudMain').classList.add('hidden');
  document.getElementById('cloudLogin').classList.remove('hidden');
}

function showCloudMain() {
  document.getElementById('cloudLogin').classList.add('hidden');
  document.getElementById('cloudMain').classList.remove('hidden');
  cloudLoadSnapshots();
}

// 現在モードのデバイスのスナップショット一覧を取得
async function cloudLoadSnapshots() {
  const device = modeCfg().device;
  const res = await fetch(
    CLOUD_URL + '/rest/v1/tce_backups?select=id,day,record_count,updated_at' +
    '&device=eq.' + encodeURIComponent(device) + '&order=day.desc&limit=90',
    { headers: cloudHeaders() }
  );
  if (res.status === 401) { cloudLogout(); return; }
  if (!res.ok) { showToast('クラウドの読み込みに失敗しました (HTTP ' + res.status + ')', 'error'); return; }
  cloudSnapshots = await res.json();

  const sel = document.getElementById('cloudSnapshotSel');
  sel.innerHTML = cloudSnapshots.map(s =>
    `<option value="${s.id}">${s.day}（${s.record_count}件）</option>`).join('');

  if (cloudSnapshots.length === 0) {
    cloudCurrent = null;
    document.getElementById('cloudSnapMeta').textContent = 'まだバックアップが届いていません。';
    document.getElementById('cloudPeriodSel').innerHTML = '';
    document.getElementById('cloudSummary').innerHTML = '';
    return;
  }
  cloudLoadSnapshot(cloudSnapshots[0].id);
}

async function cloudLoadSnapshot(id) {
  const res = await fetch(
    CLOUD_URL + '/rest/v1/tce_backups?id=eq.' + id + '&select=day,updated_at,record_count,data',
    { headers: cloudHeaders() }
  );
  if (res.status === 401) { cloudLogout(); return; }
  const rows = await res.json();
  cloudCurrent = rows[0] || null;
  if (!cloudCurrent) return;

  const t = new Date(cloudCurrent.updated_at);
  document.getElementById('cloudSnapMeta').textContent =
    `最終送信: ${t.toLocaleString('ja-JP')} ／ 総記録数: ${cloudCurrent.record_count}件`;

  // 期間プルダウンを作成
  const data = cloudData();
  const periods = new Set();
  for (const name in data) {
    for (const d in data[name]) periods.add(periodKeyOf(d));
  }
  const sorted = [...periods].sort().reverse();
  document.getElementById('cloudPeriodSel').innerHTML =
    sorted.map(p => `<option value="${p}">${periodLabelOf(p)}</option>`).join('');

  renderCloudSummary();
}

// クラウドの生データ（旧combined形式にも対応）
function cloudData() {
  if (!cloudCurrent || !cloudCurrent.data) return {};
  const d = cloudCurrent.data;
  if (d.normal !== undefined || d.swim !== undefined) return d[getMode()] || {};
  return d;
}

// 選択期間だけに絞ったクラウドデータ
function cloudDataForPeriod() {
  const period = document.getElementById('cloudPeriodSel').value;
  const data = cloudData();
  const filtered = {};
  for (const name in data) {
    for (const d in data[name]) {
      if (periodKeyOf(d) !== period) continue;
      if (!filtered[name]) filtered[name] = {};
      filtered[name][d] = data[name][d];
    }
  }
  return filtered;
}

// ミニビューワー（名前・日数・合計時間）
function renderCloudSummary() {
  const data = cloudDataForPeriod();
  const names = Object.keys(data).sort((a, b) => a.localeCompare(b, 'ja'));
  const box = document.getElementById('cloudSummary');
  if (names.length === 0) {
    box.innerHTML = '<p class="hint">この期間のデータはありません。</p>';
    return;
  }
  let html = '<table><tr><th>名前</th><th>日数</th><th>合計</th></tr>';
  for (const name of names) {
    let days = 0, total = 0, incomplete = 0;
    for (const d in data[name]) {
      days++;
      for (const rec of data[name][d]) {
        if (rec.checkIn && rec.checkOut) total += calculateTimeDifference(rec.checkIn, rec.checkOut);
        else if (rec.checkIn) incomplete++;
      }
    }
    html += `<tr><td>${esc(name)}</td><td>${days}日</td><td>${total.toFixed(1)}h${incomplete ? ' ⚠' : ''}</td></tr>`;
  }
  html += '</table>';
  box.innerHTML = html;
}

// クラウド → ローカルへ取込（重複は除外）
async function cloudImport() {
  if (!cloudCurrent) { showToast('クラウドデータが読み込まれていません', 'warning'); return; }
  const incoming = cloudDataForPeriod();
  if (Object.keys(incoming).length === 0) { showToast('この期間のデータがありません', 'warning'); return; }

  const period = document.getElementById('cloudPeriodSel').value;
  const ok = await showConfirmDialog(
    'クラウドデータ取込',
    `【${modeCfg().label}】${periodLabelOf(period)} のデータをローカルに統合します。\n既にある記録と同じものは追加されません。`,
    '取り込む'
  );
  if (!ok) return;

  const { merged, added } = mergeTimeCards(loadData(), incoming);
  saveData(merged);
  refreshNameList();
  displayTimeCards();
  showToast(added > 0 ? `${added}件の記録を取り込みました` : '新しい記録はありませんでした（すべて取込済み）', 'success');
}

// 元データの月次ローカル保存
function cloudRawSave() {
  if (!cloudCurrent) { showToast('クラウドデータが読み込まれていません', 'warning'); return; }
  const data = cloudDataForPeriod();
  if (Object.keys(data).length === 0) { showToast('この期間のデータがありません', 'warning'); return; }

  const period = document.getElementById('cloudPeriodSel').value;
  const py = +period.slice(0, 4);
  const pm = +period.slice(5, 7);
  const nm = pm === 12 ? 1 : pm + 1;

  const blob = new Blob([JSON.stringify(data)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  // Excelと同じ R{令和}.{開始月}-{終了月} 形式
  a.download = `勤怠元データ_${modeCfg().label}_R${py - 2018}.${pm}-${nm}.json`;
  a.click();
  URL.revokeObjectURL(url);

  localStorage.setItem('lastRawSave_' + getMode(), new Date().toISOString());
  updateRawSaveInfo();
  showToast('元データを保存しました', 'success');
}

// 前回の元データ保存日を表示（1ヶ月超なら注意表示）
function updateRawSaveInfo() {
  const el = document.getElementById('rawSaveInfo');
  const last = localStorage.getItem('lastRawSave_' + getMode());
  if (!last) {
    el.innerHTML = '⚠ <b>元データはまだ保存されていません。</b>月1回の保存をおすすめします。';
    return;
  }
  const d = new Date(last);
  const days = Math.floor((Date.now() - d.getTime()) / 86400000);
  const dateStr = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
  el.innerHTML = days > 31
    ? `⚠ 前回の元データ保存: ${dateStr}（${days}日前）<b>そろそろ保存してください。</b>`
    : `前回の元データ保存: ${dateStr}`;
}

/* ============ 8. 入力・編集・分割・削除 ============ */

/* --- 有給用カレンダー（クリックで複数日選択） --- */

let paidSelectedDates = new Set();   // 選択中の日付 'YYYY-MM-DD'
let paidCalCursor = new Date();      // 表示中の月

function renderPaidCalendar() {
  const box = document.getElementById('paidCalendar');
  const y = paidCalCursor.getFullYear();
  const m = paidCalCursor.getMonth();
  const startDow = new Date(y, m, 1).getDay();
  const daysInMonth = new Date(y, m + 1, 0).getDate();
  const todayStr = toDateStr(new Date());

  let html = `
    <div class="cal-header">
      <button type="button" class="cal-nav" data-nav="-1">◀</button>
      <span>${y}年${m + 1}月</span>
      <button type="button" class="cal-nav" data-nav="1">▶</button>
    </div>
    <div class="cal-grid">
      ${['日','月','火','水','木','金','土'].map(w => `<div class="cal-dow">${w}</div>`).join('')}`;
  for (let i = 0; i < startDow; i++) html += '<div></div>';
  for (let d = 1; d <= daysInMonth; d++) {
    const ds = toDateStr(new Date(y, m, d));
    const cls = ['cal-day'];
    if (paidSelectedDates.has(ds)) cls.push('sel');
    if (ds === todayStr) cls.push('today');
    html += `<div class="${cls.join(' ')}" data-date="${ds}">${d}</div>`;
  }
  html += '</div>';
  if (paidSelectedDates.size > 0) {
    html += `<div class="cal-count">${paidSelectedDates.size}日選択中</div>`;
  }
  box.innerHTML = html;
}

// カレンダーの選択をテキスト欄（MM-DD表示）と西暦欄に反映
function syncDateMultiFromCalendar() {
  const dates = [...paidSelectedDates].sort();
  document.getElementById('dateMulti').value = dates.map(d => d.slice(5)).join(', ');
  if (dates.length > 0) document.getElementById('paidLeaveYear').value = +dates[0].slice(0, 4);
}

function saveTimeCard() {
  const name = document.getElementById('name').value.trim();
  const isPaidLeave = document.getElementById('isPaidLeave').checked;
  if (!name) { showToast('名前を入力してください', 'warning'); return; }

  const checkIn = document.getElementById('checkIn').value;
  const checkOut = document.getElementById('checkOut').value;
  if (!checkIn) { showToast('出勤時間を入力してください', 'warning'); return; }
  if (!checkOut) { showToast('退勤時間を入力してください（出勤・退勤の両方が必要です）', 'warning'); return; }

  // 有給の時間帯②（中抜けの人向け・任意）
  const checkIn2 = isPaidLeave ? document.getElementById('checkIn2').value : '';
  const checkOut2 = isPaidLeave ? document.getElementById('checkOut2').value : '';
  if ((checkIn2 && !checkOut2) || (!checkIn2 && checkOut2)) {
    showToast('時間帯②は出勤・退勤の両方を入力してください', 'warning');
    return;
  }

  const finalDates = collectDates(isPaidLeave);
  if (!finalDates) return;

  const cardsToAdd = [{ checkIn, checkOut: checkOut || null, isPaidLeave }];
  if (checkIn2 && checkOut2) cardsToAdd.push({ checkIn: checkIn2, checkOut: checkOut2, isPaidLeave });

  const data = loadData();
  if (!data[name]) data[name] = {};
  finalDates.forEach(dateStr => {
    if (!data[name][dateStr]) data[name][dateStr] = [];
    cardsToAdd.forEach(c => data[name][dateStr].push({ ...c }));
  });
  saveData(data);
  updateNameUsageOrder(name);
  refreshNameList();

  if (!isPaidLeave) {
    resetForm();
    showToast('送信が完了しました', 'success');
  } else {
    // 連続入力しやすいよう、有給モードのまま次の人の入力へ
    document.getElementById('name').value = '';
    document.getElementById('checkIn').value = '';
    document.getElementById('checkOut').value = '';
    document.getElementById('checkIn2').value = '';
    document.getElementById('checkOut2').value = '';
    document.getElementById('dateMulti').value = '';
    paidSelectedDates.clear();
    renderPaidCalendar();
    showToast(`送信が完了しました（有給 ${name}・${finalDates.length}日分）`, 'success');
    document.getElementById('name').focus();
  }
  displayTimeCards();
}

// フォームから日付（複数）を取り出す
function collectDates(isPaidLeave) {
  if (isPaidLeave) {
    // カレンダーで選択していればそれを優先（年またぎもそのまま使える）
    if (paidSelectedDates.size > 0) return [...paidSelectedDates].sort();
    const year = document.getElementById('paidLeaveYear').value.trim();
    const multi = document.getElementById('dateMulti').value.trim();
    if (!year)  { showToast('有給用の西暦を入力してください', 'warning'); return null; }
    if (!multi) { showToast('月日を入力してください', 'warning'); return null; }
    const splitted = multi.split(',')
      .map(s => s.trim())
      .filter(s => s !== '')
      .map(s => (s.length === 4 && s.indexOf('-') === -1) ? s.slice(0, 2) + '-' + s.slice(2) : s);
    if (splitted.length === 0) { showToast('月日を入力してください', 'warning'); return null; }
    return splitted.map(md => `${year}-${md}`);
  }
  const single = document.getElementById('dateSingle').value;
  if (!single) { showToast('月日を入力してください', 'warning'); return null; }
  return [single];
}

function editTimeCard(name, date, index) {
  const data = loadData();
  const card = data[name] && data[name][date] && data[name][date][index];
  if (!card) return;

  editingRecord = { name, date, index };
  document.getElementById('name').value = name;
  document.getElementById('dateSingle').value = date;
  document.getElementById('checkIn').value = card.checkIn || '';
  document.getElementById('checkOut').value = card.checkOut || '';
  document.getElementById('isPaidLeave').checked = !!card.isPaidLeave;

  if (card.isPaidLeave) {
    const yearMatch = date.match(/^(\d{4})-/);
    if (yearMatch) {
      document.getElementById('paidLeaveYear').value = yearMatch[1];
      document.getElementById('dateMulti').value = formatDate(date);
    }
    setPaidLeaveFields(true);
    // 編集は1件単位なので、カレンダー選択と時間帯②は使わない
    paidSelectedDates.clear();
    renderPaidCalendar();
    document.getElementById('secondTimeGroup').classList.add('hidden');
  }

  document.getElementById('submitBtn').textContent = '更新';
  document.getElementById('cancelEditBtn').classList.remove('hidden');
  document.getElementById('timeCardForm').scrollIntoView({ behavior: 'smooth' });
}

function updateTimeCard() {
  const name = document.getElementById('name').value.trim();
  const isPaidLeave = document.getElementById('isPaidLeave').checked;
  if (!name) { showToast('名前を入力してください', 'warning'); return; }

  const checkIn = document.getElementById('checkIn').value;
  const checkOut = document.getElementById('checkOut').value;
  if (!checkIn) { showToast('出勤時間を入力してください', 'warning'); return; }
  if (!checkOut) { showToast('退勤時間を入力してください（出勤・退勤の両方が必要です）', 'warning'); return; }

  const finalDates = collectDates(isPaidLeave);
  if (!finalDates) return;

  const data = loadData();

  // 元のレコードを削除
  data[editingRecord.name][editingRecord.date].splice(editingRecord.index, 1);
  if (data[editingRecord.name][editingRecord.date].length === 0) delete data[editingRecord.name][editingRecord.date];
  if (Object.keys(data[editingRecord.name]).length === 0) delete data[editingRecord.name];

  // 新しい内容で追加
  const card = { checkIn, checkOut: checkOut || null, isPaidLeave };
  if (!data[name]) data[name] = {};
  finalDates.forEach(dateStr => {
    if (!data[name][dateStr]) data[name][dateStr] = [];
    data[name][dateStr].push(card);
  });

  saveData(data);
  updateNameUsageOrder(name);
  refreshNameList();
  showToast('更新が完了しました', 'success');
  cancelEdit();
  displayTimeCards();
}

function cancelEdit() {
  editingRecord = null;
  resetForm();
  document.getElementById('submitBtn').textContent = '送信';
  document.getElementById('cancelEditBtn').classList.add('hidden');
}

function resetForm() {
  document.getElementById('timeCardForm').reset();
  document.getElementById('isPaidLeave').checked = false;
  setPaidLeaveFields(false);
}

function setPaidLeaveFields(on) {
  document.getElementById('normalDateGroup').classList.toggle('hidden', on);
  document.getElementById('paidLeaveYearGroup').classList.toggle('hidden', !on);
  document.getElementById('multiDateGroup').classList.toggle('hidden', !on);
  document.getElementById('secondTimeGroup').classList.toggle('hidden', !on);
  document.getElementById('dateSingle').required = !on;
  document.getElementById('paidLeaveYear').required = on;
  document.getElementById('dateMulti').required = on;
  if (on) {
    renderPaidCalendar();
  } else {
    paidSelectedDates.clear();
    document.getElementById('checkIn2').value = '';
    document.getElementById('checkOut2').value = '';
  }
}

async function deleteTimeCard(name, date, index) {
  const ok = await showConfirmDialog('記録の削除', `${name} ${formatDate(date)} の記録を削除しますか？`, '削除');
  if (!ok) return;
  const data = loadData();
  data[name][date].splice(index, 1);
  if (data[name][date].length === 0) delete data[name][date];
  if (Object.keys(data[name]).length === 0) delete data[name];
  saveData(data);
  displayTimeCards();
  showToast('記録を削除しました', 'success');
}

// 分割（中抜け対応）
async function splitTimeCard(name, date, index) {
  const data = loadData();
  const card = data[name] && data[name][date] && data[name][date][index];
  if (!card || !card.checkIn || !card.checkOut) {
    showToast('出勤・退勤時間が両方登録されている記録のみ分割できます', 'warning');
    return;
  }

  const breakStartInput = await showPromptDialog(
    '中抜け開始時間',
    `中抜け開始時間を入力してください（例: 12:00 または 1200）\n現在の記録: ${card.checkIn} 〜 ${card.checkOut}`);
  if (!breakStartInput) return;
  const breakStart = normalizeTimeInput(breakStartInput);
  if (!breakStart) { showToast('無効な時刻形式です（例: 12:00、1200、820）', 'error'); return; }

  const breakEndInput = await showPromptDialog(
    '中抜け終了時間',
    `中抜け終了時間を入力してください（例: 13:00 または 1300）\n中抜け開始: ${breakStart}`);
  if (!breakEndInput) return;
  const breakEnd = normalizeTimeInput(breakEndInput);
  if (!breakEnd) { showToast('無効な時刻形式です（例: 13:00、1300、830）', 'error'); return; }

  if (!isValidTimeOrder(card.checkIn, breakStart, breakEnd, card.checkOut)) {
    showToast(`時間の順序が正しくありません。\n出勤 < 中抜け開始 < 中抜け終了 < 退勤\n\n出勤: ${card.checkIn} / 開始: ${breakStart} / 終了: ${breakEnd} / 退勤: ${card.checkOut}`, 'error');
    return;
  }

  const ok = await showConfirmDialog(
    '記録を分割',
    `以下のように分割しますか？\n\n【1】${card.checkIn} 〜 ${breakStart}\n【2】${breakEnd} 〜 ${card.checkOut}`,
    '分割する');
  if (!ok) return;

  data[name][date].splice(index, 1);
  data[name][date].push({ checkIn: card.checkIn, checkOut: breakStart, isPaidLeave: !!card.isPaidLeave });
  data[name][date].push({ checkIn: breakEnd, checkOut: card.checkOut, isPaidLeave: !!card.isPaidLeave });
  saveData(data);
  showToast('記録を分割しました', 'success');
  displayTimeCards();
}

function isValidTimeOrder(t1, t2, t3, t4) {
  const ts = [t1, t2, t3, t4].map(t => new Date(`1970-01-01T${t}`));
  return ts[0] < ts[1] && ts[1] < ts[2] && ts[2] < ts[3];
}

function normalizeTimeInput(input) {
  if (!input) return null;
  input = input.trim();
  let h, m;
  if (/^\d{1,2}:\d{2}$/.test(input))      [h, m] = input.split(':').map(Number);
  else if (/^\d{4}$/.test(input))         { h = +input.slice(0, 2); m = +input.slice(2); }
  else if (/^\d{3}$/.test(input))         { h = +input.slice(0, 1); m = +input.slice(1); }
  else return null;
  if (h >= 0 && h <= 23 && m >= 0 && m <= 59) {
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  }
  return null;
}

async function requestPasswordAndClearData() {
  const password = await showPromptDialog('データクリア', `【${modeCfg().label}】の全データを削除します。パスワードを入力してください:`, '', 'password');
  if (password === null) return;
  if (password !== ADMIN_PASSWORD) { showToast('パスワードが違います', 'error'); return; }
  const ok = await showConfirmDialog('データクリア確認', `本当に【${modeCfg().label}】のすべてのデータをクリアしますか？`, '削除');
  if (!ok) return;
  localStorage.removeItem(modeCfg().storageKey);
  refreshNameList();
  displayTimeCards();
  showToast('すべてのデータを削除しました', 'success');
}

/* ============ 9. 一覧表示 ============ */

function refreshNameList() {
  const data = loadData();
  const usage = JSON.parse(localStorage.getItem(modeCfg().nameKey) || '{}');
  const names = Object.keys(data).sort((a, b) => (usage[b] || 0) - (usage[a] || 0));
  document.getElementById('nameList').innerHTML =
    names.map(n => `<option value="${esc(n)}">`).join('');
}

function displayTimeCards() {
  const filter = document.getElementById('searchName').value.trim().toLowerCase();
  const data = loadData();
  const resultDiv = document.getElementById('result');
  resultDiv.innerHTML = '';

  const names = Object.keys(data)
    .filter(name => name.toLowerCase().includes(filter))
    .sort((a, b) => a.localeCompare(b, 'ja'));

  let hasRecords = false;

  names.forEach(name => {
    const group = document.createElement('div');
    group.className = 'name-group';
    group.innerHTML = `<div class="name-header">${esc(name)}</div>`;

    Object.keys(data[name]).sort().forEach(date => {
      if (!Array.isArray(data[name][date])) return;
      data[name][date].forEach((card, index) => {
        if (!card || !card.checkIn) return;
        hasRecords = true;

        const paidLabel = card.isPaidLeave ? '<span class="record-status status-paid">有給</span>' : '';
        const checkOutDisplay = card.checkOut
          ? `<span class="record-time">${esc(card.checkOut)}</span>`
          : '<span class="record-status status-incomplete">未登録</span>';
        const splitBtn = card.checkOut
          ? `<button class="btn-xs split" data-action="split" data-name="${esc(name)}" data-date="${date}" data-index="${index}">分割</button>` : '';

        const row = document.createElement('div');
        row.className = 'record-card';
        row.innerHTML = `
          <div class="record-info">
            <input type="checkbox" class="rec-check" title="一括編集の対象に選択" data-name="${esc(name)}" data-date="${date}" data-index="${index}">
            <span class="record-date">${formatDate(date)}</span>
            <span class="record-time">${esc(card.checkIn)}</span>
            <span>→</span>
            ${checkOutDisplay}
            ${paidLabel}
          </div>
          <div class="record-buttons">
            <button class="btn-xs edit" data-action="edit" data-name="${esc(name)}" data-date="${date}" data-index="${index}">編集</button>
            ${splitBtn}
            <button class="btn-xs delete" data-action="delete" data-name="${esc(name)}" data-date="${date}" data-index="${index}">削除</button>
          </div>`;
        group.appendChild(row);
      });
    });

    if (group.children.length > 1) resultDiv.appendChild(group);
  });

  if (!hasRecords) {
    resultDiv.innerHTML = '<div class="no-data">まだ記録がありません</div>';
  }
  updateBulkBar();
}

/* --- 一括編集（チェックした記録をまとめて変更・削除） --- */

function selectedRecords() {
  return [...document.querySelectorAll('.rec-check:checked')].map(cb => ({
    name: cb.dataset.name, date: cb.dataset.date, index: +cb.dataset.index
  }));
}

function updateBulkBar() {
  const n = document.querySelectorAll('.rec-check:checked').length;
  document.getElementById('bulkBar').classList.toggle('hidden', n === 0);
  document.getElementById('bulkCount').textContent = `${n}件選択中`;
}

function showBulkEditDialog(count) {
  return new Promise(resolve => {
    const dialog = document.createElement('div');
    dialog.className = 'ios-dialog';
    dialog.innerHTML = `
      <div class="ios-dialog-content">
        <div class="ios-dialog-header">
          <div class="ios-dialog-title">まとめて時間変更（${count}件）</div>
          <div class="ios-dialog-message">空欄の項目は変更されません。\n（出勤だけ・退勤だけの修正もできます）</div>
          <div class="form-group" style="margin-top:14px;">
            <label>出勤時間</label>
            <input type="time" id="bulkCheckIn">
          </div>
          <div class="form-group">
            <label>退勤時間</label>
            <input type="time" id="bulkCheckOut">
          </div>
        </div>
        <div class="ios-dialog-buttons">
          <button class="ios-dialog-button">キャンセル</button>
          <button class="ios-dialog-button primary">変更する</button>
        </div>
      </div>`;
    document.body.appendChild(dialog);
    const [cancelBtn, okBtn] = dialog.querySelectorAll('.ios-dialog-button');
    const close = result => { dialog.style.opacity = '0'; setTimeout(() => dialog.remove(), 250); resolve(result); };
    cancelBtn.addEventListener('click', () => close(null));
    okBtn.addEventListener('click', () => close({
      checkIn: dialog.querySelector('#bulkCheckIn').value,
      checkOut: dialog.querySelector('#bulkCheckOut').value
    }));
    dialog.addEventListener('click', e => { if (e.target === dialog) close(null); });
  });
}

async function bulkEditTimes() {
  const recs = selectedRecords();
  if (recs.length === 0) return;
  const v = await showBulkEditDialog(recs.length);
  if (!v) return;
  if (!v.checkIn && !v.checkOut) { showToast('変更する時間を入力してください', 'warning'); return; }

  const data = loadData();
  let changed = 0;
  recs.forEach(r => {
    const card = data[r.name] && data[r.name][r.date] && data[r.name][r.date][r.index];
    if (!card) return;
    if (v.checkIn) card.checkIn = v.checkIn;
    if (v.checkOut) card.checkOut = v.checkOut;
    changed++;
  });
  saveData(data);
  displayTimeCards();
  showToast(`${changed}件の時間を変更しました`, 'success');
}

async function bulkDelete() {
  const recs = selectedRecords();
  if (recs.length === 0) return;
  const ok = await showConfirmDialog('まとめて削除', `選択した${recs.length}件の記録を削除しますか？`, '削除');
  if (!ok) return;

  const data = loadData();
  // 同じ名前・日付の中で index がずれないよう、大きい index から削除
  recs.sort((a, b) => b.index - a.index);
  recs.forEach(r => {
    if (!data[r.name] || !data[r.name][r.date]) return;
    data[r.name][r.date].splice(r.index, 1);
    if (data[r.name][r.date].length === 0) delete data[r.name][r.date];
    if (Object.keys(data[r.name]).length === 0) delete data[r.name];
  });
  saveData(data);
  refreshNameList();
  displayTimeCards();
  showToast(`${recs.length}件削除しました`, 'success');
}

/* ============ 10. Excelエクスポート ============ */

function calculateTimeDifference(startTime, endTime) {
  const start = new Date(`1970-01-01T${startTime}`);
  const end = new Date(`1970-01-01T${endTime}`);
  return parseFloat(((end - start) / 3600000).toFixed(2));
}

function calculateEarlyMorningTime(startTime, endTime) {
  const limit = new Date('1970-01-01T08:30');
  const start = new Date(`1970-01-01T${startTime}`);
  const end = new Date(`1970-01-01T${endTime}`);
  if (end <= limit) return calculateTimeDifference(startTime, endTime);
  if (start < limit) return calculateTimeDifference(startTime, '08:30');
  return 0;
}

function calculateEveningTime(startTime, endTime) {
  const limit = new Date('1970-01-01T16:00');
  const start = new Date(`1970-01-01T${startTime}`);
  const end = new Date(`1970-01-01T${endTime}`);
  if (start >= limit) return calculateTimeDifference(startTime, endTime);
  if (end > limit) return calculateTimeDifference('16:00', endTime);
  return 0;
}

// 基本の並び順（通常・スイミング共通の印刷順）。
// ツール上で並べ替えていないとき（初回や別PCで開いたとき）はこの順番が使われる。
// 順番を恒久的に変えたいときはここを書き換える。
const DEFAULT_EXPORT_ORDER = [
  '穂積', 'かおり', '外崎', 'るな', '新井郁絵', '延谷', '森田', '並',
  '佐藤雅恵', 'あやこ', '梅原瑞穂', '清水隆晟', '潮', '市川弘美',
  '古巣麻衣', '山内ちぐさ', '加藤厚子', '吉岡陽子', '岡本', '萩原'
];

// データが無くてもシートを出力する人（Excelに手書きで記入する人）
const ALWAYS_EXPORT_EMPTY = ['清水隆晟'];

// 並び順の保存先（通常・スイミング共通の1本）
const ORDER_KEY = 'exportOrderUnified';

// 保存済みの従業員並び順（Excelのシート順）。未登録の名前は末尾に五十音順で追加
function loadExportOrder() {
  try {
    const arr = JSON.parse(localStorage.getItem(ORDER_KEY) || '[]');
    if (Array.isArray(arr) && arr.length > 0) return arr;
  } catch {}
  return DEFAULT_EXPORT_ORDER;
}

function saveExportOrder(names) {
  localStorage.setItem(ORDER_KEY, JSON.stringify(names));
}

// 従業員ごとの属性（資格有無・バス添乗）。通常／スイミング共通で名前をキーに保存。
// 合計まとめシートの日給ランク（A〜E）振り分けに使う。
const EMP_ATTR_KEY = 'employeeAttrs';

// 初期値（資格＝「あり」、バス添乗＝「バス朝」。未保存の人はここの値が使われる）
const DEFAULT_EMP_ATTRS = {
  '穂積':   { qualified: true },
  '外崎':   { qualified: true },
  '新井郁絵': { bus: true },
  '延谷':   { qualified: true },
  'あやこ': { bus: true },
  '潮':     { qualified: true },
  '市川弘美': { qualified: true },
  '古巣麻衣': { qualified: true },
  '岡本':   { qualified: true },
  '萩原':   { qualified: true }
};

function loadEmployeeAttrs() {
  try {
    const o = JSON.parse(localStorage.getItem(EMP_ATTR_KEY) || '{}');
    return (o && typeof o === 'object' && !Array.isArray(o)) ? o : {};
  } catch { return {}; }
}

function getEmployeeAttr(name) {
  // 保存済みが最優先。未保存ならデフォルト（画像の初期値）。
  const saved = loadEmployeeAttrs();
  const a = saved[name] || DEFAULT_EMP_ATTRS[name] || {};
  return { qualified: !!a.qualified, bus: !!a.bus };
}

function setEmployeeAttr(name, attr) {
  const all = loadEmployeeAttrs();
  all[name] = { qualified: !!attr.qualified, bus: !!attr.bus };
  localStorage.setItem(EMP_ATTR_KEY, JSON.stringify(all));
}

// 日給ランク A〜E の時給（円）。設定で変更可、未設定は初期値。
const WAGE_KEY = 'rankWages';
const DEFAULT_WAGES = { A: 1300, B: 1200, C: 1500, D: 1400, E: 1500 };

function loadWages() {
  let saved = {};
  try {
    const o = JSON.parse(localStorage.getItem(WAGE_KEY) || '{}');
    if (o && typeof o === 'object' && !Array.isArray(o)) saved = o;
  } catch {}
  const w = {};
  for (const k of ['A', 'B', 'C', 'D', 'E']) {
    const v = Number(saved[k]);
    w[k] = Number.isFinite(v) && v > 0 ? v : DEFAULT_WAGES[k];
  }
  return w;
}

function saveWages(wages) {
  localStorage.setItem(WAGE_KEY, JSON.stringify(wages));
}

function orderedUserNames(names) {
  const saved = loadExportOrder();
  const known = saved.filter(n => names.includes(n));
  const rest = names.filter(n => !saved.includes(n)).sort((a, b) => a.localeCompare(b, 'ja'));
  return [...known, ...rest];
}

// エクスポート対象の名前一覧（scope: 'current'＝今のモードのみ／'both'＝通常＋スイミング）
function exportCandidateNames(scope) {
  const modes = scope === 'both' ? ['normal', 'swim'] : [getMode()];
  const names = new Set();
  modes.forEach(m => Object.keys(loadDataFor(m)).forEach(n => names.add(n)));
  // データが無くても出力する人（手書き用）は常にリストに載せる
  if (modes.includes('normal')) ALWAYS_EXPORT_EMPTY.forEach(n => names.add(n));
  return orderedUserNames([...names]);
}

function showExportDialog() {
  if (exportCandidateNames('current').length === 0) {
    showToast('エクスポートするデータがありません', 'warning');
    return;
  }

  const { start, end } = getLatestDataPeriod();

  const dialog = document.createElement('div');
  dialog.className = 'ios-dialog';
  dialog.innerHTML = `
    <div class="ios-dialog-content" style="width: 440px;">
      <div class="ios-dialog-header">
        <div class="ios-dialog-title">Excelエクスポート</div>
      </div>
      <div class="export-body">
        <div class="form-group">
          <label>出力対象</label>
          <select id="exportScope">
            <option value="current" selected>${esc(modeCfg().label)}のみ</option>
            <option value="both">通常＋スイミング（1つのファイルにまとめる）</option>
          </select>
        </div>
        <div class="form-group">
          <label>エクスポートする従業員（ドラッグまたは▲▼で並び順を変更・自動保存、上から順にシートになります）</label>
          <div class="user-checkbox-list" id="userOrderList"></div>
          <div class="rank-legend">
            <div class="rank-legend-title">合計まとめ A〜E 振り分け</div>
            <div><b>A</b> = 資格有 × 通常</div>
            <div><b>B</b> = 資格なし × 通常</div>
            <div><b>C</b> = 資格有 × 朝 ／ 資格有 × 夕 ／ 資格なし × 朝（バス添乗なし）</div>
            <div><b>D</b> = 資格なし × 朝（バス添乗あり） ／ 資格なし × 夕</div>
            <div><b>E</b> = スイミング（時間帯・資格問わず全時間）</div>
          </div>
        </div>
        <div class="form-group">
          <label>A〜E の時給（円・変更すると自動保存。Excelのヘッダーに表示されます）</label>
          <div class="wage-inputs" id="wageInputs"></div>
        </div>
        <div class="form-group">
          <label>エクスポート期間</label>
          <select id="exportPeriodType">
            <option value="all">すべてのデータ</option>
            <option value="latest" selected>最新の月（11日〜翌月10日）</option>
            <option value="custom">期間指定</option>
          </select>
        </div>
        <div class="form-group hidden" id="customPeriodFields">
          <label>開始日</label>
          <input type="date" id="exportStartDate" value="${toDateStr(start)}">
          <label style="margin-top:10px;">終了日</label>
          <input type="date" id="exportEndDate" value="${toDateStr(end)}">
        </div>
        <div class="form-group">
          <label class="checkbox-label"><input type="checkbox" id="enableRaiseDate"> 昇給日で分割する</label>
        </div>
        <div class="form-group hidden" id="raiseDateFields">
          <label>昇給日（この日付で期間を分割して別々に集計）</label>
          <input type="date" id="raiseDate">
        </div>
      </div>
      <div class="ios-dialog-buttons">
        <button class="ios-dialog-button" id="exportCancelBtn">キャンセル</button>
        <button class="ios-dialog-button primary" id="exportExecuteBtn">エクスポート</button>
      </div>
    </div>`;
  document.body.appendChild(dialog);

  const closeDialog = () => { dialog.style.opacity = '0'; setTimeout(() => dialog.remove(), 250); };

  // A〜E の時給入力（変更で自動保存）
  const wageBox = dialog.querySelector('#wageInputs');
  const wages = loadWages();
  wageBox.innerHTML = ['A', 'B', 'C', 'D', 'E'].map(k => `
    <label class="wage-field"><span class="wage-rank">${k}</span>
      <input type="number" class="wageInput" data-rank="${k}" min="0" step="10" value="${wages[k]}">円</label>`).join('');
  wageBox.addEventListener('change', e => {
    if (!e.target.classList.contains('wageInput')) return;
    const cur = loadWages();
    const v = Number(e.target.value);
    if (Number.isFinite(v) && v > 0) cur[e.target.dataset.rank] = Math.round(v);
    saveWages(cur);
  });

  const orderList = dialog.querySelector('#userOrderList');

  // 従業員リストの描画（出力対象の切り替えで作り直す）
  const userRowHtml = n => {
    const a = getEmployeeAttr(n);
    return `
    <div class="user-row" data-name="${esc(n)}" draggable="true">
      <span class="drag-handle" title="ドラッグで並べ替え">⠿</span>
      <label><input type="checkbox" class="userCheckbox" value="${esc(n)}" checked> ${esc(n)}</label>
      <span class="emp-attrs" title="合計まとめのA〜E振り分けに使用">
        <label title="資格有"><input type="checkbox" class="attrQualified" ${a.qualified ? 'checked' : ''}> 資格</label>
        <label title="バス添乗"><input type="checkbox" class="attrBus" ${a.bus ? 'checked' : ''}> バス</label>
      </span>
      <span class="order-btns">
        <button type="button" class="order-btn" data-dir="up" title="上へ">▲</button>
        <button type="button" class="order-btn" data-dir="down" title="下へ">▼</button>
      </span>
    </div>`;
  };

  const renderUserList = scope => {
    orderList.innerHTML =
      `<label class="select-all"><input type="checkbox" id="selectAllUsers" checked> すべて選択</label>` +
      exportCandidateNames(scope).map(userRowHtml).join('');
  };
  renderUserList('current');

  dialog.querySelector('#exportScope').addEventListener('change', function () {
    renderUserList(this.value);
  });

  // 全選択チェックボックス（リストを作り直しても効くように委譲で処理）
  orderList.addEventListener('change', e => {
    // 資格・バス添乗の変更は即保存
    if (e.target.classList.contains('attrQualified') || e.target.classList.contains('attrBus')) {
      const row = e.target.closest('.user-row');
      setEmployeeAttr(row.dataset.name, {
        qualified: row.querySelector('.attrQualified').checked,
        bus: row.querySelector('.attrBus').checked
      });
      return;
    }
    const selectAll = orderList.querySelector('#selectAllUsers');
    const boxes = [...orderList.querySelectorAll('.userCheckbox')];
    if (e.target === selectAll) {
      boxes.forEach(cb => { cb.checked = selectAll.checked; });
      selectAll.indeterminate = false;
      return;
    }
    if (e.target.classList.contains('userCheckbox')) {
      const all = boxes.every(c => c.checked);
      const some = boxes.some(c => c.checked);
      selectAll.checked = all;
      selectAll.indeterminate = some && !all;
    }
  });

  // 並べ替え：▲▼ボタン＋ドラッグ&ドロップ。動かしたら即保存。
  // リストに表示されていない名前（別モードの人）は元の位置を保ったままマージする
  const saveOrderFromDom = () => {
    const domNames = [...orderList.querySelectorAll('.user-row')].map(r => r.dataset.name);
    const old = loadExportOrder();
    const merged = [...domNames];
    old.forEach((name, i) => {
      if (merged.includes(name)) return;
      let insertAt = 0;
      for (let j = i - 1; j >= 0; j--) {
        const idx = merged.indexOf(old[j]);
        if (idx !== -1) { insertAt = idx + 1; break; }
      }
      merged.splice(insertAt, 0, name);
    });
    saveExportOrder(merged);
  };

  orderList.addEventListener('click', e => {
    const btn = e.target.closest('.order-btn');
    if (!btn) return;
    const row = btn.closest('.user-row');
    if (btn.dataset.dir === 'up') {
      const prev = row.previousElementSibling;
      if (prev && prev.classList.contains('user-row')) orderList.insertBefore(row, prev);
    } else {
      const next = row.nextElementSibling;
      if (next) orderList.insertBefore(next, row);
    }
    saveOrderFromDom();
  });

  let dragRow = null;
  orderList.addEventListener('dragstart', e => {
    dragRow = e.target.closest('.user-row');
    if (!dragRow) return;
    dragRow.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    // Firefoxはこれが無いとドラッグが始まらない
    e.dataTransfer.setData('text/plain', dragRow.dataset.name);
  });
  orderList.addEventListener('dragover', e => {
    if (!dragRow) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    const over = e.target.closest('.user-row');
    if (!over || over === dragRow) return;
    const rect = over.getBoundingClientRect();
    const below = e.clientY > rect.top + rect.height / 2;
    orderList.insertBefore(dragRow, below ? over.nextSibling : over);
  });
  orderList.addEventListener('drop', e => e.preventDefault());
  orderList.addEventListener('dragend', () => {
    if (!dragRow) return;
    dragRow.classList.remove('dragging');
    dragRow = null;
    saveOrderFromDom();
  });

  dialog.querySelector('#exportPeriodType').addEventListener('change', function () {
    dialog.querySelector('#customPeriodFields').classList.toggle('hidden', this.value !== 'custom');
  });
  dialog.querySelector('#enableRaiseDate').addEventListener('change', function () {
    dialog.querySelector('#raiseDateFields').classList.toggle('hidden', !this.checked);
  });

  dialog.querySelector('#exportCancelBtn').addEventListener('click', closeDialog);
  dialog.querySelector('#exportExecuteBtn').addEventListener('click', () => performExport(dialog, closeDialog));
  dialog.addEventListener('click', e => { if (e.target === dialog) closeDialog(); });
}

function performExport(dialog, closeDialog) {
  const selectedUsers = [...dialog.querySelectorAll('.userCheckbox:checked')].map(cb => cb.value);
  if (selectedUsers.length === 0) { showToast('少なくとも1人の従業員を選択してください', 'warning'); return; }

  const scope = dialog.querySelector('#exportScope').value;
  const exportType = dialog.querySelector('#exportPeriodType').value;
  const enableRaiseDate = dialog.querySelector('#enableRaiseDate').checked;
  const raiseDateValue = enableRaiseDate ? dialog.querySelector('#raiseDate').value : null;

  let startDate, endDate;
  if (exportType === 'all') {
    startDate = new Date('1900-01-01');
    endDate = new Date('2100-12-31');
  } else if (exportType === 'latest') {
    const { start, end } = getLatestDataPeriod();
    startDate = start;
    endDate = end;
  } else {
    const s = dialog.querySelector('#exportStartDate').value;
    const e = dialog.querySelector('#exportEndDate').value;
    if (!s || !e) { showToast('開始日と終了日を入力してください', 'warning'); return; }
    startDate = new Date(s);
    endDate = new Date(e);
    if (startDate > endDate) { showToast('開始日は終了日より前である必要があります', 'warning'); return; }
  }

  if (enableRaiseDate) {
    if (!raiseDateValue) { showToast('昇給日を入力してください', 'warning'); return; }
    const rd = new Date(raiseDateValue);
    if (rd <= startDate || rd >= endDate) {
      showToast('昇給日は開始日と終了日の間の日付を指定してください', 'warning');
      return;
    }
  }

  closeDialog();
  exportToExcel(startDate, endDate, exportType, selectedUsers, raiseDateValue ? new Date(raiseDateValue) : null, scope);
}

function exportToExcel(startDate, endDate, exportType, selectedUsers, raiseDate, scope) {
  // scope: 'current'＝今のモードのみ／'both'＝通常＋スイミングを1つのワークブックに
  const modesToExport = scope === 'both' ? ['normal', 'swim'] : [getMode()];
  const dataByMode = {};
  modesToExport.forEach(m => { dataByMode[m] = loadDataFor(m); });
  const workbook = XLSX.utils.book_new();
  let sheetCount = 0;
  let actualMinDate = null;
  let actualMaxDate = null;

  const adjustedEndDate = new Date(endDate);
  adjustedEndDate.setHours(23, 59, 59, 999);

  // 昇給日が指定されている場合は期間を2つに分割
  let periods;
  if (raiseDate) {
    const beforeEnd = new Date(raiseDate);
    beforeEnd.setDate(beforeEnd.getDate() - 1);
    beforeEnd.setHours(23, 59, 59, 999);
    periods = [
      { start: startDate, end: beforeEnd, label: '昇給前' },
      { start: new Date(raiseDate), end: adjustedEndDate, label: '昇給後' }
    ];
  } else {
    periods = [{ start: startDate, end: adjustedEndDate, label: '' }];
  }

  // 出力シート（1人×1モード＝1シート）をダイアログの表示順に組み立てる。
  // 両モードにデータがある人は 通常→スイミング の順で並べ、シート名に（通常）（スイミング）を付ける
  const tasks = [];
  for (const name of selectedUsers) {
    const targets = modesToExport.filter(m =>
      (dataByMode[m][name] && Object.keys(dataByMode[m][name]).length > 0) ||
      (m === 'normal' && ALWAYS_EXPORT_EMPTY.includes(name)));
    const dual = targets.length > 1;
    for (const m of targets) tasks.push({ name, mode: m, dual });
  }

  // 最後のまとめシート用：各人の全期間合計を日給ランク A〜E に振り分けて集める
  //   A=資格有×通常 / B=資格なし×通常
  //   C=資格有×朝 + 資格有×夕 + 資格なし×朝（バス無）
  //   D=資格なし×朝（バス有）+ 資格なし×夕 / E=スイミング（全時間）
  // 通常とスイミング両方ある人は1行にまとめる（A〜Dは通常、Eはスイミング由来）
  const summaryByName = {};
  const summaryOrder = [];

  for (const { name, mode, dual } of tasks) {
    const cards = dataByMode[mode][name] || {};
    const workType = MODES[mode].label;
    // データが無くてもシートを出す人（手書き用）
    const forceEmpty = mode === 'normal' && ALWAYS_EXPORT_EMPTY.includes(name);

    let sheetName = name;
    if (dual) {
      sheetName = `${name}(${workType})`;
    } else if (exportType !== 'custom' && !raiseDate) {
      const monthNames = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"];
      sheetName = name + monthNames[startDate.getMonth()];
    }

    const sheetData = [];
    sheetData.push([`名前: ${name} (${workType})`]);
    if (!raiseDate) {
      sheetData.push(["日付","出勤時間","退勤時間","合計時間","早朝勤務","夕方勤務","朝夕合計","通常合計","勤務種別"]);
    }

    let hasAnyData = false;
    // この人（×モード）の全期間を通した合計
    let gTotal = 0, gEarly = 0, gEvening = 0, gNormal = 0;

    for (const { start: pStart, end: pEnd, label: pLabel } of periods) {
      const periodData = [];

      for (const date in cards) {
        if (!Array.isArray(cards[date])) continue;
        const d = new Date(date);
        if (d < pStart || d > pEnd) continue;
        if (!actualMinDate || d < actualMinDate) actualMinDate = d;
        if (!actualMaxDate || d > actualMaxDate) actualMaxDate = d;

        cards[date].forEach(card => {
          if (!card || !card.checkIn) return;
          let totalHours = 0, earlyMorning = 0, evening = 0;
          if (card.checkOut) {
            totalHours = calculateTimeDifference(card.checkIn, card.checkOut);
            earlyMorning = calculateEarlyMorningTime(card.checkIn, card.checkOut);
            evening = calculateEveningTime(card.checkIn, card.checkOut);
          }
          periodData.push({
            date,
            checkIn: card.checkIn,
            checkOut: card.checkOut || '未登録',
            totalHours, earlyMorning, evening,
            isPaidLeave: card.isPaidLeave
          });
        });
      }

      if (periodData.length === 0) continue;
      hasAnyData = true;

      if (pLabel) {
        const startStr = `${pStart.getMonth() + 1}/${pStart.getDate()}`;
        const endStr = `${pEnd.getMonth() + 1}/${pEnd.getDate()}`;
        sheetData.push([]);
        sheetData.push([`■ ${pLabel} (${startStr}〜${endStr})`]);
        sheetData.push(["日付","出勤時間","退勤時間","合計時間","早朝勤務","夕方勤務","朝夕合計","通常合計","勤務種別"]);
      }

      let sumTotal = 0, sumEarly = 0, sumEvening = 0, sumME = 0, sumNormal = 0;
      periodData.sort((a, b) => new Date(a.date) - new Date(b.date));

      periodData.forEach(rec => {
        const morningEvening = rec.earlyMorning + rec.evening;
        const normalHours = rec.totalHours - morningEvening;
        if (rec.checkOut !== '未登録') {
          sumTotal += rec.totalHours;
          sumEarly += rec.earlyMorning;
          sumEvening += rec.evening;
          sumME += morningEvening;
          sumNormal += normalHours;
        }
        const noOut = rec.checkOut === '未登録';
        sheetData.push([
          formatDate(rec.date),
          rec.checkIn,
          rec.checkOut,
          noOut ? '－' : rec.totalHours.toFixed(2),
          noOut ? '－' : rec.earlyMorning.toFixed(2),
          noOut ? '－' : rec.evening.toFixed(2),
          noOut ? '－' : morningEvening.toFixed(2),
          noOut ? '－' : normalHours.toFixed(2),
          rec.isPaidLeave ? "有給" : ""
        ]);
      });

      sheetData.push([]);
      sheetData.push([
        pLabel ? `【${pLabel}合計】` : "合計", "", "",
        sumTotal.toFixed(2), sumEarly.toFixed(2), sumEvening.toFixed(2),
        sumME.toFixed(2), sumNormal.toFixed(2), ""
      ]);

      gTotal += sumTotal; gEarly += sumEarly; gEvening += sumEvening;
      gNormal += sumNormal;
    }

    if (hasAnyData) {
      let row = summaryByName[name];
      if (!row) { row = summaryByName[name] = { name, a: 0, b: 0, c: 0, d: 0, e: 0 }; summaryOrder.push(name); }
      if (mode === 'swim') {
        row.e += gTotal;                       // スイミングは全時間 → E
      } else {
        const attr = getEmployeeAttr(name);
        if (attr.qualified) {
          row.a += gNormal;                    // 資格有×通常 → A
          row.c += gEarly + gEvening;          // 資格有×朝夕 → C
        } else {
          row.b += gNormal;                    // 資格なし×通常 → B
          if (attr.bus) {
            row.d += gEarly + gEvening;         // 資格なし×朝(バス) + 夕 → D
          } else {
            row.c += gEarly;                    // 資格なし×朝(バス無) → C
            row.d += gEvening;                  // 資格なし×夕 → D
          }
        }
      }
    }

    if (!hasAnyData) {
      if (!forceEmpty) continue;
      // 手書き用の空シート：昇給分割時は見出し行が未出力なので足しておく
      if (raiseDate) {
        sheetData.push(["日付","出勤時間","退勤時間","合計時間","早朝勤務","夕方勤務","朝夕合計","通常合計","勤務種別"]);
      }
    }
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    // 列幅：ヘッダー（出勤時間・退勤時間 等の4文字）が切れず全部見えるように広めに確保
    worksheet['!cols'] = [
      { wch: 12 }, // 日付
      { wch: 11 }, // 出勤時間
      { wch: 11 }, // 退勤時間
      { wch: 11 }, // 合計時間
      { wch: 11 }, // 早朝勤務
      { wch: 11 }, // 夕方勤務
      { wch: 11 }, // 朝夕合計
      { wch: 11 }, // 通常合計
      { wch: 11 }  // 勤務種別
    ];
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.slice(0, 31));
    sheetCount++;
  }

  if (sheetCount === 0) {
    showToast('選択した期間にデータがありません', 'warning');
    return;
  }

  // 最後のシート：各人の合計時間を日給ランク A〜E に振り分けてまとめる
  if (summaryOrder.length > 0) {
    const wages = loadWages();
    const summaryData = [];
    summaryData.push(["合計時間まとめ（単位：時間）"]);
    summaryData.push([]);
    summaryData.push([
      "名前",
      `A ${wages.A}円`,
      `B ${wages.B}円`,
      `C ${wages.C}円`,
      `D ${wages.D}円`,
      `E ${wages.E}円`
    ]);

    summaryOrder.forEach(n => {
      const r = summaryByName[n];
      summaryData.push([
        r.name,
        r.a.toFixed(2),
        r.b.toFixed(2),
        r.c.toFixed(2),
        r.d.toFixed(2),
        r.e.toFixed(2)
      ]);
    });

    // A〜E 振り分け表（凡例）。各行は6列ぶん結合して見やすく表示
    summaryData.push([]);
    const legendStart = summaryData.length;
    const legendLines = [
      "【A〜E 振り分け】",
      "A = 資格有 × 通常",
      "B = 資格なし × 通常",
      "C = 資格有 × 朝 ／ 資格有 × 夕 ／ 資格なし × 朝（バス添乗なし）",
      "D = 資格なし × 朝（バス添乗あり） ／ 資格なし × 夕",
      "E = スイミング（時間帯・資格問わず全時間）"
    ];
    legendLines.forEach(line => summaryData.push([line]));

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
    // 列幅（名前は広め、数値列もゆったり）
    summarySheet['!cols'] = [
      { wch: 24 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }
    ];
    // 全行に高さを設定して詰まりを解消（タイトル＝大、区切り＝小、その他＝ゆったり）
    summarySheet['!rows'] = summaryData.map((row, i) => {
      if (i === 0) return { hpt: 30 };                 // タイトル
      if (row.length === 0) return { hpt: 10 };        // 区切りの空行
      return { hpt: 24 };                              // 見出し・データ・凡例
    });
    // タイトル＋凡例の各行を6列ぶん結合
    summarySheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }
    ];
    legendLines.forEach((_, i) => {
      const r = legendStart + i;
      summarySheet['!merges'].push({ s: { r, c: 0 }, e: { r, c: 5 } });
    });
    XLSX.utils.book_append_sheet(workbook, summarySheet, "合計まとめ");
  }

  // ファイル名: R{令和}.{開始月}-{終了月}（スイミングは_SW付き）
  let filename;
  if (exportType === 'all' && actualMinDate && actualMaxDate) {
    const minP = getMonthPeriod(toDateStr(actualMinDate));
    const maxP = getMonthPeriod(toDateStr(actualMaxDate));
    filename = `R${minP.start.getFullYear() - 2018}.${minP.start.getMonth() + 1}-${maxP.end.getMonth() + 1}`;
  } else if (exportType === 'latest') {
    const { start, end } = getLatestDataPeriod();
    filename = `R${start.getFullYear() - 2018}.${start.getMonth() + 1}-${end.getMonth() + 1}`;
  } else {
    filename = `R${startDate.getFullYear() - 2018}.${startDate.getMonth() + 1}.${startDate.getDate()}-${endDate.getMonth() + 1}.${endDate.getDate()}`;
  }
  if (scope === 'both') filename += '_通常+SW';
  else if (getMode() === 'swim') filename += '_SW';
  filename += '.xlsx';

  XLSX.writeFile(workbook, filename);
  showToast('Excelファイルを出力しました', 'success');
}

/* ============ 11. バックアップ・リストア ============ */

function backupData() {
  const data = loadData();
  const now = new Date();
  const stamp = toDateStr(now).replace(/-/g, '');
  // Excelと同じ R{令和}.{開始月}-{終了月} を付ける（同じ期間に複数回とっても区別できるよう日付も残す）
  const { start, end } = getLatestDataPeriod();
  const period = `R${start.getFullYear() - 2018}.${start.getMonth() + 1}-${end.getMonth() + 1}`;
  const blob = new Blob([JSON.stringify(data)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `勤怠バックアップ_${modeCfg().label}_${period}_${stamp}.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('データをバックアップしました', 'success');
}

function showRestoreDialog(file) {
  const dialog = document.createElement('div');
  dialog.className = 'ios-dialog';
  dialog.innerHTML = `
    <div class="ios-dialog-content">
      <div class="ios-dialog-header">
        <div class="ios-dialog-title">データ復元方法</div>
        <div class="ios-dialog-message">【${esc(modeCfg().label)}】への復元方法を選択してください</div>
      </div>
      <div class="ios-dialog-buttons" style="flex-direction: column;">
        <button class="ios-dialog-button" data-act="merge" style="border-left:none;border-bottom:1px solid var(--border);">統合（既存データに追加・重複は除外）</button>
        <button class="ios-dialog-button primary" data-act="overwrite" style="border-left:none;border-bottom:1px solid var(--border);">上書き（既存データを置き換え）</button>
        <button class="ios-dialog-button" data-act="cancel" style="border-left:none;">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(dialog);

  const close = () => { dialog.style.opacity = '0'; setTimeout(() => dialog.remove(), 250); };

  dialog.addEventListener('click', e => {
    if (e.target === dialog) { close(); return; }
    const act = e.target.dataset && e.target.dataset.act;
    if (!act) return;
    close();
    if (act === 'merge') restoreData(file, 'merge');
    else if (act === 'overwrite') restoreData(file, 'overwrite');
  });
}

function restoreData(file, mode) {
  const reader = new FileReader();
  reader.onload = function (event) {
    try {
      const incoming = JSON.parse(event.target.result);
      if (!incoming || typeof incoming !== 'object' || Array.isArray(incoming)) {
        showToast('無効なデータ形式です', 'error');
        return;
      }
      if (mode === 'merge') {
        const { merged, added } = mergeTimeCards(loadData(), incoming);
        saveData(merged);
        showToast(`データを統合しました（${added}件追加）`, 'success');
      } else {
        saveData(incoming);
        showToast('データを上書きしました', 'success');
      }
      refreshNameList();
      displayTimeCards();
    } catch {
      showToast('データの読み込み中にエラーが発生しました', 'error');
    }
  };
  reader.readAsText(file);
}

/* ============ 12. 初期化 ============ */

document.addEventListener('DOMContentLoaded', () => {
  applyTheme(localStorage.getItem('theme') || 'light');
  document.getElementById('themeToggle').addEventListener('click', toggleTheme);

  // モード切替
  document.querySelectorAll('#modeSwitch button').forEach(btn =>
    btn.addEventListener('click', () => setMode(btn.dataset.mode)));

  // 有給チェック（パスワード式）
  document.getElementById('isPaidLeave').addEventListener('change', async function () {
    if (this.checked) {
      const pwd = await showPromptDialog('有給登録', '有給として登録するためのパスワードを入力してください:', '', 'password');
      if (pwd === ADMIN_PASSWORD) {
        setPaidLeaveFields(true);
        document.getElementById('paidLeaveYear').value = new Date().getFullYear();
      } else {
        if (pwd) showToast('パスワードが違います', 'error');
        this.checked = false;
      }
    } else {
      setPaidLeaveFields(false);
    }
  });

  // 有給カレンダー（クリックで日付を選択／解除、◀▶で月移動）
  document.getElementById('paidCalendar').addEventListener('click', e => {
    const nav = e.target.closest('.cal-nav');
    if (nav) {
      paidCalCursor = new Date(paidCalCursor.getFullYear(), paidCalCursor.getMonth() + (+nav.dataset.nav), 1);
      renderPaidCalendar();
      return;
    }
    const day = e.target.closest('.cal-day');
    if (!day) return;
    const ds = day.dataset.date;
    if (paidSelectedDates.has(ds)) paidSelectedDates.delete(ds);
    else paidSelectedDates.add(ds);
    syncDateMultiFromCalendar();
    renderPaidCalendar();
  });

  // 日付欄を手で書き換えたらカレンダー選択は解除（手入力を優先）
  document.getElementById('dateMulti').addEventListener('input', () => {
    if (paidSelectedDates.size > 0) {
      paidSelectedDates.clear();
      renderPaidCalendar();
    }
  });

  // フォーム送信
  document.getElementById('timeCardForm').addEventListener('submit', e => {
    e.preventDefault();
    if (editingRecord) updateTimeCard();
    else saveTimeCard();
  });

  // Enterで次の入力欄へ移動（最後の欄でEnterすると送信）
  document.getElementById('timeCardForm').addEventListener('keydown', e => {
    if (e.key !== 'Enter') return;
    const t = e.target;
    if (t.tagName !== 'INPUT' || t.type === 'checkbox') return;
    e.preventDefault();
    const fields = [...e.currentTarget.querySelectorAll('input')]
      .filter(el => el.type !== 'checkbox' && el.offsetParent !== null);
    const next = fields[fields.indexOf(t) + 1];
    if (next) next.focus();
    else e.currentTarget.requestSubmit();
  });
  document.getElementById('cancelEditBtn').addEventListener('click', cancelEdit);

  // クラウドカードの折りたたみ（状態を記憶）
  const setCloudCollapsed = collapsed => {
    document.getElementById('cloudBody').classList.toggle('hidden', collapsed);
    document.getElementById('cloudArrow').textContent = collapsed ? '▸ ひらく' : '▾';
    localStorage.setItem('cloudCardCollapsed', collapsed ? '1' : '');
  };
  document.getElementById('cloudCardToggle').addEventListener('click', () =>
    setCloudCollapsed(!document.getElementById('cloudBody').classList.contains('hidden')));
  setCloudCollapsed(localStorage.getItem('cloudCardCollapsed') === '1');

  // 修正申請シート：現在モードのシートを画面右半分のウィンドウで開く（アプリと並べて見られる）
  document.getElementById('fixSheetBtn').addEventListener('click', () => {
    const w = Math.floor(screen.availWidth / 2);
    const h = screen.availHeight;
    window.open(FIX_SHEET_URLS[getMode()], 'fixSheet_' + getMode(), `width=${w},height=${h},left=${w},top=0`);
  });

  // ツールバー
  document.getElementById('exportBtn').addEventListener('click', showExportDialog);
  document.getElementById('backupBtn').addEventListener('click', backupData);
  document.getElementById('clearDataBtn').addEventListener('click', requestPasswordAndClearData);
  document.getElementById('restoreBtn').addEventListener('click', () => document.getElementById('restoreFile').click());
  document.getElementById('restoreFile').addEventListener('change', function (e) {
    if (e.target.files[0]) {
      showRestoreDialog(e.target.files[0]);
      e.target.value = '';
    }
  });

  // Googleドライブ（保存先フォルダ）：画面右半分のウィンドウで開く
  document.getElementById('driveBtn').addEventListener('click', () => {
    const w = Math.floor(screen.availWidth / 2);
    const h = screen.availHeight;
    window.open(DRIVE_FOLDER_URL, 'driveFolder', `width=${w},height=${h},left=${w},top=0`);
  });

  // 一括編集バー
  document.getElementById('result').addEventListener('change', e => {
    if (e.target.classList.contains('rec-check')) updateBulkBar();
  });
  document.getElementById('bulkEditBtn').addEventListener('click', bulkEditTimes);
  document.getElementById('bulkDeleteBtn').addEventListener('click', bulkDelete);
  document.getElementById('bulkClearBtn').addEventListener('click', () => {
    document.querySelectorAll('.rec-check:checked').forEach(cb => { cb.checked = false; });
    updateBulkBar();
  });

  // 検索
  document.getElementById('searchName').addEventListener('input', displayTimeCards);
  document.getElementById('searchClearBtn').addEventListener('click', () => {
    document.getElementById('searchName').value = '';
    displayTimeCards();
  });

  // 入力欄に名前を打つと一覧も自動で絞り込み（検索バーに反映されるので手で変更・解除も可能）
  document.getElementById('name').addEventListener('input', function () {
    document.getElementById('searchName').value = this.value;
    displayTimeCards();
  });

  // 一覧のボタン（イベント委譲：名前に記号があっても安全）
  document.getElementById('result').addEventListener('click', e => {
    const btn = e.target.closest('button[data-action]');
    if (!btn) return;
    const { action, name, date, index } = btn.dataset;
    if (action === 'edit') editTimeCard(name, date, +index);
    else if (action === 'split') splitTimeCard(name, date, +index);
    else if (action === 'delete') deleteTimeCard(name, date, +index);
  });

  // クラウド
  document.getElementById('cloudLoginBtn').addEventListener('click', cloudLogin);
  document.getElementById('cloudPassword').addEventListener('keydown', e => { if (e.key === 'Enter') cloudLogin(); });
  document.getElementById('cloudLogoutBtn').addEventListener('click', cloudLogout);
  document.getElementById('cloudReloadBtn').addEventListener('click', cloudLoadSnapshots);
  document.getElementById('cloudSnapshotSel').addEventListener('change', function () { cloudLoadSnapshot(this.value); });
  document.getElementById('cloudPeriodSel').addEventListener('change', renderCloudSummary);
  document.getElementById('cloudImportBtn').addEventListener('click', cloudImport);
  document.getElementById('cloudRawSaveBtn').addEventListener('click', cloudRawSave);

  // 初期表示
  setMode(getMode());
  if (cloudToken) showCloudMain();
});
