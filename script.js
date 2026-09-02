/* ============================================================
 *  年次有給休暇 管理ツール 1.0.0
 *
 *  入力はこのツールだけ。付与・取得・残日数を自動計算し、
 *  TC5（勤怠ツール）へは「リストア→統合」で読める形のJSONを出力する。
 *
 *   1. 定数・ユーティリティ
 *   2. データ層（localStorage）
 *   3. 付与・消化の計算エンジン
 *   3.5 削除（個別・まとめて・条件で絞って）
 *   4. 従業員マスタ
 *   5. 有給入力
 *   6. 残日数一覧・管理簿（パート方式＝見本Excelと同じ体裁）
 *   6.5 正職集計（正職方式＝年度・時間）
 *   7. 既存Excelの取込
 *   8. 書き出し（TC5用JSON・バックアップ）
 *   8.5 年度ごとのファイル保存（フォルダ連携）
 *   9. 初期化
 * ============================================================ */

/* ============ 1. 定数・ユーティリティ ============ */

const STORE_KEY = 'yk_data_v1';
const TC5_MODES = { normal: '通常', swim: 'スイミング' };

// 労働基準法の付与日数表。キー＝週の所定労働日数、値＝[6ヶ月, 1年6ヶ月, … 6年6ヶ月以降]
const GRANT_TABLE = {
  5: [10, 11, 12, 14, 16, 18, 20],
  4: [ 7,  8,  9, 10, 12, 13, 15],
  3: [ 5,  6,  6,  8,  9, 10, 11],
  2: [ 3,  4,  4,  5,  6,  6,  7],
  1: [ 1,  2,  2,  2,  3,  3,  3]
};

// 年5日の取得義務が始まった日
const DUTY_START = '2019-04-01';

// 管理方式。従業員ごとにどちらかを選ぶ
//   part  … パート方式：見本Excelの管理簿（入社日基準で付与・繰越・2年で失効）
//   staff … 正職方式：年度（4月〜翌3月）ごとに時間で記録し、1日の勤務時間で日数換算
const LEDGER_TYPES = { part: 'パート方式（管理簿）', staff: '正職方式（年度・時間）' };

// 正職方式の長期休暇の最大取得日数（年度ごとに変更できる）
const DEFAULT_MAX_LEAVE = { summer: 16, winter: 5, spring: 3 };

// 長期休暇の入力の呼び方。どちらを選んでも「最大日数 − その数 ＝ 取得日数」で同じ。
// 数え方の言い方が人によって違うので選べるようにしてある。
const TERM_MODES = { work: '出勤数', absent: '欠勤数' };

// 年度の月の並び（4月始まり）と曜日
const FISCAL_MONTHS = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3];
const DOW = ['日', '月', '火', '水', '木', '金', '土'];
const TERM_FIELDS = [
  { key: 'sWork',  max: 'summer', label: '夏期' },
  { key: 'wWork',  max: 'winter', label: '冬期' },
  { key: 'spWork', max: 'spring', label: '春期' }
];

// TC5の登録名。ソースには名前を持たせない（このツールは公開の場所に置くため）。
// 「👤 従業員情報」→「TC5の登録名」で貼り付けると DB.tcNames に保存される。
const DEFAULT_TC5_NAMES = [];

const pad2 = n => String(n).padStart(2, '0');
const esc  = s => String(s == null ? '' : s).replace(/[&<>"']/g,
  c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

// 長期休暇の入力を「出勤数」で書くか「欠勤数」で書くか
function termMode() {
  return (DB && DB.termMode === 'absent') ? 'absent' : 'work';
}
function termModeLabel() { return TERM_MODES[termMode()]; }

// 実際に使うTC5の登録名（保存済みがあればそれ、無ければ初期値）
function tcNames() {
  return (DB && Array.isArray(DB.tcNames) && DB.tcNames.length) ? DB.tcNames : DEFAULT_TC5_NAMES;
}

function toDateStr(d) { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }
function parseDate(s) { const [y, m, d] = String(s).split('-').map(Number); return new Date(y, m - 1, d); }
function todayStr() { return toDateStr(new Date()); }

// 月を足す（末日は月末に丸める）
function addMonths(dateStr, months) {
  const d = parseDate(dateStr);
  const day = d.getDate();
  const t = new Date(d.getFullYear(), d.getMonth() + months, 1);
  const last = new Date(t.getFullYear(), t.getMonth() + 1, 0).getDate();
  t.setDate(Math.min(day, last));
  return toDateStr(t);
}

function addDays(dateStr, days) {
  const d = parseDate(dateStr);
  d.setDate(d.getDate() + days);
  return toDateStr(d);
}

function fmtDate(s) {
  if (!s) return '';
  const [y, m, d] = s.split('-');
  return `${y}/${+m}/${+d}`;
}

// 小数の日数を「◯日◯時間」表記に
function fmtDays(days, dailyHours) {
  const h = dailyHours || 8;
  const sign = days < 0 ? '-' : '';
  let n = Math.abs(Math.round(days * 1000) / 1000);
  let d = Math.floor(n + 1e-9);
  let rest = Math.round((n - d) * h * 4) / 4;
  if (rest >= h - 1e-9) { d += 1; rest = 0; }
  if (rest === 0) return `${sign}${d}日`;
  return `${sign}${d}日${rest}時間`;
}

function showToast(message, type = 'info') {
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = message;
  document.getElementById('toastArea').appendChild(el);
  setTimeout(() => { el.style.opacity = '0'; setTimeout(() => el.remove(), 300); }, 3200);
}

function download(filename, text) {
  const blob = new Blob([text], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

// 10日締め（11日〜翌月10日）の期間を返す
function closingPeriod(baseDate, offset = 0) {
  const d = parseDate(baseDate);
  let y = d.getFullYear(), m = d.getMonth();      // 0-based
  if (d.getDate() <= 10) m -= 1;                  // 10日以前なら前月11日始まり
  m += offset;
  const start = new Date(y, m, 11);
  const end   = new Date(y, m + 1, 10);
  return { start: toDateStr(start), end: toDateStr(end) };
}

// 年度（4月始まり）。2026-03-31 は 2025年度
function fiscalYearOf(dateStr) {
  const [y, m] = dateStr.split('-').map(Number);
  return m >= 4 ? y : y - 1;
}
function fiscalRange(year) {
  return { start: `${year}-04-01`, end: `${year + 1}-03-31` };
}

// "2:30" や "2.5" を時間数に（正職ツールと同じ読み方）
function parseHM(str) {
  if (str == null) return 0;
  let t = String(str).trim();
  if (t === '') return 0;
  let neg = false;
  if (t[0] === '-') { neg = true; t = t.slice(1); }
  let v;
  if (t.includes(':')) {
    const [h, m] = t.split(':');
    v = (parseFloat(h) || 0) + (parseFloat(m) || 0) / 60;
  } else v = parseFloat(t);
  if (isNaN(v)) return 0;
  return neg ? -v : v;
}

// 時間数を "2:30" 形式に
function fmtHM(hours) {
  const sign = hours < 0 ? '-' : '';
  const mins = Math.round(Math.abs(hours) * 60);
  return `${sign}${Math.floor(mins / 60)}:${pad2(mins % 60)}`;
}

// 正職ツールと同じ「○日＋○時間」表記
function toDaysPlus(total, daily) {
  if (!daily || daily <= 0) return `${round2(total)}時間`;
  const d = Math.trunc(total / daily);
  return `${d}日＋${round2(total - d * daily)}時間`;
}

const round2 = n => Math.round(n * 100) / 100;

// R8.5-6 のような期間ラベル
function periodLabel(fromStr, toStr) {
  const f = parseDate(fromStr), t = parseDate(toStr);
  return `R${f.getFullYear() - 2018}.${f.getMonth() + 1}-${t.getMonth() + 1}`;
}

/* ============ 2. データ層 ============ */

// { employees: [ … ], leaves: [ … ], fiscal: { '2026': { maxLeave, work } } }
let DB = { employees: [], leaves: [], fiscal: {}, tcNames: [], termMode: 'work' };

function loadDB() {
  try {
    const o = JSON.parse(localStorage.getItem(STORE_KEY) || 'null');
    if (o && Array.isArray(o.employees) && Array.isArray(o.leaves)) {
      if (!o.fiscal || typeof o.fiscal !== 'object') o.fiscal = {};
      if (!Array.isArray(o.tcNames)) o.tcNames = [...DEFAULT_TC5_NAMES];
      if (o.termMode !== 'absent') o.termMode = 'work';
      // 方式が入っていない古いデータはパート方式とみなす
      o.employees.forEach(e => { if (!e.type) e.type = 'part'; });
      return o;
    }
  } catch {}
  return { employees: [], leaves: [], fiscal: {}, tcNames: [...DEFAULT_TC5_NAMES], termMode: 'work' };
}

// 正職方式の年度設定（長期休暇の最大日数と、各人の出勤数）
function getFiscal(year) {
  const k = String(year);
  if (!DB.fiscal) DB.fiscal = {};
  if (!DB.fiscal[k]) DB.fiscal[k] = { maxLeave: { ...DEFAULT_MAX_LEAVE }, work: {} };
  const f = DB.fiscal[k];
  if (!f.maxLeave) f.maxLeave = { ...DEFAULT_MAX_LEAVE };
  if (!f.work) f.work = {};
  return f;
}

function getTermWork(year, empId) {
  const w = getFiscal(year).work;
  if (!w[empId]) w[empId] = { sWork: '', wWork: '', spWork: '' };
  return w[empId];
}

// これまでに記録のある年度＋今年度
function knownFiscalYears() {
  const set = new Set(DB.leaves.map(l => fiscalYearOf(l.date)));
  Object.keys(DB.fiscal || {}).forEach(y => set.add(Number(y)));
  set.add(fiscalYearOf(todayStr()));
  return [...set].sort((a, b) => b - a);
}

function saveDB() {
  localStorage.setItem(STORE_KEY, JSON.stringify(DB));
  if (typeof scheduleFileWrite === 'function') scheduleFileWrite();   // フォルダ連携中なら年度ファイルにも
}

function newId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function getEmp(id) { return DB.employees.find(e => e.id === id) || null; }

function leavesOf(empId) {
  return DB.leaves.filter(l => l.empId === empId)
                  .sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0);
}

// 表示順：区分 → TC5の並び順 → 氏名
function sortedEmployees() {
  return [...DB.employees].sort((a, b) => {
    if (a.mode !== b.mode) return a.mode === 'normal' ? -1 : 1;
    const order = tcNames();
    const ia = order.indexOf(a.tcName), ib = order.indexOf(b.tcName);
    if (ia !== ib) return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib);
    return (a.name || '').localeCompare(b.name || '', 'ja');
  });
}

/* ============ 3. 付与・消化の計算エンジン ============ */
/*
 *  見本Excel「年次有給休暇管理簿」の 1行＝1回の付与 と同じ計算をする。
 *
 *    有給休暇発生日 B … 入社日＋6ヶ月、以後1年ごと
 *    付与       E … 週労働日数と勤続年数から法定表を引く（「無効」なら0）
 *    前年繰越   F … IF(前の繰越 >= 前の取得, 前の付与, 前の残日数)
 *                    ＝古い付与から先に使い、使い残しは2年で失効、という意味
 *    計         G … E + F
 *    取得       H … 対象期間（発生日〜1年後の前日）に取った日数
 *    義務残     I … 5 − 取得（付与10日以上・2019/4/1以降。時間単位は含めない）
 *    残日数     J … G − H
 */

// 1件の有給が何日ぶんか
function leaveDays(leave, emp) {
  if (leave.type === 'full')  return 1;
  if (leave.type === 'am' || leave.type === 'pm') return 0.5;
  if (leave.type === 'hours') {
    const h = Number(leave.hours) || 0;
    const dh = Number(emp && emp.dailyHours) || 8;
    return h / dh;
  }
  return 0;
}

const TYPE_LABEL = { full: '全日', am: '午前半休', pm: '午後半休', hours: '時間単位' };

// 時間単位の有給の表示（元の "2:30" 表記があればそれを使う）
function leaveHoursText(l) {
  return l.hoursText || fmtHM(Number(l.hours) || 0);
}

// 1件の有給が何時間ぶんか（正職方式はこちらが基準）
function leaveHours(leave, emp) {
  const dh = Number(emp && emp.dailyHours) || 8;
  if (leave.type === 'hours') return Number(leave.hours) || 0;
  return leaveDays(leave, emp) * dh;
}

/**
 * 正職方式の年度集計。正職ツールの集計表とまったく同じ計算をする。
 *   有給取得日数 ＝ 年度の合計時間 ÷ 1日の勤務時間（「○日＋余り時間」）
 *   長期休暇の取得日数 ＝ 最大日数 − 出勤数（出勤数が空欄なら0日）
 */
function computeStaffYear(emp, year) {
  const { start, end } = fiscalRange(year);
  const takes = leavesOf(emp.id).filter(l => l.date >= start && l.date <= end);
  const totalHours = takes.reduce((s, l) => s + leaveHours(l, emp), 0);
  const daily = Number(emp.dailyHours) || 0;
  const days = daily > 0 ? Math.trunc(totalHours / daily) : 0;
  const rest = daily > 0 ? totalHours - days * daily : totalHours;

  const f = getFiscal(year);
  const work = getTermWork(year, emp.id);
  const mode = termMode();
  const terms = TERM_FIELDS.map(t => {
    const raw = work[t.key];
    const max = Number(f.maxLeave[t.max]) || 0;
    const has = !(raw === '' || raw == null);
    const num = has ? (parseFloat(raw) || 0) : 0;
    // 出勤数で入れたとき … 取得＝最大−出勤数、残り＝出勤数
    // 欠勤数で入れたとき … 欠勤＝もう休んだ日数なので 取得＝欠勤数、残り＝最大−欠勤数
    const taken  = !has ? 0   : (mode === 'absent' ? num : max - num);
    const remain = !has ? max : (mode === 'absent' ? max - num : num);
    return { ...t, max, work: raw, value: num, has, taken, remain };
  });
  const longSum = terms.reduce((s, t) => s + t.taken, 0);
  const longRemain = terms.reduce((s, t) => s + t.remain, 0);

  return {
    takes, count: takes.length, totalHours, daily, days, rest,
    terms, longSum, longRemain,
    grandDays: days + longSum, grandRest: rest,
    maxLeave: f.maxLeave
  };
}

// 入社日から horizon までに発生する付与を全部作る
// その回の付与に使う週労働日数（Excelと同じく付与ごとに変えられる）
function weekDaysAt(emp, grantDate) {
  const per = emp.weekDaysByGrant && emp.weekDaysByGrant[grantDate];
  const w = Number(per) || Number(emp.weekDays) || 5;
  return Math.min(Math.max(Math.round(w), 1), 5);   // 週6日は週5日と同じ扱い（法定表どおり）
}

function buildGrants(emp, horizon) {
  const grants = [];
  if (!emp.hire) return grants;
  const invalid = emp.invalidGrants || {};
  for (let i = 0; i < 60; i++) {
    const months = 6 + i * 12;
    const date = addMonths(emp.hire, months);
    if (date > horizon) break;
    const weekDays = weekDaysAt(emp, date);
    const table = GRANT_TABLE[weekDays] || GRANT_TABLE[5];
    const base = (emp.grantOverrides && emp.grantOverrides[date] != null)
      ? Number(emp.grantOverrides[date])
      : table[Math.min(i, table.length - 1)];
    grants.push({ months, date, base, weekDays, invalid: !!invalid[date] });
  }
  return grants;
}

// 勤続年数の表示（見本のC6と同じ「◯年◯ヶ月」）
function serviceLength(hire, asOf) {
  if (!hire) return '';
  const h = parseDate(hire), a = parseDate(asOf);
  let months = (a.getFullYear() - h.getFullYear()) * 12 + (a.getMonth() - h.getMonth());
  if (a.getDate() < h.getDate()) months -= 1;
  if (months < 0) return '入社前';
  return `${Math.floor(months / 12)}年${months % 12}ヶ月`;
}

/**
 * 従業員1人の管理簿を丸ごと組み立てる。
 * rows が見本Excelの1行1行にそのまま対応する。
 */
function computeLedger(emp, asOf) {
  const leaves = leavesOf(emp.id);
  const maxLeaveDate = leaves.length ? leaves[leaves.length - 1].date : asOf;
  // 先の日付の有給を入れてあっても計算できるよう、少し先まで付与を作る
  let horizon = addMonths(asOf, 18);
  if (maxLeaveDate > horizon) horizon = maxLeaveDate;

  const rows = [];
  let prev = null;
  for (const g of buildGrants(emp, horizon)) {
    const from = g.date;
    const to = addDays(addMonths(from, 12), -1);          // 対象期間の終わり
    const takes = leaves.filter(l => l.date >= from && l.date <= to);
    const taken    = takes.reduce((s, l) => s + leaveDays(l, emp), 0);
    const dayTaken = takes.filter(l => l.type !== 'hours')
                          .reduce((s, l) => s + leaveDays(l, emp), 0);
    const hourTaken = taken - dayTaken;

    const grant = g.invalid ? 0 : g.base;
    // 見本の F列と同じ式
    const carry = !prev ? 0 : (prev.carry >= prev.taken ? prev.grant : prev.remain);
    const total = grant + carry;
    const remain = total - taken;

    const row = {
      months: g.months, from, to, grant, base: g.base, invalid: g.invalid,
      weekDays: g.weekDays,
      carry, total, taken, dayTaken, hourTaken, remain, takes,
      duty: (grant >= 10 && from >= DUTY_START) ? Math.max(0, 5 - dayTaken) : null,
      lapsed: prev ? Math.max(0, prev.remain - carry) : 0,   // 前の期で失効した分
      started: from <= asOf,
      current: from <= asOf && asOf <= to
    };
    rows.push(row);
    prev = row;
  }

  const current = rows.find(r => r.current) || null;
  const next    = rows.find(r => r.from > asOf) || null;
  const balance = current ? current.remain : 0;
  const expired = rows.filter(r => r.to < asOf).reduce((s, r) => s + r.lapsed, 0)
                + (current ? current.lapsed : 0);

  // 残日数がマイナスになった期＝使いすぎ
  const shortages = rows.filter(r => r.started && r.remain < -1e-9)
                        .map(r => ({ from: r.from, to: r.to, days: -r.remain }));

  // 時間単位年休は年5日ぶんが上限
  const hourlyOver = current && current.hourTaken > 5 + 1e-9 ? current.hourTaken : null;

  return {
    rows, leaves, current, next, balance, expired, shortages, hourlyOver,
    grantDays:     current ? current.grant  : 0,
    carry:         current ? current.carry  : 0,
    takenInPeriod: current ? current.taken  : 0,
    duty: (current && current.duty !== null)
      ? { taken: current.dayTaken, remain: current.duty, deadline: current.to }
      : null
  };
}

/* ============ 3.5 削除 ============ */

const SOURCE_LABEL = { manual: '手入力', excel: 'Excel取込', staffjson: '正職ツール取込' };

// 有給をまとめて消す。消した件数を返す
function removeLeaves(ids) {
  const set = new Set(ids);
  const before = DB.leaves.length;
  DB.leaves = DB.leaves.filter(l => !set.has(l.id));
  return before - DB.leaves.length;
}

// 従業員をまとめて消す（その人の有給・年度設定も一緒に消える）
function removeEmployees(ids) {
  const set = new Set(ids);
  DB.employees = DB.employees.filter(e => !set.has(e.id));
  DB.leaves = DB.leaves.filter(l => !set.has(l.empId));
  for (const y in (DB.fiscal || {})) {
    const w = DB.fiscal[y].work || {};
    for (const id of set) delete w[id];
  }
  return set.size;
}

// 画面でチェックが入っている行のid
function checkedIds(selector) {
  return [...document.querySelectorAll(selector)].filter(c => c.checked).map(c => c.dataset.id);
}

/* --- 条件で絞って有給を消す（取込をやり直したいとき用） --- */

function purgeTargets() {
  const src  = document.getElementById('purgeSource').value;
  const empId = document.getElementById('purgeEmp').value;
  const from = document.getElementById('purgeFrom').value;
  const to   = document.getElementById('purgeTo').value;
  return DB.leaves.filter(l => {
    if (src && (l.source || 'manual') !== src) return false;
    if (empId && l.empId !== empId) return false;
    if (from && l.date < from) return false;
    if (to && l.date > to) return false;
    return true;
  });
}

function renderPurgePreview() {
  const box = document.getElementById('purgePreview');
  const hits = purgeTargets();
  if (!hits.length) { box.innerHTML = '<span class="hint">条件に合う有給はありません。</span>'; return; }
  const byEmp = {};
  for (const l of hits) {
    const e = getEmp(l.empId);
    const k = e ? e.name : '(不明)';
    byEmp[k] = (byEmp[k] || 0) + 1;
  }
  const names = Object.keys(byEmp).sort();
  box.innerHTML = `<b class="danger-text">${hits.length}件</b> が対象です（${names.length}人）`
    + `<ul>${names.slice(0, 12).map(n => `<li>${esc(n)}：${byEmp[n]}件</li>`).join('')}`
    + (names.length > 12 ? `<li>ほか ${names.length - 12}人</li>` : '') + '</ul>';
}

function refreshPurgeEmpSelect() {
  const sel = document.getElementById('purgeEmp');
  const cur = sel.value;
  sel.innerHTML = '<option value="">全員</option>' + sortedEmployees()
    .map(e => `<option value="${e.id}">${esc(e.name)}${e.mode === 'swim' ? '（スイミング）' : ''}</option>`).join('');
  if ([...sel.options].some(o => o.value === cur)) sel.value = cur;
}

/* ============ 4. 従業員マスタ ============ */

function readEmpForm() {
  const num = id => Number(document.getElementById(id).value);
  return {
    id:        document.getElementById('empId').value || null,
    name:      document.getElementById('empName').value.trim(),
    kana:      document.getElementById('empKana').value.trim(),
    tcName:    document.getElementById('empTcName').value.trim(),
    type:      document.getElementById('empType').value,
    mode:      document.getElementById('empMode').value,
    hire:      document.getElementById('empHire').value,
    weekDays:  num('empWeekDays'),
    dailyHours: num('empDailyHours') || 8,
    shift: {
      in1:  document.getElementById('empT1a').value,
      out1: document.getElementById('empT1b').value,
      in2:  document.getElementById('empT2a').value,
      out2: document.getElementById('empT2b').value
    },
    active: document.getElementById('empActive').checked
  };
}

function fillEmpForm(emp) {
  const set = (id, v) => { document.getElementById(id).value = v == null ? '' : v; };
  set('empId', emp ? emp.id : '');
  set('empType', emp ? (emp.type || 'part') : 'part');
  set('empName', emp ? emp.name : '');
  set('empKana', emp ? emp.kana : '');
  set('empTcName', emp ? emp.tcName : '');
  set('empMode', emp ? emp.mode : 'normal');
  set('empHire', emp ? emp.hire : '');
  set('empWeekDays', emp ? emp.weekDays : 5);
  set('empDailyHours', emp ? emp.dailyHours : 8);
  set('empT1a', emp && emp.shift ? emp.shift.in1 : '');
  set('empT1b', emp && emp.shift ? emp.shift.out1 : '');
  set('empT2a', emp && emp.shift ? emp.shift.in2 : '');
  set('empT2b', emp && emp.shift ? emp.shift.out2 : '');
  document.getElementById('empActive').checked = emp ? emp.active !== false : true;

  applyEmpTypeFields();
  document.getElementById('empFormTitle').textContent = emp ? '従業員を編集' : '従業員を追加';
  document.getElementById('empCancel').classList.toggle('hidden', !emp);
  document.getElementById('empDelete').classList.toggle('hidden', !emp);
}

// パート方式なら入社日・週労働日数、正職方式なら1日の勤務時間が要る
function applyEmpTypeFields() {
  const type = document.getElementById('empType').value;
  const part = type === 'part';
  document.getElementById('empPartFields').classList.toggle('hidden', !part);
  // 正職はTC5を使わないので、TC5での名前と所定の勤務時間帯は出さない
  document.getElementById('empTcField').classList.toggle('hidden', !part);
  document.getElementById('empShiftField').classList.toggle('hidden', !part);
  document.getElementById('empTypeHint').textContent = part
    ? '入社日から6ヶ月・以後1年ごとに付与し、繰越・2年での失効・年5日の義務まで管理します（見本Excelと同じ）。TC5へも出せます。'
    : '年度（4月〜翌3月）ごとに、有給を時間で積み上げて「1日の勤務時間」で日数に換算します。夏期・冬期・春期の長期休暇は別枠。TC5は使いません。';
}

function saveEmployee() {
  const f = readEmpForm();
  if (!f.name && !f.tcName) { showToast('氏名を入力してください', 'warning'); return; }
  // TC5を使うのはパートだけなので、正職はTC5での名前が無くてよい
  if (f.type === 'part' && !f.tcName) {
    showToast('パート方式はTC5での名前が必須です（JSONのキーになります）', 'warning'); return;
  }
  if (f.type === 'staff') f.tcName = '';
  if (!f.name) f.name = f.tcName;
  if (f.type === 'part' && !f.hire) {
    showToast('パート方式は入社日が必要です（付与日の計算に使います）', 'warning'); return;
  }
  if (f.type === 'staff' && !(f.dailyHours > 0)) {
    showToast('正職方式は1日の勤務時間が必要です（日数換算に使います）', 'warning'); return;
  }

  const dup = DB.employees.find(e => e.tcName === f.tcName && e.mode === f.mode && e.id !== f.id);
  if (dup) { showToast(`「${f.tcName}」（${TC5_MODES[f.mode]}）は既に登録されています`, 'warning'); return; }

  if (f.id) {
    const emp = getEmp(f.id);
    Object.assign(emp, f);
  } else {
    f.id = newId();
    f.grantOverrides = {};
    DB.employees.push(f);
  }
  saveDB();
  fillEmpForm(null);
  renderAll();
  showToast('従業員を保存しました', 'success');
}

function deleteEmployee() {
  const id = document.getElementById('empId').value;
  const emp = getEmp(id);
  if (!emp) return;
  const n = leavesOf(id).length;
  if (!confirm(`「${emp.name}」を削除します。\n登録済みの有給 ${n} 件も一緒に消えます。よろしいですか？`)) return;
  DB.employees = DB.employees.filter(e => e.id !== id);
  DB.leaves = DB.leaves.filter(l => l.empId !== id);
  saveDB();
  fillEmpForm(null);
  renderAll();
  showToast('削除しました', 'success');
}

/* --- 同じ人が二重に入っていないか --- */

// 「山田」と「山田　花子」のように書き方が違っても同じ人とみなす（取込のたびに氏名の書き方が変わるため）
const normName = v => String(v || '').replace(/[\s　]/g, '');

function samePerson(a, b) {
  if (a.mode !== b.mode) return false;                 // 通常とスイミングは別扱い
  const na = normName(a.name), nb = normName(b.name);
  if (a.tcName && a.tcName === b.tcName) return true;
  if (!na || !nb) return false;
  if (na === nb) return true;
  if (na.startsWith(nb) || nb.startsWith(na)) return true;
  if (a.hire && a.hire === b.hire && na.slice(0, 2) === nb.slice(0, 2)) return true;
  return false;
}

function duplicateGroups() {
  const list = sortedEmployees();
  const used = new Set();
  const groups = [];
  for (let i = 0; i < list.length; i++) {
    if (used.has(list[i].id)) continue;
    const g = [list[i]];
    for (let j = i + 1; j < list.length; j++) {
      if (used.has(list[j].id)) continue;
      if (g.some(x => samePerson(x, list[j]))) { g.push(list[j]); used.add(list[j].id); }
    }
    if (g.length > 1) { g.forEach(e => used.add(e.id)); groups.push(g); }
  }
  return groups;
}

// keepId の人に others をまとめる（有給は日付＋単位が同じものだけ捨てる）
function mergeEmployees(keepId, otherIds) {
  const keep = getEmp(keepId);
  if (!keep) return 0;
  let moved = 0;
  for (const oid of otherIds) {
    const other = getEmp(oid);
    if (!other || other.id === keep.id) continue;

    // 空いている項目は相手の値で埋める
    for (const k of ['kana', 'tcName', 'hire', 'dailyHours']) {
      if (!keep[k] && other[k]) keep[k] = other[k];
    }
    if (!keep.shift || !keep.shift.in1) if (other.shift && other.shift.in1) keep.shift = { ...other.shift };
    keep.weekDaysByGrant = { ...(other.weekDaysByGrant || {}), ...(keep.weekDaysByGrant || {}) };
    keep.invalidGrants  = { ...(other.invalidGrants  || {}), ...(keep.invalidGrants  || {}) };
    keep.grantOverrides = { ...(other.grantOverrides || {}), ...(keep.grantOverrides || {}) };
    // 長い方の氏名を残す（「山田」より「山田　花子」）
    if (normName(other.name).length > normName(keep.name).length) keep.name = other.name;

    for (const l of DB.leaves) {
      if (l.empId !== oid) continue;
      const dup = DB.leaves.some(x =>
        x.empId === keep.id && x.date === l.date && x.type === l.type);
      if (dup) { l._drop = true; continue; }
      l.empId = keep.id;
      moved++;
    }
    // 年度設定（正職方式の長期休暇の出勤数）も引き継ぐ
    for (const y in (DB.fiscal || {})) {
      const w = DB.fiscal[y].work || {};
      if (w[oid]) { if (!w[keep.id]) w[keep.id] = w[oid]; delete w[oid]; }
    }
    DB.employees = DB.employees.filter(e => e.id !== oid);
  }
  DB.leaves = DB.leaves.filter(l => !l._drop);
  return moved;
}

function renderDupCard() {
  const card = document.getElementById('dupCard');
  const groups = duplicateGroups();
  if (!groups.length) { card.classList.add('hidden'); return; }
  card.classList.remove('hidden');

  document.getElementById('dupArea').innerHTML = groups.map(g => {
    const rows = g.map(e => {
      const n = leavesOf(e.id).length;
      const src = [...new Set(leavesOf(e.id).map(l => l.source))].join('・') || '—';
      return `<tr>
        <td>${esc(e.name)}${e.kana ? ` <span class="hint">${esc(e.kana)}</span>` : ''}</td>
        <td>${(e.type || 'part') !== 'part' ? '<span class="hint">—</span>'
              : (esc(e.tcName) || '<span class="danger-text">未設定</span>')}</td>
        <td>${(e.type || 'part') === 'part' ? 'パート' : '正職'}</td>
        <td>${fmtDate(e.hire) || '—'}</td>
        <td class="num">${n}件</td>
        <td><span class="hint">${esc(src)}</span></td>
        <td>
          <button type="button" class="btn mini" data-keep="${e.id}">これにまとめる</button>
          <button type="button" class="btn mini danger" data-drop="${e.id}">この行を削除</button>
        </td>
      </tr>`;
    }).join('');
    return `<div class="dup-group">
      <table><thead><tr>
        <th>氏名</th><th>TC5表記</th><th>方式</th><th>入社日</th><th>有給</th><th>取込元</th><th></th>
      </tr></thead><tbody>${rows}</tbody></table>
    </div>`;
  }).join('');
}

/* --- TC5の登録名（TC5の勤怠一覧に出ている名前）の管理 --- */

function renderTcNames() {
  const box = document.getElementById('tcNameArea');
  const list = tcNames();
  const used = new Set(DB.employees.filter(e => (e.type || 'part') === 'part' && e.tcName).map(e => e.tcName));
  const unused = list.filter(n => !used.has(n));
  const unknown = [...used].filter(n => !list.includes(n));

  document.getElementById('tcNameText').value = list.join('\n');
  if (!list.length) {
    box.innerHTML = '<div class="hint">まだ登録されていません。TC5の勤怠一覧に出ている名前を貼り付けて「保存」を押してください。</div>';
    return;
  }
  let html = `<div class="hint">登録名 ${list.length}件 ／ 使用中 ${used.size}件</div>`;
  if (unused.length) {
    html += `<div class="hint">まだ誰にも割り当てていない名前：`
      + unused.map(n => `<span class="take full">${esc(n)}</span>`).join('') + '</div>';
  }
  if (unknown.length && list.length) {
    html += `<div class="hint danger-text">この一覧に無い名前が使われています：`
      + unknown.map(n => `<span class="take half">${esc(n)}</span>`).join('')
      + '（TC5側の表記と違っていないか確認してください）</div>';
  }
  box.innerHTML = html;
}

/* TC5のバックアップJSON（{名前:{日付:[…]}}）から名前だけ取り出す */
function namesFromTc5Backup(obj) {
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return [];
  const names = [];
  for (const key in obj) {
    const days = obj[key];
    if (!days || typeof days !== 'object' || Array.isArray(days)) continue;
    // 値が「日付 → 配列」になっていればTC5の勤怠データとみなす
    const looksLikeDays = Object.keys(days).some(d =>
      /^\d{4}-\d{2}-\d{2}$/.test(d) && Array.isArray(days[d]));
    if (looksLikeDays) names.push(key);
  }
  return names;
}

function handleTcNameJsonFiles(files) {
  const found = [];
  let done = 0, bad = 0;
  for (const file of files) {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const names = namesFromTc5Backup(JSON.parse(ev.target.result));
        if (!names.length) bad++;
        names.forEach(n => { if (!found.includes(n)) found.push(n); });
      } catch { bad++; }
      if (++done === files.length) {
        if (!found.length) {
          showToast('TC5の勤怠データが見つかりませんでした（バックアップJSONを選んでください）', 'warning');
          return;
        }
        const cur = tcNames();
        const add = found.filter(n => !cur.includes(n));
        const already = found.length - add.length;
        if (!add.length) {
          showToast(`${found.length}人ぶん読みましたが、全員すでに登録済みです`, 'info');
          return;
        }
        if (!confirm(`TC5の勤怠データから ${found.length}人 見つかりました。\n`
          + `新しく追加する ${add.length}人：${add.join('、')}\n`
          + (already ? `（${already}人はすでに登録済み）\n` : '')
          + '\n今の一覧のうしろに追加します。よろしいですか？')) return;
        DB.tcNames = [...cur, ...add];
        saveDB();
        renderAll();
        showToast(`${add.length}人を追加しました${bad ? `（${bad}ファイルは読めませんでした）` : ''}`, 'success');
      }
    };
    reader.readAsText(file);
  }
}

function saveTcNames() {
  const list = document.getElementById('tcNameText').value
    .split(/[\n,、]/).map(v => v.trim()).filter(Boolean);
  if (!list.length) { showToast('1つ以上入力してください', 'warning'); return; }
  const dup = list.filter((n, i) => list.indexOf(n) !== i);
  if (dup.length) { showToast(`同じ名前が2回入っています：${[...new Set(dup)].join('、')}`, 'warning'); return; }
  DB.tcNames = list;
  saveDB();
  renderAll();
  showToast(`TC5の登録名を ${list.length}件 保存しました`, 'success');
}

function renderEmpTable() {
  const tbody = document.querySelector('#empTable tbody');
  const list = sortedEmployees();
  document.getElementById('empCount').textContent = `${list.length}人`;
  const ea = document.getElementById('empCheckAll'); if (ea) ea.checked = false;

  if (!list.length) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="8">まだ登録がありません。左のフォームか、「取込・出力」タブの取込で追加してください。</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(e => {
    const sh = e.shift || {};
    const t = sh.in1 ? `${sh.in1}〜${sh.out1}` + (sh.in2 ? ` / ${sh.in2}〜${sh.out2}` : '') : '—';
    const type = (e.type || 'part') === 'part'
      ? '<span class="tag full">パート方式</span>'
      : '<span class="tag hours">正職方式</span>';
    return `<tr${e.active === false ? ' style="opacity:.5"' : ''}>
      <td><input type="checkbox" class="empChk" data-id="${e.id}"></td>
      <td>${esc(e.name)}</td>
      <td>${(e.type || 'part') === 'part' ? `<b>${esc(e.tcName)}</b>` : '<span class="hint">—</span>'}</td>
      <td>${type}</td>
      <td>${e.mode === 'swim' ? '<span class="tag swim">スイミング</span>' : '通常'}</td>
      <td>${(e.type || 'part') === 'part' ? fmtDate(e.hire) + '／週' + e.weekDays + '日' : e.dailyHours + '時間/日'}</td>
      <td>${t}</td>
      <td>
        <button type="button" class="btn mini" data-edit="${e.id}">編集</button>
        <button type="button" class="btn mini danger" data-remove="${e.id}">削除</button>
      </td>
    </tr>`;
  }).join('');
}

/* ============ 5. 有給入力 ============ */

let lastLeaveRows = [];                // 一覧に出ている有給（まとめて削除に使う）
let editingLeaveId = null;             // 編集中の有給のid
let paidSelected = new Set();          // 選択中の日付
let paidCursor = new Date();           // カレンダー表示中の月
let inputType = 'full';

function renderCalendar() {
  const box = document.getElementById('paidCalendar');
  const y = paidCursor.getFullYear(), m = paidCursor.getMonth();
  const startDow = new Date(y, m, 1).getDay();
  const daysInMonth = new Date(y, m + 1, 0).getDate();
  const today = todayStr();

  // 選択中の従業員が既に有給を入れている日には印をつける
  const empId = document.getElementById('inEmp').value;
  const has = new Set(DB.leaves.filter(l => l.empId === empId).map(l => l.date));

  let html = `<div class="cal-header">
      <button type="button" class="cal-nav" data-nav="-1">◀</button>
      <span>${y}年${m + 1}月</span>
      <button type="button" class="cal-nav" data-nav="1">▶</button>
    </div><div class="cal-grid">
      ${['日','月','火','水','木','金','土'].map(w => `<div class="cal-dow">${w}</div>`).join('')}`;
  for (let i = 0; i < startDow; i++) html += '<div></div>';
  for (let d = 1; d <= daysInMonth; d++) {
    const ds = toDateStr(new Date(y, m, d));
    const cls = ['cal-day'];
    if (paidSelected.has(ds)) cls.push('sel');
    if (ds === today) cls.push('today');
    if (has.has(ds)) cls.push('has');
    html += `<div class="${cls.join(' ')}" data-date="${ds}">${d}</div>`;
  }
  html += '</div>';
  if (paidSelected.size) html += `<div class="cal-count">${paidSelected.size}日選択中</div>`;
  box.innerHTML = html;
}

function syncDatesField() {
  document.getElementById('inDates').value = [...paidSelected].sort().join(', ');
}

function collectInputDates() {
  const raw = document.getElementById('inDates').value.trim();
  if (!raw) return [];
  const out = [];
  for (const part of raw.split(',')) {
    const s = part.trim();
    if (!s) continue;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
    out.push(s);
  }
  return [...new Set(out)].sort();
}

// 選択中の従業員の所定時間帯を時間欄に入れる（全日のときだけ）
function applyShiftDefault() {
  const emp = getEmp(document.getElementById('inEmp').value);
  const hint = document.getElementById('shiftHint');
  if (!emp) { hint.textContent = ''; return; }
  const sh = emp.shift || {};
  document.getElementById('inTypeHint').textContent = (emp.type || 'part') === 'staff'
    ? '正職方式：年度ごとに時間で積み上げて集計します。'
    : 'パート方式：管理簿の日数として数え、勤務時間帯を入れるとTC5へ出せます。';
  hint.textContent = (emp.type || 'part') === 'staff'
    ? `${emp.name} の1日の勤務時間：${emp.dailyHours}時間。時間帯を入れると時間数が自動で入ります。`
    : sh.in1
      ? `${emp.name} の所定：${sh.in1}〜${sh.out1}${sh.in2 ? ` / ${sh.in2}〜${sh.out2}` : ''}（1日 ${emp.dailyHours}時間）`
      : `${emp.name} は所定の勤務時間帯が未設定です。従業員マスタで登録すると自動で入ります。`;

  if (inputType === 'full') {
    document.getElementById('inTime1a').value = sh.in1 || '';
    document.getElementById('inTime1b').value = sh.out1 || '';
    document.getElementById('inTime2a').value = sh.in2 || '';
    document.getElementById('inTime2b').value = sh.out2 || '';
  }
}

// 入力欄に書かれた時間数（"2:30" でも "2.5" でもよい）
function inputHours() {
  return parseHM(document.getElementById('inHours').value);
}

// "08:00" と "10:30" の間の分数（日をまたぐときは翌日扱い）
function minutesBetween(from, to) {
  if (!from || !to) return 0;
  const [h1, m1] = from.split(':').map(Number);
  const [h2, m2] = to.split(':').map(Number);
  let d = (h2 * 60 + m2) - (h1 * 60 + m1);
  if (d < 0) d += 24 * 60;
  return d;
}

// 入力欄の時間帯①②を足した時間数
function timesTotalHours() {
  const v = id => document.getElementById(id).value;
  return (minutesBetween(v('inTime1a'), v('inTime1b'))
        + minutesBetween(v('inTime2a'), v('inTime2b'))) / 60;
}

// 時間帯を入れたら時間数の欄を自動で埋める（時間単位のときだけ）
function syncHoursFromTimes() {
  if (inputType !== 'hours') return;
  const t = timesTotalHours();
  if (t > 0) document.getElementById('inHours').value = fmtHM(t);
  renderBalancePreview();
}

// 選んだ人の方式に合わせて入力欄の出し方を変える
//   正職方式 … 正職ツールと同じで「時間」だけを入れる。全日／半休の区別はない
//   パート方式 … 管理簿と同じで 全日／半休／時間単位 から選ぶ
function applyInputMode() {
  const emp = getEmp(document.getElementById('inEmp').value);
  const staff = emp && (emp.type || 'part') === 'staff';
  renderInputTerms();
  document.getElementById('inTypeField').classList.toggle('hidden', !!staff);
  // 時間帯はどちらの方式でも使う（正職は時間の計算用、パートはTC5へ渡す用）
  document.getElementById('inShiftLabel').textContent = staff
    ? '有給を取った時間帯（何時から何時まで）'
    : 'TC5に出す勤務時間帯';
  if (staff && inputType !== 'hours') setInputType('hours');
  else setInputType(inputType);
}

function setInputType(type) {
  inputType = type;
  const emp = getEmp(document.getElementById('inEmp').value);
  const staff = emp && (emp.type || 'part') === 'staff';
  const daily = Number(emp && emp.dailyHours) || 8;

  document.querySelectorAll('#inType button').forEach(b =>
    b.classList.toggle('active', b.dataset.type === type));
  document.getElementById('hoursField').classList.toggle('hidden', type !== 'hours');

  if (type === 'hours') {
    document.getElementById('hoursLabel').textContent = staff
      ? '有給として取得した時間' : '有給として扱う時間数';
    // 正職はまる1日休むことも時間で入れる（正職ツールと同じ）ので、目安ボタンを出す
    document.getElementById('hoursQuick').innerHTML = staff
      ? `<button type="button" data-h="${fmtHM(daily)}">1日（${fmtHM(daily)}）</button>
         <button type="button" data-h="${fmtHM(daily / 2)}">半日（${fmtHM(daily / 2)}）</button>
         <button type="button" data-h="1:00">1時間</button>
         <button type="button" data-h="0:30">30分</button>`
      : '';
    document.getElementById('hoursHint').textContent = staff
      ? `上の時間帯を入れると自動で入ります。直接「2:30」「2.5」と書いてもOK。年度の合計を1日${daily}時間で日数に換算します。`
      : `上の時間帯を入れると自動で入ります。1日${daily}時間で日数に換算します（時間単位年休は年5日分が上限）。`;
  }

  if (type === 'full') {
    applyShiftDefault();
  } else {
    // 半日・時間単位は「有給として扱う時間帯」を手で入れてもらう
    document.getElementById('inTime2a').value = '';
    document.getElementById('inTime2b').value = '';
  }
  renderBalancePreview();
}

// 入力中の日付があればその年度、無ければ今日の年度で見る
function dates0() {
  const d = collectInputDates();
  return (d && d.length) ? d[0] : todayStr();
}

/* 有給入力タブでも長期休暇を入れられるようにする（正職の人だけ） */
function renderInputTerms() {
  const field = document.getElementById('inTermField');
  const emp = getEmp(document.getElementById('inEmp').value);
  const staff = emp && (emp.type || 'part') === 'staff';
  field.classList.toggle('hidden', !staff);
  if (!staff) return;

  const year = fiscalYearOf(dates0());
  const y = computeStaffYear(emp, year);
  const cells = y.terms.map(t => `
    <div class="term-box">
      <div class="term-name">${t.label}<span class="hint">最大${t.max}日</span></div>
      <label class="term-in">${termModeLabel()}
        <input type="number" data-term="${t.key}" step="0.5" min="0" placeholder="—"
               value="${t.work == null ? '' : t.work}">
      </label>
      <div class="term-out">取得 <b>${round2(t.taken)}</b> 日</div>
      <div class="term-out${t.remain <= 0 ? ' neg' : ''}">あと <b>${round2(t.remain)}</b> 日</div>
    </div>`).join('');

  document.getElementById('inTermArea').innerHTML = `
    <div class="term-row">${cells}</div>
    <div class="hint">
      ${year}年度（${year}/4〜${year + 1}/3）／ 取得計 <b>${round2(y.longSum)}日</b> ／
      あとどれだけとれる <b>${round2(y.longRemain)}日</b>。
      ${termMode() === 'absent'
        ? '<b>欠勤数</b>＝その休みでもう休んだ日数。あと何日とれるかは「最大日数 − 欠勤数」。'
        : '<b>出勤数</b>＝その休みに出勤した日数。取得日数は「最大日数 − 出勤数」。'}
      入れた内容は「📈 正職」タブと同じもので、呼び方は正職タブで切り替えられます。
    </div>`;
}

function renderBalancePreview() {
  const box = document.getElementById('balancePreview');
  const emp = getEmp(document.getElementById('inEmp').value);
  if (!emp) { box.innerHTML = ''; return; }

  if ((emp.type || 'part') === 'staff') {
    const year = fiscalYearOf(dates0());
    const y = computeStaffYear(emp, year);
    const dts = collectInputDates() || [];
    const add = dts.length * (inputType === 'hours'
      ? inputHours()
      : (inputType === 'full' ? y.daily : y.daily / 2));
    let h = `${year}年度の取得 <b>${y.count}回</b>／<b>${toDaysPlus(y.totalHours, y.daily)}</b>（${round2(y.totalHours)}時間）`;
    if (add > 0) h += ` → 今回 ${round2(add)}時間 を足すと <b>${toDaysPlus(y.totalHours + add, y.daily)}</b>`;
    if (y.longSum || y.longRemain) {
      h += `<br>長期休暇：取得 ${round2(y.longSum)}日 ／ あと ${round2(y.longRemain)}日とれます`;
    }
    box.innerHTML = h;
    return;
  }

  const led = computeLedger(emp, todayStr());
  const dates = collectInputDates() || [];
  const per = inputType === 'hours'
    ? inputHours() / (emp.dailyHours || 8)
    : (inputType === 'full' ? 1 : 0.5);
  const willUse = dates.length * per;
  const after = led.balance - willUse;

  let html = `現在の残 <b>${fmtDays(led.balance, emp.dailyHours)}</b>`;
  if (willUse > 0) {
    html += ` → 今回 ${fmtDays(willUse, emp.dailyHours)} を引くと `
          + `<b class="${after < -1e-9 ? 'neg' : ''}">${fmtDays(after, emp.dailyHours)}</b>`;
    if (after < -1e-9) html += ' <span class="neg">※残が足りません</span>';
  }
  if (led.duty && led.duty.remain > 0) {
    html += `<br>年5日の取得義務：あと <b>${led.duty.remain}日</b>（${fmtDate(led.duty.deadline)} まで）`;
  } else if (led.duty) {
    html += '<br>年5日の取得義務：<span class="ok-text">達成済み</span>';
  }
  if (led.next) html += `<br>次回付与：${fmtDate(led.next.from)}（${led.next.grant}日）`;
  box.innerHTML = html;
}

/* 登録済みの1件をフォームに読み込んで編集する */
function startEditLeave(id) {
  const l = DB.leaves.find(x => x.id === id);
  const emp = l && getEmp(l.empId);
  if (!emp) return;

  editingLeaveId = id;
  document.getElementById('inEmp').value = emp.id;
  applyShiftDefault();
  applyInputMode();
  setInputType(l.type);

  paidSelected = new Set([l.date]);
  syncDatesField();
  document.getElementById('inHours').value = l.type === 'hours' ? leaveHoursText(l) : '';
  document.getElementById('inTime1a').value = l.in1 || '';
  document.getElementById('inTime1b').value = l.out1 || '';
  document.getElementById('inTime2a').value = l.in2 || '';
  document.getElementById('inTime2b').value = l.out2 || '';
  document.getElementById('inNote').value = l.note || '';

  document.getElementById('inSubmit').textContent = 'この内容で更新';
  document.getElementById('inCancelEdit').classList.remove('hidden');
  renderCalendar();
  renderBalancePreview();
  renderInputTerms();
  renderLeaveTable();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function cancelEditLeave(keepForm) {
  editingLeaveId = null;
  document.getElementById('inSubmit').textContent = 'この内容で登録';
  document.getElementById('inCancelEdit').classList.add('hidden');
  if (!keepForm) {
    paidSelected.clear();
    syncDatesField();
    document.getElementById('inHours').value = '';
    document.getElementById('inNote').value = '';
    renderCalendar();
    renderBalancePreview();
  }
  renderLeaveTable();
}

function submitLeave() {
  const emp = getEmp(document.getElementById('inEmp').value);
  if (!emp) { showToast('従業員を選んでください', 'warning'); return; }

  const dates = collectInputDates();
  if (dates === null) { showToast('日付は YYYY-MM-DD 形式で入力してください', 'warning'); return; }
  if (!dates.length) { showToast('取得日を選んでください', 'warning'); return; }

  const hours = inputHours();
  if (inputType === 'hours' && hours <= 0) {
    showToast('時間を入力してください（例: 2:30）', 'warning'); return;
  }

  const in1  = document.getElementById('inTime1a').value;
  const out1 = document.getElementById('inTime1b').value;
  const in2  = document.getElementById('inTime2a').value;
  const out2 = document.getElementById('inTime2b').value;
  if ((in1 && !out1) || (!in1 && out1) || (in2 && !out2) || (!in2 && out2)) {
    showToast('時間帯は開始・終了の両方を入れてください', 'warning'); return;
  }
  if (!in1 && (emp.type || 'part') === 'part') {
    if (!confirm('勤務時間帯が空です。この状態だとTC5用JSONには出力されません（有給の日数管理だけになります）。このまま登録しますか？')) return;
  }

  // 編集中なら、その1件を書き換える
  if (editingLeaveId) {
    const l = DB.leaves.find(x => x.id === editingLeaveId);
    if (!l) { cancelEditLeave(); return; }
    if (dates.length > 1) { showToast('編集は1日ずつです。日付を1つだけ選んでください', 'warning'); return; }
    Object.assign(l, {
      empId: emp.id, date: dates[0], type: inputType,
      hours: inputType === 'hours' ? hours : 0,
      hoursText: inputType === 'hours' ? fmtHM(hours) : '',
      in1, out1, in2, out2,
      note: document.getElementById('inNote').value.trim()
    });
    saveDB();
    cancelEditLeave();
    renderAll();
    showToast(`${emp.name} ${fmtDate(dates[0])} を更新しました`, 'success');
    return;
  }

  let added = 0, skipped = 0;
  for (const date of dates) {
    const dup = DB.leaves.some(l => l.empId === emp.id && l.date === date && l.type === inputType);
    if (dup) { skipped++; continue; }
    DB.leaves.push({
      id: newId(), empId: emp.id, date, type: inputType,
      hours: inputType === 'hours' ? hours : 0,
      hoursText: inputType === 'hours' ? fmtHM(hours) : '',
      in1, out1, in2, out2,
      note: document.getElementById('inNote').value.trim(),
      source: 'manual'
    });
    added++;
  }
  saveDB();

  paidSelected.clear();
  syncDatesField();
  document.getElementById('inNote').value = '';
  renderAll();
  showToast(skipped
    ? `${emp.name}：${added}日を登録（同じ内容の${skipped}日は除外）`
    : `${emp.name}：${added}日を登録しました`, 'success');
}

function renderLeaveTable() {
  const tbody = document.querySelector('#leaveTable tbody');
  const q = document.getElementById('leaveSearch').value.trim();
  const month = document.getElementById('leaveMonth').value;

  let rows = DB.leaves.map(l => ({ l, emp: getEmp(l.empId) })).filter(x => x.emp);
  if (q) rows = rows.filter(x => x.emp.name.includes(q) || x.emp.tcName.includes(q));
  if (month) rows = rows.filter(x => x.l.date.slice(0, 7) === month);
  rows.sort((a, b) => a.l.date < b.l.date ? 1 : a.l.date > b.l.date ? -1 : 0);

  lastLeaveRows = rows;
  const la = document.getElementById('leaveCheckAll'); if (la) la.checked = false;
  document.getElementById('leaveCount').textContent = `${rows.length}件`
    + (rows.length > 500 ? '（表示は500件まで）' : '');
  if (!rows.length) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="7">該当する有給がありません</td></tr>';
    return;
  }
  tbody.innerHTML = rows.slice(0, 500).map(({ l, emp }) => {
    const editing = l.id === editingLeaveId;
    const cls = l.type === 'full' ? 'full' : l.type === 'hours' ? 'hours' : 'half';
    const time = l.in1 ? `${l.in1}〜${l.out1}` + (l.in2 ? ` / ${l.in2}〜${l.out2}` : '')
      : (emp.type || 'part') === 'part'
        ? '<span class="tag warn">時間なし</span>'      // パートはTC5へ出すので時間帯が要る
        : '<span class="hint">—</span>';
    return `<tr${editing ? ' class="editing"' : ''}>
      <td><input type="checkbox" class="leaveChk" data-id="${l.id}"></td>
      <td>${fmtDate(l.date)}</td>
      <td>${esc(emp.name)}${emp.mode === 'swim' ? ' <span class="tag swim">SW</span>' : ''}</td>
      <td><span class="tag ${cls}">${TYPE_LABEL[l.type]}${l.type === 'hours' ? ` ${leaveHoursText(l)}` : ''}</span></td>
      <td class="num">${fmtDays(leaveDays(l, emp), emp.dailyHours)}</td>
      <td>${time}</td>
      <td>
        <button type="button" class="btn mini" data-edit="${l.id}">編集</button>
        <button type="button" class="btn mini danger" data-del="${l.id}">削除</button>
      </td>
    </tr>`;
  }).join('');
}

function refreshMonthFilter() {
  const sel = document.getElementById('leaveMonth');
  const cur = sel.value;
  const months = [...new Set(DB.leaves.map(l => l.date.slice(0, 7)))].sort().reverse();
  sel.innerHTML = '<option value="">すべての月</option>' +
    months.map(m => `<option value="${m}">${m.replace('-', '年')}月</option>`).join('');
  if (months.includes(cur)) sel.value = cur;
}

function refreshEmpSelects() {
  const all = sortedEmployees();
  const opt = e =>
    `<option value="${e.id}">${esc(e.name)}${e.mode === 'swim' ? '（スイミング）' : ''}${e.active === false ? '（退職）' : ''}</option>`;
  // 入力は在籍者だけ、管理簿は退職者も見られるようにする
  const sets = {
    inEmp:   all.filter(e => e.active !== false),
    bookEmp: all.filter(e => (e.type || 'part') === 'part')
  };
  for (const id in sets) {
    const sel = document.getElementById(id);
    const cur = sel.value;
    sel.innerHTML = sets[id].map(opt).join('') || '<option value="">（従業員が未登録です）</option>';
    if (sets[id].some(e => e.id === cur)) sel.value = cur;
  }
  document.getElementById('tcNameList').innerHTML =
    tcNames().map(n => `<option value="${esc(n)}"></option>`).join('');
}

/* ============ 6. 残日数一覧・管理簿 ============ */

function renderLedgerTable() {
  const tbody = document.querySelector('#ledgerTable tbody');
  const list = sortedEmployees().filter(e => (e.type || 'part') === 'part');
  if (!list.length) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="12">パート方式の従業員が登録されていません</td></tr>';
    return;
  }
  const asOf = todayStr();
  tbody.innerHTML = list.map(e => {
    const g = computeLedger(e, asOf);
    const duty = !g.duty ? '—'
      : g.duty.remain <= 0 ? '<span class="ok-text">達成</span>'
      : `<span class="danger-text">あと${g.duty.remain}日</span><br><span class="hint">${fmtDate(g.duty.deadline)}まで</span>`;
    return `<tr>
      <td>${esc(e.name)}</td>
      <td><b>${esc(e.tcName)}</b>${e.mode === 'swim' ? ' <span class="tag swim">SW</span>' : ''}</td>
      <td>${fmtDate(e.hire)}</td>
      <td class="num">${e.weekDays}</td>
      <td>${g.current ? fmtDate(g.current.from) : '未付与'}</td>
      <td class="num">${g.current ? g.grantDays : '—'}</td>
      <td class="num">${g.current ? g.carry : '—'}</td>
      <td class="num">${g.current ? g.grantDays + g.carry : '—'}</td>
      <td class="num">${fmtDays(g.takenInPeriod, e.dailyHours)}</td>
      <td class="num"><b class="${g.balance < -1e-9 ? 'danger-text' : ''}">${fmtDays(g.balance, e.dailyHours)}</b></td>
      <td>${duty}</td>
      <td>${g.next ? `${fmtDate(g.next.from)}<br><span class="hint">${g.next.grant}日</span>` : '—'}</td>
    </tr>`;
  }).join('');
}

/* --- 管理簿（見本Excelと同じ体裁） --- */

// 取得日セルの中身
function takesCell(row, emp) {
  if (!row.takes.length) return '<span class="hint">—</span>';
  return row.takes.map(l => {
    const d = leaveDays(l, emp);
    const cls = l.type === 'full' ? 'full' : l.type === 'hours' ? 'hours' : 'half';
    const sub = l.type === 'full' ? '' : `<sub>${l.type === 'hours' ? leaveHoursText(l) : '半'}</sub>`;
    return `<span class="take ${cls}" title="${TYPE_LABEL[l.type]}（${d}日）">${fmtDate(l.date).slice(5)}${sub}</span>`;
  }).join('');
}

function renderBook() {
  const box = document.getElementById('bookArea');
  const emp = getEmp(document.getElementById('bookEmp').value);
  if (!emp) {
    box.innerHTML = '<div class="hint">パート方式の従業員がいません。管理簿はパート方式（入社日基準）の人だけが対象です。'
      + '正職の人は「📈 正職集計」タブを見てください。</div>';
    return;
  }

  const asOf = todayStr();
  const g = computeLedger(emp, asOf);
  const shown = g.rows.filter(r => r.started || r === g.next);

  const body = shown.map(r => {
    const future = !r.started;
    const duty = r.duty === null ? '' : (r.duty <= 0 ? '<span class="ok-text">終了</span>'
                 : `<span class="${r.current ? 'danger-text' : ''}">${r.duty}</span>`);
    return `<tr class="${r.current ? 'book-current' : ''}${future ? ' book-future' : ''}">
      <td class="num">${r.months}</td>
      <td>${fmtDate(r.from)}</td>
      <td><input type="number" class="wd" data-week="${r.from}" min="1" max="5" step="1" value="${r.weekDays}" title="この回の付与に使う週の所定労働日数"></td>
      <td><input type="checkbox" data-invalid="${r.from}"${r.invalid ? ' checked' : ''} title="チェックすると付与0日（無効）"></td>
      <td class="num">${r.grant}${r.invalid ? `<span class="hint">(${r.base})</span>` : ''}</td>
      <td class="num">${r.carry}</td>
      <td class="num">${r.total}</td>
      <td class="num">${fmtDays(r.taken, emp.dailyHours)}</td>
      <td class="num">${duty}</td>
      <td class="num"><b class="${r.remain < -1e-9 ? 'danger-text' : ''}">${fmtDays(r.remain, emp.dailyHours)}</b></td>
      <td>${fmtDate(r.from)} 〜 ${fmtDate(r.to)}</td>
      <td class="takes">${takesCell(r, emp)}</td>
    </tr>`;
  }).join('') || '<tr class="empty-row"><td colspan="12">入社日から6ヶ月未満のため、まだ付与がありません</td></tr>';

  let warn = '';
  if (g.shortages.length) {
    warn += `<div class="balance-preview"><span class="neg">⚠ 残日数を超えて取得している期があります：</span>`
          + g.shortages.map(s => `${fmtDate(s.from)}の期（${fmtDays(s.days, emp.dailyHours)}超過）`).join('、') + '</div>';
  }
  if (g.hourlyOver) {
    warn += `<div class="balance-preview"><span class="neg">⚠ 時間単位年休が年5日ぶんを超えています</span>（${fmtDays(g.hourlyOver, emp.dailyHours)}）</div>`;
  }

  box.innerHTML = `
    <div id="bookPrint">
      <div class="book-title">年次有給休暇管理簿</div>
      <table class="book-head"><tbody>
        <tr><th>ふりがな</th><td>${esc(emp.kana || '')}</td>
            <th>入社日</th><td>${fmtDate(emp.hire)}</td></tr>
        <tr><th>氏　名</th><td class="book-name">${esc(emp.name)}</td>
            <th>勤続年数</th><td>${serviceLength(emp.hire, asOf)}</td></tr>
        <tr><th>週労働日数</th><td>${emp.weekDays}日<span class="hint">（既定）</span></td>
            <th>1日の所定</th><td>${emp.dailyHours}時間</td></tr>
      </tbody></table>
      ${warn}
      <div class="table-wrap">
        <table class="book-table"><thead><tr>
          <th>勤続<br>月数</th><th>有給休暇<br>発生日</th><th>週労働<br>日数</th><th>無効</th>
          <th>付与</th><th>前年<br>繰越</th><th>計</th><th>取得</th>
          <th>義務残</th><th>残日数</th><th>対　象　期　間</th><th>取　得　日</th>
        </tr></thead><tbody>${body}</tbody></table>
      </div>
      <div class="hint">
        前年繰越＝前の期の残り（古い付与から先に使い、使い残しは2年で失効）。
        義務残＝年5日の取得義務のうち残っている日数（時間単位年休は含みません）。
        <span class="take full">05/12</span> のような表示が取得日で、
        <span class="take half">05/13<sub>半</sub></span> は半休、
        <span class="take hours">05/14<sub>4h</sub></span> は時間単位です。
      </div>
    </div>`;
}

/* --- 管理簿をExcelで書き出す（見本と同じ列の並び） --- */

function buildBookSheet(emp, asOf) {
  const g = computeLedger(emp, asOf);
  const A = [];
  const put = (r, c, v) => { (A[r] = A[r] || [])[c] = v; };

  // 右上の付与日数表（見本の N1:U7）
  const heads = ['6ヶ月', '1年6ヶ月', '2年6ヶ月', '3年6ヶ月', '4年6ヶ月', '5年6ヶ月', '6年6ヶ月'];
  put(0, 13, '週間労働日数');
  heads.forEach((h, i) => put(0, 14 + i, h));
  for (let w = 1; w <= 6; w++) {
    const line = GRANT_TABLE[Math.min(w, 5)];
    put(w, 13, w);
    line.forEach((v, i) => put(w, 14 + i, v));
  }

  put(0, 1, '年次有給休暇管理簿');
  put(1, 1, 'ふりがな');   put(1, 2, emp.kana || '');
  put(2, 1, '氏　名');     put(2, 2, emp.name);
  put(4, 1, '入社日');     put(4, 2, fmtDate(emp.hire));
  put(5, 1, '勤続年数');   put(5, 2, serviceLength(emp.hire, asOf));

  const H = ['勤続月数', '有給休暇発生日', '週労働日数', '無効', '付与', '前年　繰越',
             '計', '取得', '義務残', '残日数', '対　象　期　間'];
  H.forEach((h, i) => put(7, i, h));
  put(7, 19, '取　　得　　日');

  const merges = [{ s: { r: 7, c: 10 }, e: { r: 7, c: 12 } }];
  const shown = g.rows.filter(r => r.started || r === g.next);

  shown.forEach((r, i) => {
    const R = 8 + i * 2;
    put(R, 0, r.months);
    put(R, 1, fmtDate(r.from));
    put(R, 2, r.weekDays);
    put(R, 3, r.invalid ? '×' : '');
    put(R, 4, r.grant);
    put(R, 5, r.carry);
    put(R, 6, r.total);
    put(R, 7, Math.round(r.taken * 100) / 100);
    put(R, 8, r.duty === null ? '' : (r.duty <= 0 ? '終了' : r.duty));
    put(R, 9, Math.round(r.remain * 100) / 100);
    put(R, 10, fmtDate(r.from));
    put(R, 11, '～');
    put(R, 12, fmtDate(r.to));
    r.takes.slice(0, 20).forEach((l, k) => {
      put(R, 13 + k, fmtDate(l.date));
      put(R + 1, 13 + k, Math.round(leaveDays(l, emp) * 1000) / 1000);
    });
    for (let c = 0; c <= 12; c++) merges.push({ s: { r: R, c }, e: { r: R + 1, c } });
  });

  // 抜けているセルを '' で埋める（列ずれ防止）
  const maxR = Math.max(8, 7 + shown.length * 2);
  for (let r = 0; r <= maxR; r++) {
    A[r] = A[r] || [];
    for (let c = 0; c <= 32; c++) if (A[r][c] === undefined) A[r][c] = '';
  }

  const ws = XLSX.utils.aoa_to_sheet(A);
  ws['!merges'] = merges;
  ws['!cols'] = [{ wch: 6 }, { wch: 13 }, { wch: 7 }, { wch: 5 }, { wch: 6 }, { wch: 7 },
                 { wch: 6 }, { wch: 6 }, { wch: 7 }, { wch: 7 }, { wch: 12 }, { wch: 4 }, { wch: 12 }]
                 .concat(Array(20).fill({ wch: 8 }));
  return ws;
}

// Excelのシート名に使えない文字を落とし、31文字に収める
function safeSheetName(name, used) {
  const base = String(name).replace(/[\\\/\?\*\[\]:]/g, '').slice(0, 28) || 'sheet';
  let n = base, i = 2;
  while (used.has(n)) n = `${base}(${i++})`;
  used.add(n);
  return n;
}

function exportBookXlsx(scope) {
  const asOf = todayStr();
  let list;
  if (scope === 'one') {
    const emp = getEmp(document.getElementById('bookEmp').value);
    if (!emp) { showToast('従業員を選んでください', 'warning'); return; }
    list = [emp];
  } else {
    list = sortedEmployees().filter(e => (e.type || 'part') === 'part');
  }
  if (!list.length) { showToast('パート方式の従業員が登録されていません', 'warning'); return; }

  const wb = XLSX.utils.book_new();
  const used = new Set();
  for (const emp of list) {
    XLSX.utils.book_append_sheet(wb, buildBookSheet(emp, asOf), safeSheetName(emp.name, used));
  }
  const stamp = asOf.replace(/-/g, '');
  XLSX.writeFile(wb, list.length === 1
    ? `年次有給休暇管理簿_${list[0].name}_${stamp}.xlsx`
    : `年次有給休暇管理簿_全員_${stamp}.xlsx`);
  showToast(`管理簿を出力しました（${list.length}人）`, 'success');
}
/* ============ 6.5 正職集計（年度・時間の方式） ============ */

function refreshStaffYearSelect() {
  const sel = document.getElementById('staffYear');
  const cur = sel.value;
  const years = knownFiscalYears();
  sel.innerHTML = years.map(y => `<option value="${y}">${y}年度（${y}/4〜${y + 1}/3）</option>`).join('');
  if (years.map(String).includes(cur)) sel.value = cur;
  else sel.value = String(fiscalYearOf(todayStr()));
}

function currentStaffYear() {
  return Number(document.getElementById('staffYear').value) || fiscalYearOf(todayStr());
}

/* 長期休暇の設定（最大日数と呼び方）。人ごとの入力は「✍️ 有給入力」タブで行う */
function renderTermInput() {
  const year = currentStaffYear();
  const f = getFiscal(year);
  const mode = termMode();

  const maxBar = TERM_FIELDS.map(t =>
    `<label>${t.label} <input type="number" class="mx" data-max="${t.max}" step="0.5" min="0" value="${f.maxLeave[t.max]}"> 日</label>`
  ).join('');

  document.getElementById('termArea').innerHTML = `
    <div class="max-bar">
      <strong>${year}年度の最大取得日数</strong>${maxBar}
    </div>
    <div class="field" style="margin-bottom:6px">
      <label>入力のしかた</label>
      <div class="seg" id="termModeSeg">
        <button type="button" data-mode="work"${mode === 'work' ? ' class="active"' : ''}>出勤数で入力</button>
        <button type="button" data-mode="absent"${mode === 'absent' ? ' class="active"' : ''}>欠勤数で入力</button>
      </div>
    </div>
    <div class="hint">
      ${mode === 'absent'
        ? '<b>欠勤数</b>（その長期休暇で<b>もう休んだ日数</b>）を入れると、'
          + '<b>取得＝欠勤数</b>、<b>残り＝最大日数−欠勤数</b>（あと何日とれるか）になります。'
        : '<b>出勤数</b>（その長期休暇中に<b>出勤した日数</b>）を入れると、'
          + '<b>取得＝最大日数−出勤数</b>、<b>残り＝出勤数</b>になります。'}<br>
      <b>人ごとの入力は「✍️ 有給入力」タブ</b>で、その人を選ぶと出てきます。
      結果は下の集計表に出ます。<br>
      <b>⚠ 入力のしかたを切り替えると、すでに入っている数字の意味が入れ替わります</b>
      （出勤数として入れた3日は、欠勤数では「3日休んだ」になります）。
    </div>`;
}

/* 📅 月別表（正職ツールと同じ形。ここは見るだけで、入力は有給入力タブ） */
let staffGridMonth = new Date().getMonth() + 1;

function renderStaffGrid() {
  const box = document.getElementById('gridArea');
  const year = currentStaffYear();
  const month = staffGridMonth;
  const actual = month >= 4 ? year : year + 1;          // 1〜3月は翌年
  const days = new Date(actual, month, 0).getDate();
  const list = sortedEmployees().filter(e => (e.type || 'part') === 'staff');

  const nav = `<div class="grid-nav no-print">
      <button type="button" class="cal-nav" data-gmonth="-1">‹</button>
      <select id="gridMonth">${FISCAL_MONTHS.map(m =>
        `<option value="${m}"${m === month ? ' selected' : ''}>${m}月</option>`).join('')}</select>
      <button type="button" class="cal-nav" data-gmonth="1">›</button>
      <span class="hint">${year}年度（${actual}年${month}月）</span>
    </div>`;

  if (!list.length) {
    box.innerHTML = nav + '<div class="hint">正職方式の従業員がいません。</div>';
    return;
  }

  // その日に入っている有給の合計時間
  const hoursOn = (emp, dateStr) => leavesOf(emp.id)
    .filter(l => l.date === dateStr)
    .reduce((s, l) => s + leaveHours(l, emp), 0);

  let head = '<th class="name-col">名前＼日</th>';
  for (let d = 1; d <= days; d++) {
    const dw = new Date(actual, month - 1, d).getDay();
    head += `<th class="${dw === 0 ? 'sun' : dw === 6 ? 'sat' : ''}">${d}<br><span class="dow">${DOW[dw]}</span></th>`;
  }
  head += '<th class="total">合計<br><span class="dow">時:分</span></th>';

  const body = list.map(e => {
    let row = `<td class="name-col">${esc(e.name)}</td>`;
    let sum = 0;
    for (let d = 1; d <= days; d++) {
      const ds = `${actual}-${pad2(month)}-${pad2(d)}`;
      const h = hoursOn(e, ds);
      sum += h;
      const dw = new Date(actual, month - 1, d).getDay();
      row += `<td class="${dw === 0 ? 'sun' : dw === 6 ? 'sat' : ''}${h ? ' has' : ''}">${h ? fmtHM(h) : ''}</td>`;
    }
    row += `<td class="total">${sum ? fmtHM(sum) : ''}</td>`;
    return `<tr>${row}</tr>`;
  }).join('');

  box.innerHTML = nav + `
    <div id="gridPrint">
      <div class="book-title no-screen">${year}年度　有給管理表（${month}月）</div>
      <div class="table-wrap">
        <table class="grid"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>
      </div>
      <div class="hint">数字はその日に入っている有給の時間（時:分）。入力は「✍️ 有給入力」タブから。</div>
    </div>`;
}

function renderStaffSummary() {
  const year = currentStaffYear();
  const f = getFiscal(year);
  const list = sortedEmployees().filter(e => (e.type || 'part') === 'staff');

  if (!list.length) {
    document.getElementById('staffArea').innerHTML = '';
    return;
  }

  let sumHours = 0, sumDays = 0, sumRest = 0, sumLong = [0, 0, 0], sumLongAll = 0;
  let sumCount = 0, sumLongRemain = 0;
  const rows = list.map(e => {
    const y = computeStaffYear(e, year);
    sumHours += y.totalHours; sumDays += y.days; sumRest += y.rest; sumLongAll += y.longSum;
    sumCount += y.count; sumLongRemain += y.longRemain;
    y.terms.forEach((t, i) => { sumLong[i] += t.taken; });
    const termCells = y.terms.map(t =>
      `<td class="num">${t.has ? t.work : '—'}</td>
       <td class="num">${round2(t.taken)} 日</td>
       <td class="num">${round2(t.remain)} 日</td>`).join('');
    return `<tr>
      <td>${esc(e.name)}${e.mode === 'swim' ? ' <span class="tag swim">SW</span>' : ''}</td>
      <td class="num">${y.count} 回</td>
      <td class="num">${round2(y.totalHours)} 時間</td>
      <td class="num">${y.daily} 時間</td>
      <td class="num"><b>${toDaysPlus(y.totalHours, y.daily)}</b></td>
      ${termCells}
      <td class="num">${round2(y.longSum)} 日</td>
      <td class="num">${round2(y.longRemain)} 日</td>
      <td class="num grand"><b>${round2(y.grandDays)}日＋${round2(y.grandRest)}時間</b></td>
    </tr>`;
  }).join('');

  document.getElementById('staffArea').innerHTML = `
    <div id="staffPrint">
      <div class="book-title">${year}年度　有給取得 集計表</div>
      <div class="table-wrap">
        <table class="sum-table"><thead>
          <tr>
            <th rowspan="2">名前</th>
            <th colspan="4">通常有給</th>
            <th colspan="3">夏期休暇<br>(最大${f.maxLeave.summer}日)</th>
            <th colspan="3">冬期休暇<br>(最大${f.maxLeave.winter}日)</th>
            <th colspan="3">春期休暇<br>(最大${f.maxLeave.spring}日)</th>
            <th rowspan="2">長期休暇<br>取得計</th>
            <th rowspan="2">長期休暇<br>あとどれだけ</th>
            <th rowspan="2">総合計<br>(通常+長期)</th>
          </tr>
          <tr>
            <th>取得回数</th><th>合計取得時間</th><th>1日の勤務時間</th><th>有給取得日数</th>
            <th>${termModeLabel()}</th><th>取得</th><th>残り</th>
            <th>${termModeLabel()}</th><th>取得</th><th>残り</th>
            <th>${termModeLabel()}</th><th>取得</th><th>残り</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
        <tfoot><tr>
          <th>合計</th>
          <th class="num">${sumCount} 回</th>
          <th class="num">${round2(sumHours)} 時間</th><th></th>
          <th class="num">${sumDays}日＋${round2(sumRest)}時間</th>
          <th></th><th class="num">${round2(sumLong[0])} 日</th><th></th>
          <th></th><th class="num">${round2(sumLong[1])} 日</th><th></th>
          <th></th><th class="num">${round2(sumLong[2])} 日</th><th></th>
          <th class="num">${round2(sumLongAll)} 日</th>
          <th class="num">${round2(sumLongRemain)} 日</th><th></th>
        </tr></tfoot>
        </table>
      </div>
      <div class="hint">
        ・「取得回数」＝ その年度に有給を入れた件数。「有給取得日数」＝ 合計取得時間 ÷ 1日の勤務時間。<br>
        ・長期休暇の「残り」＝ あと何日とれるか。<b>入力は「✍️ 有給入力」タブで</b>。<br>
        ・「1日の勤務時間」は「👤 従業員情報」で各自に合わせて変更できます。
      </div>
    </div>`;
}

function exportStaffCsv() {
  const year = currentStaffYear();
  const f = getFiscal(year);
  const list = sortedEmployees().filter(e => (e.type || 'part') === 'staff');
  if (!list.length) { showToast('正職方式の従業員がいません', 'warning'); return; }

  const ml = termModeLabel();
  const rows = [['名前', '取得回数', '合計取得時間', '1日の勤務時間', '有給取得日数',
    `夏期${ml}`, `夏期取得(最大${f.maxLeave.summer})`, '夏期残り',
    `冬期${ml}`, `冬期取得(最大${f.maxLeave.winter})`, '冬期残り',
    `春期${ml}`, `春期取得(最大${f.maxLeave.spring})`, '春期残り',
    '長期休暇取得計', '長期休暇残り計']];
  for (const e of list) {
    const y = computeStaffYear(e, year);
    rows.push([e.name, y.count, round2(y.totalHours), y.daily, toDaysPlus(y.totalHours, y.daily),
      y.terms[0].has ? y.terms[0].work : '', round2(y.terms[0].taken), round2(y.terms[0].remain),
      y.terms[1].has ? y.terms[1].work : '', round2(y.terms[1].taken), round2(y.terms[1].remain),
      y.terms[2].has ? y.terms[2].work : '', round2(y.terms[2].taken), round2(y.terms[2].remain),
      round2(y.longSum), round2(y.longRemain)]);
  }
  const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\r\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv' });   // Excelで開けるようBOM付き
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `有給集計_${year}年度.csv`; a.click();
  URL.revokeObjectURL(url);
  showToast('集計表をCSVで書き出しました', 'success');
}

/* --- 正職ツール（有給管理_YYYY年度.json）の取込 --- */

function importStaffJson(obj, fileName, log) {
  if (!obj || !Array.isArray(obj.staff) || !obj.entries) {
    log.push(`<li class="danger-text">【${esc(fileName)}】正職ツールのJSONではありません</li>`);
    return { empAdded: 0, empUpdated: 0, leaveAdded: 0, skipped: 1 };
  }
  const year = Number(obj.year) || fiscalYearOf(todayStr());
  const f = getFiscal(year);
  if (obj.maxLeave) f.maxLeave = { ...DEFAULT_MAX_LEAVE, ...obj.maxLeave };

  let empAdded = 0, empUpdated = 0, leaveAdded = 0;
  for (const st of obj.staff) {
    let emp = DB.employees.find(e =>
      samePerson(e, { name: st.name, tcName: '', mode: 'normal', hire: '' }));
    if (!emp) {
      emp = {
        id: newId(), name: st.name, kana: '',
        tcName: '',                       // 正職はTC5を使わない
        type: 'staff', mode: 'normal',
        hire: '', weekDays: 5, dailyHours: Number(st.daily) || 8,
        shift: { in1: '', out1: '', in2: '', out2: '' },
        grantOverrides: {}, invalidGrants: {}, active: true
      };
      DB.employees.push(emp);
      empAdded++;
    } else {
      emp.type = 'staff';
      emp.tcName = '';
      if (st.daily) emp.dailyHours = Number(st.daily);
      empUpdated++;
    }

    // 長期休暇の出勤数
    const work = getTermWork(year, emp.id);
    for (const t of TERM_FIELDS) {
      if (st[t.key] !== undefined && st[t.key] !== null) work[t.key] = st[t.key];
    }

    // 有給の記録（時間）
    const ent = obj.entries[st.id] || {};
    for (const date in ent) {
      const hours = parseHM(ent[date]);
      if (!hours) continue;
      if (DB.leaves.some(l => l.empId === emp.id && l.date === date)) continue;
      DB.leaves.push({
        id: newId(), empId: emp.id, date, type: 'hours', hours,
        hoursText: String(ent[date]).trim(),   // 元の "2:30" 表記も残しておく
        in1: '', out1: '', in2: '', out2: '', note: '', source: 'staffjson'
      });
      leaveAdded++;
    }
  }
  log.push(`<li>【${esc(fileName)}】${year}年度：正職 ${obj.staff.length}人、有給 ${leaveAdded}件を取り込みました</li>`);
  // 正職ツールの sWork/wWork/spWork は「出勤数」。欠勤数モードのままだと逆の意味に読まれる
  if (termMode() === 'absent') {
    log.push('<li class="danger-text">⚠ 正職ツールの長期休暇は「出勤数」で入っています。'
      + 'いまの設定は「欠勤数で入力」なので、取り込んだ数字が逆の意味に読まれます。'
      + '「📈 正職」タブで「出勤数で入力」に切り替えてください。</li>');
  }
  return { empAdded, empUpdated, leaveAdded, skipped: 0 };
}

function handleStaffJsonFiles(files) {
  const box = document.getElementById('staffImportResult');
  const log = [];
  const total = { empAdded: 0, empUpdated: 0, leaveAdded: 0, skipped: 0 };
  let done = 0;
  for (const file of files) {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const r = importStaffJson(JSON.parse(ev.target.result), file.name, log);
        for (const k in total) total[k] += r[k];
      } catch (err) {
        log.push(`<li class="danger-text">【${esc(file.name)}】読み込みに失敗しました（${esc(err.message)}）</li>`);
      }
      if (++done === files.length) {
        saveDB();
        refreshStaffYearSelect();
        renderAll();
        box.innerHTML = `<b>取込結果</b>：正職 新規${total.empAdded}人 / 更新${total.empUpdated}人、有給 ${total.leaveAdded}件<ul>${log.join('')}</ul>`
          + '<div class="hint">取込んだ有給には勤務時間帯が入っていません（正職ツールが時間数しか持たないため）。'
          + 'TC5へ出す分だけ、有給入力タブで時間帯を入れ直してください。</div>';
        showToast(`正職ツールのJSONを取り込みました（有給 ${total.leaveAdded}件）`, 'success');
      }
    };
    reader.readAsText(file);
  }
}

/* ============ 7. 取込（既存Excel・正職ツールJSON） ============ */

// Excelの日付シリアル値 → 'YYYY-MM-DD'
function serialToDateStr(n) {
  if (typeof n !== 'number' || !isFinite(n) || n < 1) return null;
  const ms = Math.round((n - 25569) * 86400000);
  const d = new Date(ms);
  if (isNaN(d.getTime())) return null;
  return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
}

function cellVal(ws, addr) {
  const c = ws[addr];
  return c && c.v !== undefined ? c.v : null;
}
function cellNum(ws, addr) {
  const v = cellVal(ws, addr);
  return typeof v === 'number' && isFinite(v) ? v : null;
}
function cellStr(ws, addr) {
  const v = cellVal(ws, addr);
  return v == null ? '' : String(v).trim();
}

/**
 * 「年次有給休暇管理簿」1シートを読む。
 *   氏名 C3／ふりがな C2／入社日 C5
 *   期間行は 9,11,13,…,55（上段）。C=週労働日数、N〜AG＝取得日、その1行下＝取得日数
 * ※ 見本が空だったため、この読み取り位置は見本のレイアウトに合わせてある。
 *    実データで位置が違っていた場合はこの関数だけ直せばよい。
 */
// Excelの取得日セルは「10/25」のように月日しか表示されず、内部の年は
// 入力した年のままになっていることがある。書かれていた行（対象期間）が正なので、
// その期間に収まる年に付け替える。
function fitYearInto(dateStr, from, to) {
  if (!from || !to) return dateStr;
  const md = dateStr.slice(5);                       // 'MM-DD'
  const y0 = Number(from.slice(0, 4));
  for (const y of [y0, y0 + 1]) {
    const cand = `${y}-${md}`;
    if (cand >= from && cand <= to) return cand;
  }
  return dateStr;                                     // どの年にも収まらなければそのまま
}

function parseLedgerSheet(ws, sheetName) {
  // 実ファイルは 姓＝C / 名＝E の2セルに分かれている（見本では結合されていて見えない）
  const join = (a, b) => [cellStr(ws, a), cellStr(ws, b)].filter(Boolean).join('　');
  const name = join('C3', 'E3') || sheetName;
  const kana = join('C2', 'E2');
  const hire = serialToDateStr(cellNum(ws, 'C5'));

  const COLS = [];
  for (let i = 13; i <= 32; i++) COLS.push(XLSX.utils.encode_col(i));   // N〜AG

  let weekDays = null;
  const weekDaysByGrant = {};
  const invalidGrants = {};
  const takes = [];

  for (let r = 9; r <= 55; r += 2) {
    // その行の付与日。B列（数式の計算結果）を優先し、無ければ入社日＋勤続月数から出す
    const months = cellNum(ws, 'A' + r);
    let from = serialToDateStr(cellNum(ws, 'B' + r));
    if (!from && hire && months != null) from = addMonths(hire, months);
    const to = from ? addDays(addMonths(from, 12), -1) : null;

    const c = cellNum(ws, 'C' + r);
    if (c != null && c >= 1 && c <= 6) {
      weekDays = c;                                   // 既定＝最後に見つかった値
      if (from) weekDaysByGrant[from] = c;            // 付与ごとの値
    }
    if (from && cellStr(ws, 'D' + r)) invalidGrants[from] = true;   // 「無効」列

    for (const col of COLS) {
      const ds = serialToDateStr(cellNum(ws, col + r));
      if (!ds) continue;
      let days = cellNum(ws, col + (r + 1));
      if (days == null || days <= 0) days = 1;
      takes.push({ date: fitYearInto(ds, from, to), days, periodFrom: from });
    }
  }
  return { sheetName, name, kana, hire, weekDays, weekDaysByGrant, invalidGrants, takes };
}

function importWorkbook(wb, fileName, log) {
  let empAdded = 0, empUpdated = 0, leaveAdded = 0, skipped = 0;

  for (const sheetName of wb.SheetNames) {
    const parsed = parseLedgerSheet(wb.Sheets[sheetName], sheetName);
    if (!parsed.hire) {
      log.push(`<li>【${esc(fileName)} / ${esc(sheetName)}】入社日（C5）が読めなかったので飛ばしました</li>`);
      skipped++;
      continue;
    }

    // 既存の従業員と突き合わせ（氏名 → TC5表記 → シート名）
    // 氏名の書き方が変わっても同じ人だと分かるようにする
    const cand = {
      name: parsed.name, tcName: '', mode: 'normal', hire: parsed.hire,
      kana: parsed.kana
    };
    let emp = DB.employees.find(e => samePerson(e, cand))
           || DB.employees.find(e => samePerson(e, { ...cand, name: parsed.sheetName }));

    if (!emp) {
      emp = {
        id: newId(),
        name: parsed.name || parsed.sheetName,
        kana: parsed.kana,
        tcName: tcNames().find(n => (parsed.name || '').includes(n) || parsed.sheetName.includes(n)) || '',
        type: 'part',
        mode: 'normal',
        hire: parsed.hire,
        weekDays: parsed.weekDays || 5,
        weekDaysByGrant: { ...parsed.weekDaysByGrant },
        invalidGrants: { ...parsed.invalidGrants },
        dailyHours: 8,
        shift: { in1: '', out1: '', in2: '', out2: '' },
        grantOverrides: {},
        active: true
      };
      DB.employees.push(emp);
      empAdded++;
    } else {
      emp.type = 'part';
      // 「山田」→「山田　花子」のように詳しくなったら氏名を更新する
      if (normName(parsed.name).length > normName(emp.name).length) emp.name = parsed.name;
      emp.hire = parsed.hire;
      if (parsed.weekDays) emp.weekDays = parsed.weekDays;
      emp.weekDaysByGrant = { ...(emp.weekDaysByGrant || {}), ...parsed.weekDaysByGrant };
      emp.invalidGrants = { ...(emp.invalidGrants || {}), ...parsed.invalidGrants };
      if (parsed.kana && !emp.kana) emp.kana = parsed.kana;
      empUpdated++;
    }

    for (const t of parsed.takes) {
      if (DB.leaves.some(l => l.empId === emp.id && l.date === t.date)) continue;
      const type = t.days >= 1 ? 'full' : t.days === 0.5 ? 'am' : 'hours';
      DB.leaves.push({
        id: newId(), empId: emp.id, date: t.date, type,
        hours: type === 'hours' ? Math.round(t.days * (emp.dailyHours || 8) * 100) / 100 : 0,
        in1: '', out1: '', in2: '', out2: '',
        note: '', source: 'excel'
      });
      leaveAdded++;
    }

    const wdList = Object.keys(parsed.weekDaysByGrant).sort()
      .map(d => `${fmtDate(d)}=週${parsed.weekDaysByGrant[d]}日`).join('、');
    log.push(`<li>【${esc(sheetName)}】${esc(emp.name)}：入社日 ${fmtDate(parsed.hire)}／取得 ${parsed.takes.length}件`
      + (wdList ? `<br><span class="hint">週労働日数：${wdList}</span>` : '')
      + (emp.tcName ? '' : ' <span class="danger-text">← TC5での名前が未設定</span>') + '</li>');
  }
  return { empAdded, empUpdated, leaveAdded, skipped };
}

function handleXlsxFiles(files) {
  const box = document.getElementById('importResult');
  const log = [];
  let total = { empAdded: 0, empUpdated: 0, leaveAdded: 0, skipped: 0 };
  let done = 0;

  for (const file of files) {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        const r = importWorkbook(wb, file.name, log);
        for (const k in total) total[k] += r[k];
      } catch (err) {
        log.push(`<li class="danger-text">【${esc(file.name)}】読み込みに失敗しました（${esc(err.message)}）</li>`);
      }
      if (++done === files.length) {
        saveDB();
        renderAll();
        box.innerHTML = `<b>取込結果</b>：従業員 新規${total.empAdded}人 / 更新${total.empUpdated}人、`
          + `有給 ${total.leaveAdded}件を追加${total.skipped ? `、${total.skipped}シートを飛ばしました` : ''}<ul>${log.join('')}</ul>`
          + `<div class="hint">取込んだ有給には勤務時間帯が入っていません（Excelに時間が無いため）。`
          + `過去分をTC5へ出す必要がなければそのままで大丈夫です。</div>`;
        showToast(`Excelを取り込みました（有給 ${total.leaveAdded}件）`, 'success');
      }
    };
    reader.readAsArrayBuffer(file);
  }
}

/* ============ 8. 書き出し ============ */

// TC5の形： { 名前: { 'YYYY-MM-DD': [ {checkIn, checkOut, isPaidLeave} ] } }
function buildTc5Json(from, to, mode) {
  const out = {};
  const noTime = [];
  const noName = new Set();
  let count = 0;

  for (const l of DB.leaves) {
    if (l.date < from || l.date > to) continue;
    const emp = getEmp(l.empId);
    if (!emp || emp.mode !== mode) continue;
    if ((emp.type || 'part') !== 'part') continue;      // 正職はTC5を使わない
    if (!emp.tcName) { noName.add(emp.name); continue; }
    if (!l.in1 || !l.out1) { noTime.push(`${emp.name} ${fmtDate(l.date)}`); continue; }

    if (!out[emp.tcName]) out[emp.tcName] = {};
    if (!out[emp.tcName][l.date]) out[emp.tcName][l.date] = [];
    out[emp.tcName][l.date].push({ checkIn: l.in1, checkOut: l.out1, isPaidLeave: true });
    count++;
    if (l.in2 && l.out2) {
      out[emp.tcName][l.date].push({ checkIn: l.in2, checkOut: l.out2, isPaidLeave: true });
      count++;
    }
  }
  return { out, count, noTime, noName: [...noName] };
}

function renderExportPreview() {
  const from = document.getElementById('expFrom').value;
  const to   = document.getElementById('expTo').value;
  const mode = document.getElementById('expMode').value;
  const box  = document.getElementById('expPreview');
  if (!from || !to) { box.innerHTML = ''; return; }

  const { out, count, noTime, noName } = buildTc5Json(from, to, mode);
  const people = Object.keys(out);
  let html = `<b>${fmtDate(from)} 〜 ${fmtDate(to)}（${TC5_MODES[mode]}）</b>：`
           + `${people.length}人 / ${count}レコード`;
  if (people.length) html += `<ul>${people.map(n =>
    `<li>${esc(n)}：${Object.keys(out[n]).length}日</li>`).join('')}</ul>`;
  if (noName.length) html += `<div class="danger-text">⚠ TC5での名前が未設定のため出力できない人：${noName.map(esc).join('、')}</div>`;
  if (noTime.length) html += `<div class="hint">⚠ 勤務時間帯が空で出力されない有給 ${noTime.length}件：${noTime.slice(0, 8).map(esc).join('、')}${noTime.length > 8 ? ' ほか' : ''}</div>`;
  box.innerHTML = html;
}

function exportTc5Json() {
  const from = document.getElementById('expFrom').value;
  const to   = document.getElementById('expTo').value;
  const mode = document.getElementById('expMode').value;
  if (!from || !to) { showToast('期間を指定してください', 'warning'); return; }
  if (from > to) { showToast('開始日が終了日より後になっています', 'warning'); return; }

  const { out, count } = buildTc5Json(from, to, mode);
  if (!count) { showToast('この期間・区分に出力できる有給がありません', 'warning'); return; }

  download(`有給_TC5取込_${TC5_MODES[mode]}_${periodLabel(from, to)}.json`, JSON.stringify(out));
  showToast(`${count}レコードを書き出しました。TC5で「リストア→統合」してください`, 'success');
}

function exportLedgerJson() {
  const asOf = todayStr();
  const data = {
    exportedAt: new Date().toISOString(),
    asOf,
    employees: sortedEmployees().map(e => {
      const g = computeLedger(e, asOf);
      return {
        name: e.name, kana: e.kana, tcName: e.tcName, mode: e.mode,
        hire: e.hire, weekDays: e.weekDays, dailyHours: e.dailyHours,
        currentGrantDate: g.current ? g.current.from : null,
        grantDays: g.grantDays,
        carryOver: g.carry,
        takenInPeriod: Math.round(g.takenInPeriod * 100) / 100,
        balance: Math.round(g.balance * 100) / 100,
        dutyRemain: g.duty ? g.duty.remain : null,
        dutyDeadline: g.duty ? g.duty.deadline : null,
        nextGrant: g.next ? { date: g.next.from, days: g.next.grant } : null,
        // 管理簿の1行ぶんがそのまま入る（見本Excelの列と同じ意味）
        book: g.rows.filter(r => r.started).map(r => ({
          serviceMonths: r.months,
          grantDate: r.from, periodEnd: r.to,
          grant: r.grant, carry: r.carry, total: r.total,
          taken: Math.round(r.taken * 100) / 100,
          dutyRemain: r.duty,
          remain: Math.round(r.remain * 100) / 100,
          invalid: r.invalid,
          takenDates: r.takes.map(l => l.date)
        })),
        leaves: g.leaves.map(l => ({
          date: l.date, type: l.type, hours: l.hours,
          days: Math.round(leaveDays(l, e) * 1000) / 1000,
          in1: l.in1, out1: l.out1, in2: l.in2, out2: l.out2, note: l.note
        }))
      };
    })
  };
  download(`有給データ_${asOf.replace(/-/g, '')}.json`, JSON.stringify(data, null, 2));
  showToast('有給データを書き出しました', 'success');
}

function backupAll() {
  download(`有給管理バックアップ_${todayStr().replace(/-/g, '')}.json`, JSON.stringify(DB));
  showToast('バックアップを保存しました', 'success');
}

function restoreAll(file) {
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const o = JSON.parse(ev.target.result);
      if (!o || !Array.isArray(o.employees) || !Array.isArray(o.leaves)) {
        showToast('このツールのバックアップファイルではありません', 'error'); return;
      }
      if (!confirm(`現在のデータを置き換えます。\n従業員 ${o.employees.length}人 / 有給 ${o.leaves.length}件\nよろしいですか？`)) return;
      DB = o;
      saveDB();
      renderAll();
      showToast('復元しました', 'success');
    } catch {
      showToast('ファイルを読めませんでした', 'error');
    }
  };
  reader.readAsText(file);
}

/* ============ 8.5 年度ごとのファイル保存（フォルダ連携） ============ */
/*
 *  データはこのツールの中では1つにまとまっているが、保存は年度（4月〜翌3月）ごとに
 *  ファイルを分ける。1つのファイルに「従業員マスタ全部 ＋ その年度の有給 ＋ その年度の設定」が
 *  入るので、どの年度のファイルからでも復元できる。
 *
 *  ⚠ ファイル名は今までの正職ツール（有給管理_2026年度.json）と<b>わざと変えている</b>。
 *     同じ名前にすると、フォルダを間違えたときに正職ツールのファイルを壊してしまうため。
 */

const FS_SUPPORTED = typeof window !== 'undefined' && 'showDirectoryPicker' in window;
const YEAR_FILE_RE = /^有給管理データ_(\d{4})年度\.json$/;

let dirHandle = null;          // 連携中のフォルダ
let fileWriteTimer = null;
let knownFileYears = new Set();  // フォルダにあった年度（有給が0件になっても書き戻すため）

function yearFileName(year) { return `有給管理データ_${year}年度.json`; }

/* --- フォルダの記憶（IndexedDB） --- */

function idbOpen() {
  return new Promise((res, rej) => {
    const r = indexedDB.open('yk_fs', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('kv');
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function idbSet(key, val) {
  const db = await idbOpen();
  await new Promise((res, rej) => {
    const tx = db.transaction('kv', 'readwrite');
    tx.objectStore('kv').put(val, key);
    tx.oncomplete = res; tx.onerror = () => rej(tx.error);
  });
  db.close();
}
async function idbGet(key) {
  const db = await idbOpen();
  const v = await new Promise((res, rej) => {
    const tx = db.transaction('kv', 'readonly');
    const q = tx.objectStore('kv').get(key);
    q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error);
  });
  db.close();
  return v;
}

/* --- 書き出す中身 --- */

// 保存対象の年度（有給のある年度・設定のある年度・ファイルのあった年度）
function yearsToSave() {
  const set = new Set(knownFileYears);
  DB.leaves.forEach(l => set.add(fiscalYearOf(l.date)));
  Object.keys(DB.fiscal || {}).forEach(y => set.add(Number(y)));
  set.add(fiscalYearOf(todayStr()));
  return [...set].filter(y => Number.isFinite(y)).sort();
}

function yearFileContent(year) {
  const { start, end } = fiscalRange(year);
  return {
    format: 'yukyu-kanri-data',      // 正職ツールのファイルと見分けるための目印
    version: 1,
    fiscalYear: year,
    savedAt: new Date().toISOString(),
    employees: DB.employees,                                   // マスタは全ファイルに入れる
    tcNames: DB.tcNames,                                       // TC5の登録名も一緒に持たせる
    termMode: termMode(),                                      // 長期休暇の呼び方
    leaves: DB.leaves.filter(l => l.date >= start && l.date <= end),
    fiscal: (DB.fiscal && DB.fiscal[year]) ? { [year]: DB.fiscal[year] } : {}
  };
}

// 読み込んだ年度ファイルを1つのデータにまとめる（新しく保存されたものを優先）
function mergeYearFiles(list) {
  const sorted = [...list].sort((a, b) => String(b.savedAt || '').localeCompare(String(a.savedAt || '')));
  const emps = new Map();
  const leaves = new Map();
  const fiscal = {};
  let tcNames = null;
  let termMode = null;
  for (const c of sorted) {
    if (!tcNames && Array.isArray(c.tcNames) && c.tcNames.length) tcNames = c.tcNames;
    if (!termMode && (c.termMode === 'work' || c.termMode === 'absent')) termMode = c.termMode;
    (c.employees || []).forEach(e => { if (!emps.has(e.id)) emps.set(e.id, e); });
    (c.leaves || []).forEach(l => {
      const k = `${l.empId}|${l.date}|${l.type}`;
      if (!leaves.has(k)) leaves.set(k, l);
    });
    for (const y in (c.fiscal || {})) if (!fiscal[y]) fiscal[y] = c.fiscal[y];
  }
  return { employees: [...emps.values()], leaves: [...leaves.values()], fiscal,
           tcNames: tcNames || [...DEFAULT_TC5_NAMES], termMode: termMode || 'work' };
}

/* --- 保存 --- */

async function ensurePerm(handle, mode) {
  if (!handle) return false;
  if ((await handle.queryPermission({ mode })) === 'granted') return true;
  return (await handle.requestPermission({ mode })) === 'granted';
}

async function writeYearFiles(quiet) {
  if (!dirHandle) return;
  try {
    if (!(await ensurePerm(dirHandle, 'readwrite'))) { setFsStatus('⚠ フォルダへの書き込みが許可されていません', 'warn'); return; }
    const years = yearsToSave();
    for (const y of years) {
      const fh = await dirHandle.getFileHandle(yearFileName(y), { create: true });
      const w = await fh.createWritable();
      await w.write(JSON.stringify(yearFileContent(y), null, 1));
      await w.close();
      knownFileYears.add(y);
    }
    const t = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' });
    setFsStatus(`📁 ${dirHandle.name} ／ ${years.length}年度分（${t} 保存済）`, 'ok');
    updateSaveButton(t);
    if (!quiet) showToast(`${years.length}年度分を保存しました`, 'success');
  } catch (err) {
    setFsStatus(`⚠ 保存できませんでした（${err.message}）。バックアップを取ってください`, 'warn');
  }
}

// 変更のたびに呼ばれる。まとめて1秒後に書く
function scheduleFileWrite() {
  if (!dirHandle) return;
  clearTimeout(fileWriteTimer);
  fileWriteTimer = setTimeout(() => writeYearFiles(true), 1000);
}

/* --- 読み込み --- */

async function readYearFilesFrom(handle) {
  const list = [];
  knownFileYears = new Set();
  for await (const [name, h] of handle.entries()) {
    const m = name.match(YEAR_FILE_RE);
    if (!m || h.kind !== 'file') continue;
    try {
      const text = await (await h.getFile()).text();
      const o = JSON.parse(text);
      if (o && o.format === 'yukyu-kanri-data') { list.push(o); knownFileYears.add(Number(m[1])); }
    } catch {}
  }
  return list;
}

async function linkFolder(loadFirst) {
  if (!FS_SUPPORTED) { showToast('このブラウザはフォルダ保存に対応していません（Chrome / Edge をお使いください）', 'warning'); return; }
  try {
    const dh = await window.showDirectoryPicker({ mode: 'readwrite' });
    if (!(await ensurePerm(dh, 'readwrite'))) { showToast('書き込みが許可されませんでした', 'warning'); return; }
    dirHandle = dh;
    await idbSet('dirHandle', dh);

    const files = await readYearFilesFrom(dh);
    if (files.length && loadFirst) {
      const merged = mergeYearFiles(files);
      const msg = `フォルダ「${dh.name}」に ${files.length}年度分のファイルがありました。\n`
        + `従業員 ${merged.employees.length}人 / 有給 ${merged.leaves.length}件\n\n`
        + 'OK：ファイルの内容を読み込む（いまの画面のデータは置き換わります）\n'
        + 'キャンセル：読み込まず、いまのデータをこのフォルダに保存する';
      if (confirm(msg)) {
        DB = { employees: merged.employees, leaves: merged.leaves, fiscal: merged.fiscal, tcNames: merged.tcNames, termMode: merged.termMode };
        saveDB();
        refreshStaffYearSelect();
        renderAll();
        showToast(`${files.length}年度分を読み込みました`, 'success');
        updateFsStatus();
        return;
      }
    }
    await writeYearFiles();
  } catch (err) {
    if (err.name !== 'AbortError') showToast(`フォルダを選べませんでした（${err.message}）`, 'error');
  }
}

async function loadFromFolder() {
  if (!dirHandle) { await linkFolder(true); return; }
  try {
    if (!(await ensurePerm(dirHandle, 'read'))) { showToast('読み取りが許可されませんでした', 'warning'); return; }
    const files = await readYearFilesFrom(dirHandle);
    if (!files.length) { showToast('このフォルダに年度ファイルがありません', 'warning'); return; }
    const merged = mergeYearFiles(files);
    if (!confirm(`${files.length}年度分を読み込みます。\n従業員 ${merged.employees.length}人 / 有給 ${merged.leaves.length}件\n\nいまの画面のデータは置き換わります。よろしいですか？`)) return;
    DB = { employees: merged.employees, leaves: merged.leaves, fiscal: merged.fiscal, tcNames: merged.tcNames, termMode: merged.termMode };
    saveDB();
    refreshStaffYearSelect();
    renderAll();
    showToast(`${files.length}年度分を読み込みました`, 'success');
  } catch (err) {
    showToast(`読み込めませんでした（${err.message}）`, 'error');
  }
}

/* --- フォルダが使えないブラウザ用（ダウンロード／ファイル選択） --- */

function downloadYearFiles() {
  const years = yearsToSave();
  years.forEach((y, i) => {
    // まとめてダウンロードするとブラウザに弾かれるので少しずらす
    setTimeout(() => download(yearFileName(y), JSON.stringify(yearFileContent(y), null, 1)), i * 400);
  });
  showToast(`${years.length}年度分を書き出します`, 'success');
}

function loadYearFilesFromInput(files) {
  const list = [];
  let done = 0;
  for (const f of files) {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const o = JSON.parse(ev.target.result);
        if (o && o.format === 'yukyu-kanri-data') list.push(o);
      } catch {}
      if (++done === files.length) {
        if (!list.length) { showToast('年度ファイル（有給管理データ_◯◯年度.json）が見つかりませんでした', 'warning'); return; }
        const merged = mergeYearFiles(list);
        if (!confirm(`${list.length}年度分を読み込みます。\n従業員 ${merged.employees.length}人 / 有給 ${merged.leaves.length}件\n\nいまのデータは置き換わります。よろしいですか？`)) return;
        DB = { employees: merged.employees, leaves: merged.leaves, fiscal: merged.fiscal, tcNames: merged.tcNames, termMode: merged.termMode };
        saveDB();
        refreshStaffYearSelect();
        renderAll();
        showToast(`${list.length}年度分を読み込みました`, 'success');
      }
    };
    reader.readAsText(f);
  }
}

/* --- 状態表示 --- */

function setFsStatus(text, kind) {
  const el = document.getElementById('fsStatus');
  if (!el) return;
  el.textContent = text;
  el.className = 'fs-status' + (kind ? ' ' + kind : '');
}

// 右上の保存ボタンの表示を状態に合わせる
// 印刷は押したカードだけ（同じタブに一覧と管理簿が並んでいるため）
function printCard(innerId) {
  const inner = document.getElementById(innerId);
  const card = inner && inner.closest('.card');
  if (!card) { window.print(); return; }
  document.body.classList.add('printing');
  card.classList.add('print-target');
  const clean = () => {
    document.body.classList.remove('printing');
    card.classList.remove('print-target');
    window.removeEventListener('afterprint', clean);
  };
  window.addEventListener('afterprint', clean);
  window.print();
  setTimeout(clean, 1000);          // afterprint が来ないブラウザ向けの保険
}

function updateSaveButton(savedAt) {
  const b = document.getElementById('saveBtn');
  if (!b) return;
  if (!FS_SUPPORTED) {
    b.textContent = '💾 保存';
    b.title = 'このブラウザはフォルダ保存に対応していません。「取込・出力」タブから書き出してください。';
    b.classList.remove('unset');
    return;
  }
  if (!dirHandle) {
    b.textContent = '💾 保存先を決める';
    b.title = '保存するフォルダをまだ決めていません。押すとフォルダを選べます。';
    b.classList.add('unset');
    return;
  }
  b.textContent = savedAt ? `💾 ${savedAt} 保存済` : '💾 保存';
  b.title = `${dirHandle.name} に自動保存しています（Ctrl+S でも保存）`;
  b.classList.remove('unset');
}

function updateFsStatus() {
  updateSaveButton();
  if (!FS_SUPPORTED) {
    setFsStatus('このブラウザはフォルダ保存に対応していません。下のボタンでファイルの読み書きをしてください。', 'warn');
    return;
  }
  if (!dirHandle) {
    setFsStatus('フォルダ未連携（このブラウザの中だけに保存されています）', 'warn');
    return;
  }
  setFsStatus(`📁 ${dirHandle.name} と連携中。変更するたびに年度ごとのファイルへ自動保存します。`, 'ok');
}

// 起動時：前回のフォルダを覚えていれば、許可済みなら黙って読み込む
async function restoreFolder() {
  if (!FS_SUPPORTED) { updateFsStatus(); return; }
  try {
    const dh = await idbGet('dirHandle');
    if (dh && (await dh.queryPermission({ mode: 'readwrite' })) === 'granted') {
      dirHandle = dh;
      const files = await readYearFilesFrom(dh);
      if (files.length) {
        const merged = mergeYearFiles(files);
        DB = { employees: merged.employees, leaves: merged.leaves, fiscal: merged.fiscal, tcNames: merged.tcNames, termMode: merged.termMode };
        saveDB();
        refreshStaffYearSelect();
        renderAll();
      }
    }
  } catch {}
  updateFsStatus();
}

/* ============ 9. 初期化 ============ */

function renderAll() {
  refreshEmpSelects();
  refreshMonthFilter();
  renderCalendar();
  applyShiftDefault();
  applyInputMode();
  renderInputTerms();
  renderBalancePreview();
  renderLeaveTable();
  renderEmpTable();
  renderTcNames();
  renderDupCard();
  refreshPurgeEmpSelect();
  renderPurgePreview();
  renderLedgerTable();
  renderBook();
  renderTermInput();
  renderStaffGrid();
  renderStaffSummary();
  renderExportPreview();
}

function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  document.getElementById('themeToggle').textContent = theme === 'dark' ? '☀️' : '🌙';
  localStorage.setItem('yk_theme', theme);
}

document.addEventListener('DOMContentLoaded', () => {
  DB = loadDB();
  applyTheme(localStorage.getItem('yk_theme') || 'light');

  // タブ
  document.getElementById('tabs').addEventListener('click', e => {
    const b = e.target.closest('button[data-tab]');
    if (!b) return;
    document.querySelectorAll('#tabs button').forEach(x => x.classList.toggle('active', x === b));
    document.querySelectorAll('.tab-panel').forEach(p =>
      p.classList.toggle('active', p.id === 'tab-' + b.dataset.tab));
  });

  // 右上の保存ボタン（フォルダ未設定ならフォルダ選びから）
  document.getElementById('saveBtn').addEventListener('click', () => {
    if (dirHandle) writeYearFiles(); else linkFolder(true);
  });
  document.getElementById('themeToggle').addEventListener('click', () =>
    applyTheme((localStorage.getItem('yk_theme') || 'light') === 'dark' ? 'light' : 'dark'));

  // --- 有給入力 ---
  document.getElementById('inEmp').addEventListener('change', () => {
    // 前に選んでいた人の時間帯が残らないように一度消す（全日ならこのあと所定が入る）
    for (const id of ['inTime1a', 'inTime1b', 'inTime2a', 'inTime2b']) {
      document.getElementById(id).value = '';
    }
    document.getElementById('inHours').value = '';
    applyShiftDefault(); applyInputMode(); renderCalendar(); renderBalancePreview();
  });
  document.getElementById('paidCalendar').addEventListener('click', e => {
    const nav = e.target.closest('.cal-nav');
    if (nav) {
      paidCursor = new Date(paidCursor.getFullYear(), paidCursor.getMonth() + (+nav.dataset.nav), 1);
      renderCalendar(); return;
    }
    const cell = e.target.closest('.cal-day');
    if (!cell) return;
    const ds = cell.dataset.date;
    if (paidSelected.has(ds)) paidSelected.delete(ds); else paidSelected.add(ds);
    syncDatesField(); renderCalendar(); renderBalancePreview(); renderInputTerms();
  });
  document.getElementById('inDates').addEventListener('input', () => {
    const ds = collectInputDates();
    paidSelected = new Set(ds || []);
    renderCalendar(); renderBalancePreview(); renderInputTerms();
  });
  document.getElementById('inType').addEventListener('click', e => {
    const b = e.target.closest('button[data-type]');
    if (b) setInputType(b.dataset.type);
  });
  // 有給入力タブからの長期休暇の入力
  document.getElementById('inTermArea').addEventListener('change', e => {
    const inp = e.target.closest('input[data-term]');
    if (!inp) return;
    const emp = getEmp(document.getElementById('inEmp').value);
    if (!emp) return;
    const year = fiscalYearOf(dates0());
    const w = getTermWork(year, emp.id);
    w[inp.dataset.term] = inp.value.trim() === '' ? '' : (parseFloat(inp.value) || 0);
    saveDB();
    renderAll();
    showToast(`${year}年度の長期休暇を更新しました`, 'success');
  });

  document.getElementById('inHours').addEventListener('input', renderBalancePreview);
  document.getElementById('hoursQuick').addEventListener('click', e => {
    const b = e.target.closest('button[data-h]');
    if (!b) return;
    document.getElementById('inHours').value = b.dataset.h;
    renderBalancePreview();
  });
  document.getElementById('inSubmit').addEventListener('click', submitLeave);
  document.getElementById('inCancelEdit').addEventListener('click', () => {
    cancelEditLeave();
    showToast('編集をやめました', 'info');
  });
  // 時間帯を入れたら時間数を自動計算
  for (const id of ['inTime1a', 'inTime1b', 'inTime2a', 'inTime2b']) {
    document.getElementById(id).addEventListener('change', syncHoursFromTimes);
  }
  document.getElementById('inClear').addEventListener('click', () => {
    cancelEditLeave(true);
    paidSelected.clear(); syncDatesField();
    document.getElementById('inHours').value = '';
    document.getElementById('inNote').value = '';
    applyInputMode();
    renderCalendar(); renderBalancePreview();
  });

  // 有給：チェックしてまとめて削除
  document.getElementById('leaveCheckAll').addEventListener('change', e => {
    document.querySelectorAll('.leaveChk').forEach(c => { c.checked = e.target.checked; });
  });
  document.getElementById('leaveDeleteChecked').addEventListener('click', () => {
    const ids = checkedIds('.leaveChk');
    if (!ids.length) { showToast('削除する行にチェックを入れてください', 'warning'); return; }
    if (!confirm(`チェックした有給 ${ids.length}件 を削除します。よろしいですか？`)) return;
    const n = removeLeaves(ids);
    saveDB(); renderAll();
    showToast(`${n}件の有給を削除しました`, 'success');
  });
  document.getElementById('leaveDeleteFiltered').addEventListener('click', () => {
    if (!lastLeaveRows.length) { showToast('削除する有給がありません', 'warning'); return; }
    if (!confirm(`いま絞り込んで表示している ${lastLeaveRows.length}件 をすべて削除します。\n元に戻せません。よろしいですか？`)) return;
    const n = removeLeaves(lastLeaveRows.map(x => x.l.id));
    saveDB(); renderAll();
    showToast(`${n}件の有給を削除しました`, 'success');
  });

  document.getElementById('leaveSearch').addEventListener('input', renderLeaveTable);
  document.getElementById('leaveMonth').addEventListener('change', renderLeaveTable);
  document.querySelector('#leaveTable tbody').addEventListener('click', e => {
    const ed = e.target.closest('button[data-edit]');
    if (ed) { startEditLeave(ed.dataset.edit); return; }
    const b = e.target.closest('button[data-del]');
    if (!b) return;
    const l = DB.leaves.find(x => x.id === b.dataset.del);
    if (!l) return;
    const emp = getEmp(l.empId);
    if (!confirm(`${emp ? emp.name : ''} ${fmtDate(l.date)} の有給を削除します。よろしいですか？`)) return;
    DB.leaves = DB.leaves.filter(x => x.id !== l.id);
    saveDB(); renderAll();
    showToast('削除しました', 'success');
  });

  // --- 残日数一覧 ---
  document.getElementById('bookEmp').addEventListener('change', renderBook);
  document.getElementById('bookPrintBtn').addEventListener('click', () => printCard('bookPrint'));
  document.getElementById('bookXlsxOne').addEventListener('click', () => exportBookXlsx('one'));
  document.getElementById('bookXlsxAll').addEventListener('click', () => exportBookXlsx('all'));

  // --- 正職集計 ---
  document.getElementById('staffYear').addEventListener('change', () => {
    renderTermInput(); renderStaffGrid(); renderStaffSummary();
  });
  document.getElementById('gridArea').addEventListener('click', e => {
    const b = e.target.closest('button[data-gmonth]');
    if (!b) return;
    const i = FISCAL_MONTHS.indexOf(staffGridMonth);
    const next = i + Number(b.dataset.gmonth);
    if (next < 0 || next >= FISCAL_MONTHS.length) return;   // 年度の外へは動かさない
    staffGridMonth = FISCAL_MONTHS[next];
    renderStaffGrid();
  });
  document.getElementById('gridArea').addEventListener('change', e => {
    if (!e.target.closest('#gridMonth')) return;
    staffGridMonth = Number(e.target.value);
    renderStaffGrid();
  });
  document.getElementById('gridPrintBtn').addEventListener('click', () => printCard('gridPrint'));
  document.getElementById('staffPrintBtn').addEventListener('click', () => printCard('staffPrint'));
  document.getElementById('staffCsvBtn').addEventListener('click', exportStaffCsv);
  // 長期休暇の入力（最大日数と、人ごとの出勤数）
  document.getElementById('termArea').addEventListener('click', e => {
    const b = e.target.closest('#termModeSeg button[data-mode]');
    if (!b || b.dataset.mode === termMode()) return;
    // すでに数字が入っているなら、意味が入れ替わることを伝える
    const hasValue = Object.values(DB.fiscal || {}).some(f =>
      Object.values(f.work || {}).some(w =>
        TERM_FIELDS.some(t => w[t.key] !== '' && w[t.key] != null)));
    if (hasValue && !confirm(
      `入力のしかたを「${TERM_MODES[b.dataset.mode]}」に切り替えます。\n\n`
      + `すでに入っている数字は入れ替えずにそのまま残るので、\n`
      + `「${termModeLabel()}として入れた数字」が「${TERM_MODES[b.dataset.mode]}」として読み替えられ、\n`
      + '取得日数と残り日数が変わります。よろしいですか？')) return;
    DB.termMode = b.dataset.mode;
    saveDB();
    renderAll();
    showToast(`長期休暇は「${termModeLabel()}」で入力します`, 'success');
  });
  document.getElementById('termArea').addEventListener('change', e => {
    const year = currentStaffYear();
    const mx = e.target.closest('input[data-max]');
    if (mx) {
      getFiscal(year).maxLeave[mx.dataset.max] = parseFloat(mx.value) || 0;
      saveDB(); renderTermInput(); renderStaffSummary(); return;
    }
  });
  document.getElementById('staffJsonBtn').addEventListener('click', () =>
    document.getElementById('staffJsonFile').click());
  document.getElementById('staffJsonFile').addEventListener('change', e => {
    if (e.target.files.length) handleStaffJsonFiles([...e.target.files]);
    e.target.value = '';
  });
  // 「無効」チェック＝その回の付与を0日にする（見本のD列と同じ）
  document.getElementById('bookArea').addEventListener('change', e => {
    const emp = getEmp(document.getElementById('bookEmp').value);
    if (!emp) return;
    // 週労働日数はその回の付与ごとに持つ（Excelと同じ）
    const wd = e.target.closest('input[data-week]');
    if (wd) {
      const v = Math.min(Math.max(Math.round(Number(wd.value) || 5), 1), 5);
      emp.weekDaysByGrant = emp.weekDaysByGrant || {};
      emp.weekDaysByGrant[wd.dataset.week] = v;
      saveDB(); renderAll();
      showToast(`${fmtDate(wd.dataset.week)}の付与を週${v}日で計算しました`, 'success');
      return;
    }
    const cb = e.target.closest('input[data-invalid]');
    if (!cb) return;
    emp.invalidGrants = emp.invalidGrants || {};
    if (cb.checked) emp.invalidGrants[cb.dataset.invalid] = true;
    else delete emp.invalidGrants[cb.dataset.invalid];
    saveDB();
    renderAll();
    showToast(cb.checked ? 'この回の付与を無効（0日）にしました' : '無効を解除しました', 'success');
  });

  // --- 従業員マスタ ---
  // 二重登録のまとめ／削除
  document.getElementById('dupArea').addEventListener('click', e => {
    const keep = e.target.closest('button[data-keep]');
    if (keep) {
      const group = duplicateGroups().find(g => g.some(x => x.id === keep.dataset.keep));
      if (!group) return;
      const others = group.filter(x => x.id !== keep.dataset.keep);
      const target = getEmp(keep.dataset.keep);
      if (!confirm(`「${target.name}」にまとめます。\n`
        + others.map(o => `・${o.name}（有給${leavesOf(o.id).length}件）`).join('\n')
        + `\nの有給を移してから、この${others.length}人を削除します。よろしいですか？`)) return;
      const moved = mergeEmployees(keep.dataset.keep, others.map(o => o.id));
      saveDB(); fillEmpForm(null); renderAll();
      showToast(`${target.name} にまとめました（有給 ${moved}件を移動）`, 'success');
      return;
    }
    const drop = e.target.closest('button[data-drop]');
    if (drop) {
      const emp = getEmp(drop.dataset.drop);
      if (!emp) return;
      const n = leavesOf(emp.id).length;
      if (!confirm(`「${emp.name}」を削除します。\nこの人に付いている有給 ${n} 件も一緒に消えます。よろしいですか？`)) return;
      DB.employees = DB.employees.filter(x => x.id !== emp.id);
      DB.leaves = DB.leaves.filter(l => l.empId !== emp.id);
      saveDB(); fillEmpForm(null); renderAll();
      showToast('削除しました', 'success');
    }
  });

  document.getElementById('tcNameSave').addEventListener('click', saveTcNames);
  document.getElementById('tcNameJsonBtn').addEventListener('click', () =>
    document.getElementById('tcNameJsonFile').click());
  document.getElementById('tcNameJsonFile').addEventListener('change', e => {
    if (e.target.files.length) handleTcNameJsonFiles([...e.target.files]);
    e.target.value = '';
  });
  document.getElementById('tcNameReset').addEventListener('click', () => {
    if (!confirm('TC5の登録名をすべて消します。よろしいですか？')) return;
    DB.tcNames = [];
    saveDB(); renderAll();
    showToast('TC5の登録名を消しました', 'success');
  });

  document.getElementById('empType').addEventListener('change', applyEmpTypeFields);
  document.getElementById('empSave').addEventListener('click', saveEmployee);
  document.getElementById('empCancel').addEventListener('click', () => fillEmpForm(null));
  document.getElementById('empDelete').addEventListener('click', deleteEmployee);
  document.querySelector('#empTable tbody').addEventListener('click', e => {
    const b = e.target.closest('button[data-edit]');
    if (b) { fillEmpForm(getEmp(b.dataset.edit)); window.scrollTo({ top: 0, behavior: 'smooth' }); return; }
    const d = e.target.closest('button[data-remove]');
    if (d) {
      const emp = getEmp(d.dataset.remove);
      if (!emp) return;
      const n = leavesOf(emp.id).length;
      if (!confirm(`「${emp.name}」を削除します。\nこの人の有給 ${n}件 も一緒に消えます。よろしいですか？`)) return;
      removeEmployees([emp.id]);
      saveDB(); fillEmpForm(null); renderAll();
      showToast('削除しました', 'success');
    }
  });
  document.getElementById('empCheckAll').addEventListener('change', e => {
    document.querySelectorAll('.empChk').forEach(c => { c.checked = e.target.checked; });
  });
  document.getElementById('empDeleteChecked').addEventListener('click', () => {
    const ids = checkedIds('.empChk');
    if (!ids.length) { showToast('削除する人にチェックを入れてください', 'warning'); return; }
    const names = ids.map(i => getEmp(i)).filter(Boolean);
    const leaves = names.reduce((s, e) => s + leavesOf(e.id).length, 0);
    if (!confirm(`${ids.length}人を削除します。\n`
      + names.slice(0, 10).map(e => '・' + e.name).join('\n')
      + (names.length > 10 ? `\nほか${names.length - 10}人` : '')
      + `\n\nこの人たちの有給 ${leaves}件 も一緒に消えます。よろしいですか？`)) return;
    removeEmployees(ids);
    saveDB(); fillEmpForm(null); renderAll();
    showToast(`${ids.length}人を削除しました`, 'success');
  });

  // --- 取込・出力 ---
  document.getElementById('xlsxBtn').addEventListener('click', () =>
    document.getElementById('xlsxFile').click());
  document.getElementById('xlsxFile').addEventListener('change', e => {
    if (e.target.files.length) handleXlsxFiles([...e.target.files]);
    e.target.value = '';
  });

  const setPeriod = offset => {
    const p = closingPeriod(todayStr(), offset);
    document.getElementById('expFrom').value = p.start;
    document.getElementById('expTo').value = p.end;
    renderExportPreview();
  };
  document.getElementById('expPeriodPrev').addEventListener('click', () => setPeriod(-1));
  document.getElementById('expPeriodCur').addEventListener('click',  () => setPeriod(0));
  document.getElementById('expPeriodNext').addEventListener('click', () => setPeriod(1));
  document.getElementById('expFrom').addEventListener('change', renderExportPreview);
  document.getElementById('expTo').addEventListener('change', renderExportPreview);
  document.getElementById('expMode').addEventListener('change', renderExportPreview);
  document.getElementById('expBtn').addEventListener('click', exportTc5Json);
  document.getElementById('expLedgerBtn').addEventListener('click', exportLedgerJson);

  // 条件で絞って有給を削除
  for (const id of ['purgeSource', 'purgeEmp', 'purgeFrom', 'purgeTo']) {
    document.getElementById(id).addEventListener('change', renderPurgePreview);
  }
  document.getElementById('purgeBtn').addEventListener('click', () => {
    const hits = purgeTargets();
    if (!hits.length) { showToast('条件に合う有給がありません', 'warning'); return; }
    const src = document.getElementById('purgeSource').value;
    if (!confirm(`条件に合う有給 ${hits.length}件 を削除します。`
      + (src ? `\n対象：${SOURCE_LABEL[src]}のもの` : '\n対象：すべての取込元')
      + '\n\n従業員そのものは消えません。元に戻せません。よろしいですか？')) return;
    const n = removeLeaves(hits.map(l => l.id));
    saveDB(); renderAll();
    showToast(`${n}件の有給を削除しました`, 'success');
  });

  // 年度ごとのファイル保存
  document.getElementById('fsLinkBtn').addEventListener('click', () => linkFolder(true));
  document.getElementById('fsLoadBtn').addEventListener('click', loadFromFolder);
  document.getElementById('fsSaveBtn').addEventListener('click', () => writeYearFiles());
  document.getElementById('fsDownloadBtn').addEventListener('click', downloadYearFiles);
  document.getElementById('fsPickBtn').addEventListener('click', () =>
    document.getElementById('fsPickFile').click());
  document.getElementById('fsPickFile').addEventListener('change', e => {
    if (e.target.files.length) loadYearFilesFromInput([...e.target.files]);
    e.target.value = '';
  });
  // Ctrl+S で保存（エクセル感覚）
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === 's') {
      e.preventDefault();
      if (dirHandle) writeYearFiles(); else linkFolder(true);
    }
  });
  if (!FS_SUPPORTED) document.getElementById('fsSupported').classList.add('hidden');
  else document.getElementById('fsFallback').classList.add('hidden');

  document.getElementById('backupBtn').addEventListener('click', backupAll);
  document.getElementById('restoreBtn').addEventListener('click', () =>
    document.getElementById('restoreFile').click());
  document.getElementById('restoreFile').addEventListener('change', e => {
    if (e.target.files[0]) restoreAll(e.target.files[0]);
    e.target.value = '';
  });
  document.getElementById('wipeBtn').addEventListener('click', () => {
    if (!confirm('従業員マスタと有給の記録を全部消します。元に戻せません。よろしいですか？')) return;
    if (!confirm('本当に削除しますか？先にバックアップを取っておくことを強くおすすめします。')) return;
    DB = { employees: [], leaves: [], fiscal: {}, tcNames: [...DEFAULT_TC5_NAMES], termMode: 'work' };
    saveDB(); fillEmpForm(null); renderAll();
    showToast('全データを削除しました', 'success');
  });

  setPeriod(0);
  refreshStaffYearSelect();
  fillEmpForm(null);
  renderAll();
  restoreFolder();
});
