/**
 * 人生の代理店 お繋ぎ申請 集計ダッシュボード
 *
 * - 元データ（お繋ぎ申請タブ）は読むだけ。一切書き換えない。
 * - 書き込むのは「集計ダッシュボード」タブだけ（無ければ作る）。
 * - 列は見出し名で探すので、列の並びが変わっても動く。
 *
 * 使い方: 拡張機能 > Apps Script に貼り付け → updateDashboard を実行。
 *        毎朝自動にしたい場合は installDailyTrigger を1回だけ実行。
 */

const CONFIG = {
  sourceSheet: 'お繋ぎ申請',
  outputSheet: '集計ダッシュボード',
  noReplyDays: 14,      // 打診後この日数以上返事なし → 催促候補
  stuckDays: 21,        // 申請からこの日数以上 未完了 → 長期滞留
  receptionDays: 3,     // 受付のままこの日数以上 → 打診漏れ疑い
  bulkPerDay: 20,       // 同じ人が1日にこの件数以上申請 → 大量申請
  statuses: ['受付', '確認中', '承諾', '成立', '保留', '見送り'],
};

function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.sourceSheet);
  if (!src) throw new Error('タブが見つかりません: ' + CONFIG.sourceSheet);

  const values = src.getDataRange().getValues();
  const header = values[0].map(String);
  const col = (name) => {
    const i = header.indexOf(name);
    if (i < 0) throw new Error('見出しが見つかりません: ' + name);
    return i;
  };
  const C = {
    date: col('申請日時'), from: col('申請者名'), to: col('申請先名'),
    status: col('状態'), memo: col('運営メモ'), sentAt: col('打診した日'),
    reply: col('相手の返事'), sent: col('打診済み'),
  };

  const today = startOfDay(new Date());
  const rows = values.slice(1)
    .map((r, i) => ({ r, rowNo: i + 2 }))
    .filter(({ r }) => String(r[C.status]).trim() !== '');

  // ---- 1. 状態別件数 ----
  const byStatus = {};
  CONFIG.statuses.forEach(s => (byStatus[s] = 0));
  rows.forEach(({ r }) => {
    const s = String(r[C.status]).trim();
    byStatus[s] = (byStatus[s] || 0) + 1;
  });
  const total = rows.length;
  const decided = byStatus['成立'] + byStatus['見送り'];
  const successRate = decided ? byStatus['成立'] / decided : 0;

  // ---- 2. 月別 申請数・成立数 ----
  const byMonth = {};
  rows.forEach(({ r }) => {
    const d = toDate(r[C.date]);
    if (!d) return;
    const key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM');
    byMonth[key] = byMonth[key] || { apply: 0, done: 0 };
    byMonth[key].apply++;
    if (String(r[C.status]).trim() === '成立') byMonth[key].done++;
  });

  // ---- 3. 要対応アラート ----
  const alerts = [];
  const push = (type, x, note) => alerts.push([
    type, x.rowNo, fmt(toDate(x.r[C.date])), x.r[C.from], x.r[C.to],
    x.r[C.status], note,
  ]);
  const isTrue = (v) => v === true || String(v).toUpperCase() === 'TRUE';
  const holdMemo = (x) => /打診停止|大畑さん/.test(String(x.r[C.memo]));

  // 大量申請（同じ人・同じ日）
  const bulk = {};
  rows.forEach(x => {
    const d = toDate(x.r[C.date]);
    if (!d) return;
    const k = x.r[C.from] + '|' + fmt(d);
    bulk[k] = (bulk[k] || 0) + 1;
  });

  rows.forEach(x => {
    const s = String(x.r[C.status]).trim();
    const applied = toDate(x.r[C.date]);
    const sentAt = toDate(x.r[C.sentAt]);
    const ageApply = applied ? daysBetween(applied, today) : null;
    const ageSent = sentAt ? daysBetween(sentAt, today) : null;
    const hasReply = String(x.r[C.reply]).trim() !== '';
    if (holdMemo(x) && s !== '承諾') return; // 運営判断待ちは静観

    if (s === '確認中' && !isTrue(x.r[C.sent])) {
      push('①確認中なのに未打診', x, '状態と打診済みが不一致');
    }
    if (s === '受付' && ageApply !== null && ageApply >= CONFIG.receptionDays) {
      push('②受付のまま', x, ageApply + '日経過');
    }
    if (s === '承諾') {
      push('③承諾のまま（案内待ち）', x, sentAt ? '打診から' + ageSent + '日' : '');
    }
    if (s === '確認中' && sentAt && !hasReply && ageSent >= CONFIG.noReplyDays) {
      push('④返事なし（催促候補）', x, '打診から' + ageSent + '日');
    }
    if (['受付', '確認中', '承諾'].includes(s) && ageApply !== null && ageApply >= CONFIG.stuckDays) {
      push('⑤長期滞留', x, '申請から' + ageApply + '日');
    }
  });

  const bulkList = Object.keys(bulk)
    .filter(k => bulk[k] >= CONFIG.bulkPerDay)
    .map(k => { const [name, day] = k.split('|'); return [name, day, bulk[k]]; });

  // ---- 書き出し ----
  let out = ss.getSheetByName(CONFIG.outputSheet);
  if (!out) out = ss.insertSheet(CONFIG.outputSheet);
  out.clear();

  const blocks = [];
  blocks.push([['集計ダッシュボード（自動作成・手で書き込まない）', '最終更新 ' + fmtTime(new Date())]]);
  blocks.push([['■ 全体'], ['申請 合計', total], ['成立率（成立÷(成立+見送り)）', pct(successRate)]]);
  blocks.push([['■ 状態別', '件数', '割合']].concat(
    Object.keys(byStatus).map(s => [s, byStatus[s], pct(total ? byStatus[s] / total : 0)])));
  blocks.push([['■ 月別', '申請', '成立']].concat(
    Object.keys(byMonth).sort().map(m => [m, byMonth[m].apply, byMonth[m].done])));
  blocks.push([['■ 大量申請（1人1日' + CONFIG.bulkPerDay + '件以上）', '日付', '件数']].concat(
    bulkList.length ? bulkList : [['なし']]));
  blocks.push([['■ 要対応（' + alerts.length + '件）', '行', '申請日', '申請者', '申請先', '状態', 'メモ']].concat(
    alerts.length ? alerts.sort((a, b) => a[0] < b[0] ? -1 : 1) : [['なし']]));

  let row = 1;
  blocks.forEach(b => {
    const width = Math.max.apply(null, b.map(r => r.length));
    const grid = b.map(r => r.concat(Array(width - r.length).fill('')));
    out.getRange(row, 1, grid.length, width).setValues(grid);
    out.getRange(row, 1, 1, width).setFontWeight('bold');
    row += grid.length + 1;
  });
  out.setFrozenRows(1);
  out.autoResizeColumns(1, 7);
}

/** 毎朝6:30頃に自動更新（1回だけ実行） */
function installDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateDashboard')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('updateDashboard').timeBased().everyDays(1).atHour(6).create();
}

// ---- 小道具 ----
function toDate(v) {
  if (v instanceof Date) return v;
  const s = String(v).trim();
  const m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
}
function startOfDay(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function daysBetween(a, b) { return Math.floor((startOfDay(b) - startOfDay(a)) / 86400000); }
function fmt(d) { return d ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd') : ''; }
function fmtTime(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'); }
function pct(x) { return (x * 100).toFixed(1) + '%'; }
