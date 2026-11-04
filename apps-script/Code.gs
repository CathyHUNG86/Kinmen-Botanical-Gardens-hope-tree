/**
 * 金門植物園希望樹 — Google Apps Script 後端
 *
 * 試算表分頁：
 *   回饋：每一列是一則回饋（＝一片葉子）。同一位老師（以 Email 辨識）同一堂課最多 2 則，
 *         且文字與照片都不能和自己之前送出的重複。送出後立即長出葉子，內容也立即公開（不含姓名與聯絡方式）。
 *         若有不適合公開的內容，把「審核狀態」改成「不公開」即可隱藏文字與照片，葉子仍保留。
 * 照片存在雲端硬碟「希望樹回饋照片」資料夾（僅您可見，不公開）。
 */

const SHEET_FEEDBACK = '回饋';
// 「抽獎結果」「領獎狀態」兩欄是舊版抽獎功能留下的，保留欄位以免舊資料錯位
const HEADERS = ['送出時間', '葉子編號', '學校', '教師姓名', 'Email', '電話', '來園年月', '參與課程',
  '回饋內容', '照片連結', '同意公開', '審核狀態', '抽獎結果', '照片指紋', '領獎狀態'];
const COL = { school: 2, course: 7, visit: 6, email: 4, text: 8, photo: 9, isPublic: 10, status: 11, prize: 12, hash: 13, claim: 14 };

/** 來園年月：試算表可能把「2026-10」自動轉成日期，這裡統一轉回「2026-10」文字 */
function monthText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM');
  return String(v || '');
}
const MAX_PER_COURSE = 2; // 每位老師同一堂課程最多回饋次數

/** 第一次使用：在編輯器選 setup 後按「執行」 */
function setup() {
  const ss = SpreadsheetApp.getActive();
  if (/^未命名|^Untitled/.test(ss.getName())) ss.rename('金門植物園希望樹｜回饋資料');
  let fb = ss.getSheetByName(SHEET_FEEDBACK) || ss.insertSheet(SHEET_FEEDBACK);
  if (fb.getLastRow() === 0) { fb.appendRow(HEADERS); fb.setFrozenRows(1); }
  fb.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold'); // 舊版表頭也一併補上新欄位


  // 審核狀態下拉選單
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['待審核', '已公開', '不公開'], true).build();
  fb.getRange(2, COL.status + 1, 1000, 1).setDataValidation(rule);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PHOTO_FOLDER')) {
    props.setProperty('PHOTO_FOLDER', DriveApp.createFolder('希望樹回饋照片').getId());
  }
}

// 網頁本體放在 GitHub：更新 GitHub 上的 index.html，正式網站約 5 分鐘內自動跟著更新，不需要再改 Apps Script。
const PAGE_URL = 'https://raw.githubusercontent.com/CathyHUNG86/Kinmen-Botanical-Gardens-hope-tree/main/index.html';

function loadPage_() {
  const cache = CacheService.getScriptCache();
  const fresh = cache.get('page');
  if (fresh) return fresh;
  try {
    const res = UrlFetchApp.fetch(PAGE_URL, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const html = res.getContentText('UTF-8');
      if (html.length < 95000) { cache.put('page', html, 300); cache.put('page-backup', html, 21600); }
      return html;
    }
  } catch (e) { /* GitHub 暫時連不上時，改用備份 */ }
  return cache.get('page-backup') || '<p style="font-family:sans-serif;padding:24px">希望樹暫時無法載入，請稍後重新整理。</p>';
}

function doGet() {
  return HtmlService.createHtmlOutput(loadPage_())
    .setTitle('金門植物園希望樹')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 前端讀取整棵樹：只回傳公開需要的欄位，不含姓名與聯絡方式 */
function getTree() {
  const fb = SpreadsheetApp.getActive().getSheetByName(SHEET_FEEDBACK);
  const rows = fb.getLastRow() > 1 ? fb.getRange(2, 1, fb.getLastRow() - 1, HEADERS.length).getValues() : [];
  const leaves = rows.map((r, i) => ({
    n: i + 1,
    school: String(r[COL.school]),
    course: String(r[COL.course]),
    month: monthText_(r[COL.visit]),
    text: r[COL.status] === '不公開' ? '' : String(r[COL.text]),
    hasPhoto: r[COL.status] !== '不公開' && !!r[COL.photo]
  }));
  return { leaves };
}

/** 點葉子時讀取照片：只給公開中的回饋，照片本身仍保持私人，不必開放雲端硬碟分享 */
function getLeafPhotos(n) {
  const fb = SpreadsheetApp.getActive().getSheetByName(SHEET_FEEDBACK);
  n = Math.floor(Number(n));
  if (!(n >= 1) || n + 1 > fb.getLastRow()) return [];
  const r = fb.getRange(n + 1, 1, 1, HEADERS.length).getValues()[0];
  if (r[COL.status] === '不公開') return [];
  return String(r[COL.photo] || '').split('\n').filter(Boolean).slice(0, 3).map(u => {
    const id = (u.match(/[-\w]{25,}/) || [])[0];
    if (!id) return '';
    try {
      const b = DriveApp.getFileById(id).getBlob();
      return 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes());
    } catch (e) { return ''; }
  }).filter(Boolean);
}

/** 收到回饋：存資料、存照片，回傳葉子編號 */
function submitFeedback(p) {
  const clean = (v, n) => String(v || '').trim().slice(0, n);
  const school = clean(p.school, 40), teacher = clean(p.teacher, 20), email = clean(p.email, 120).toLowerCase();
  const text = clean(p.text, 800), photos = (p.photos || []).slice(0, 3);
  if (!school || !teacher) throw new Error('請填寫學校與教師姓名');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email 格式不正確');
  const course = clean(p.course, 20);
  if (!course) throw new Error('請選擇參與的課程');
  if (!text && !photos.length) throw new Error('請寫下回饋或上傳照片');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const fb = SpreadsheetApp.getActive().getSheetByName(SHEET_FEEDBACK);
    const prev = fb.getLastRow() > 1 ? fb.getRange(2, 1, fb.getLastRow() - 1, HEADERS.length).getValues() : [];
    const mine = prev.filter(r => String(r[COL.email]).toLowerCase() === email);

    // 規則一：同一堂課最多回饋 2 次
    const sameCourse = mine.filter(r => String(r[COL.course]) === course).length;
    if (sameCourse >= MAX_PER_COURSE) {
      throw new Error(`您在「${course}」已經回饋 ${MAX_PER_COURSE} 次了，每堂課最多 ${MAX_PER_COURSE} 次。歡迎分享其他課程的回饋`);
    }
    // 規則二：文字不可和自己之前送出的重複（忽略空白與標點）
    const key = normText_(text);
    if (key && mine.some(r => normText_(r[COL.text]) === key)) {
      throw new Error('這段回饋內容和您之前送出的一樣，請寫下不同的觀察或故事');
    }
    // 規則三：照片不可和自己之前上傳的重複
    const hashes = photos.map(ph => md5_(ph.data || ''));
    const oldHashes = new Set(mine.flatMap(r => String(r[COL.hash] || '').split(',').filter(Boolean)));
    if (hashes.some(h => oldHashes.has(h)) || new Set(hashes).size < hashes.length) {
      throw new Error('有照片之前已經上傳過了，請換一張不同的照片');
    }

    // 照片
    const links = [];
    if (photos.length) {
      const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PHOTO_FOLDER'));
      photos.forEach((ph, i) => {
        if (!ph.data || ph.data.length > 6000000) return;
        const blob = Utilities.newBlob(Utilities.base64Decode(ph.data), 'image/jpeg',
          `${Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmmss')}_${school}_${i + 1}.jpg`);
        links.push(folder.createFile(blob).getUrl());
      });
    }

    const leafNo = fb.getLastRow(); // 標題列佔 1 列，所以新列的編號 = 目前列數

    fb.appendRow([new Date(), leafNo, school, teacher, email, clean(p.phone, 20), clean(p.visit, 10) ? "'" + clean(p.visit, 10) : '', course,
      text, links.join('\n'), '是', '已公開', '', hashes.join(',')]);

    return { leafNo };
  } finally {
    lock.releaseLock();
  }
}

/** 比對用：去掉空白、標點與大小寫差異 */
function normText_(t) {
  return String(t || '').toLowerCase().replace(/[\s\p{P}\p{S}]/gu, '');
}

function md5_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/* ============================================================
 * 管理面板：試算表上方選單「希望樹管理 → 開啟管理面板」
 * 只有試算表擁有者能開啟（面板帶有管理金鑰，網站訪客無法呼叫這些功能）。
 * ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🌳 希望樹管理')
    .addItem('開啟管理面板', 'openAdmin')
    .addToUi();
}

function openAdmin() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('ADMIN_TOKEN');
  if (!token) { token = Utilities.getUuid(); props.setProperty('ADMIN_TOKEN', token); }
  const html = HtmlService.createHtmlOutput(ADMIN_HTML.replace('__TOKEN__', token))
    .setWidth(1000).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, '希望樹管理面板');
}

function checkToken_(token) {
  const ok = token && token === PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN');
  if (!ok) throw new Error('沒有管理權限');
}

function adminData(token) {
  checkToken_(token);
  const ss = SpreadsheetApp.getActive();
  const fb = ss.getSheetByName(SHEET_FEEDBACK);
  const n = fb.getLastRow() - 1;
  const vals = n > 0 ? fb.getRange(2, 1, n, HEADERS.length).getValues() : [];
  const rows = vals.map((r, i) => ({
    row: i + 2,
    ts: r[0] instanceof Date ? r[0].getTime() : 0,
    time: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Taipei', 'yyyy/MM/dd HH:mm') : String(r[0]),
    leaf: i + 1,
    school: String(r[COL.school]), teacher: String(r[3]), email: String(r[COL.email]), phone: String(r[5]),
    course: String(r[COL.course]), visit: monthText_(r[COL.visit]), text: String(r[COL.text]),
    photos: String(r[COL.photo] || '').split('\n').filter(Boolean),
    hidden: r[COL.status] === '不公開'
  }));
  let siteUrl = '';
  try { siteUrl = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  return { rows, sheetUrl: ss.getUrl(), siteUrl };
}

function findRow_(fb, row, ts) {
  if (row < 2 || row > fb.getLastRow()) return false;
  const v = fb.getRange(row, 1).getValue();
  return (v instanceof Date ? v.getTime() : 0) === ts;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function adminSetHidden(token, row, ts, hidden) {
  checkToken_(token);
  return withLock_(() => {
    const fb = SpreadsheetApp.getActive().getSheetByName(SHEET_FEEDBACK);
    if (!findRow_(fb, row, ts)) throw new Error('資料已經變動，請按「重新整理」後再試一次');
    fb.getRange(row, COL.status + 1).setValue(hidden ? '不公開' : '已公開');
  });
}

/** 刪除回饋：整列刪除、照片移到垃圾桶（30 天內可從雲端硬碟垃圾桶救回），葉子編號重新排序 */
function adminDelete(token, items) {
  checkToken_(token);
  return withLock_(() => {
    const ss = SpreadsheetApp.getActive();
    const fb = ss.getSheetByName(SHEET_FEEDBACK);
    const list = (items || []).slice().sort((a, b) => b.row - a.row);
    for (const it of list) {
      if (!findRow_(fb, it.row, it.ts)) throw new Error('資料已經變動，請按「重新整理」後再試一次');
    }
    let deleted = 0, photos = 0;
    for (const it of list) {
      const r = fb.getRange(it.row, 1, 1, HEADERS.length).getValues()[0];
      String(r[COL.photo] || '').split('\n').forEach(u => {
        const id = (u.match(/[-\w]{25,}/) || [])[0];
        if (id) { try { DriveApp.getFileById(id).setTrashed(true); photos++; } catch (e) {} }
      });
      fb.deleteRow(it.row);
      deleted++;
    }
    const n = fb.getLastRow() - 1;
    if (n > 0) fb.getRange(2, 2, n, 1).setValues(Array.from({ length: n }, (_, i) => [i + 1]));
    return { deleted, photos };
  });
}

const ADMIN_HTML = String.raw`<!doctype html><html><head><meta charset="utf-8">
<style>
:root{--ink:#1D2A21;--muted:#5A6A5E;--line:#D2DBCC;--bg:#F4F7F1;--card:#FFFFFF;--green:#2D6A48;--gold:#B8862A;--gold-soft:#F7EBCB;--red:#B23B2A;--red-soft:#FBE7E3}
[hidden]{display:none!important}
*{box-sizing:border-box}body{margin:0;font-family:"Noto Sans TC","Microsoft JhengHei","PingFang TC",sans-serif;font-size:14px;color:var(--ink);background:var(--bg)}
header{position:sticky;top:0;background:var(--bg);padding:10px 16px;border-bottom:1px solid var(--line);z-index:5;display:flex;flex-direction:column;gap:10px}
.tabs{display:flex;gap:4px;align-items:center}
.tab{border:none;background:none;padding:6px 14px;border-radius:99px;font-weight:700;color:var(--muted)}
.tab.on{background:var(--green);color:#fff}
.tabs .sp{flex:1}
.stats{display:flex;flex-wrap:wrap;gap:8px}
.stat{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:5px 12px;display:flex;align-items:baseline;gap:6px}
.stat b{font-size:19px;font-variant-numeric:tabular-nums;color:var(--green)}
.stat.gold b{color:var(--gold)}
.stat small{color:var(--muted)}
.tools{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
input,select{font:inherit;padding:6px 9px;border:1px solid var(--line);border-radius:8px;background:#fff;color:var(--ink)}
input[type=search]{flex:1;min-width:200px}
button{font:inherit;cursor:pointer;border-radius:8px;border:1px solid var(--line);background:#fff;padding:6px 12px;color:var(--ink)}
button:hover{border-color:var(--green)}
button.primary{background:var(--green);border-color:var(--green);color:#fff}
button.danger{background:var(--red);border-color:var(--red);color:#fff}
button.ghost-danger{color:var(--red)}
button:disabled{opacity:.5;cursor:wait}
a{color:var(--green)}
.bulk{display:none;align-items:center;gap:8px;background:var(--red-soft);border:1px solid var(--red);border-radius:10px;padding:6px 10px}
.bulk.on{display:flex}
section{padding:12px 16px 28px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:10px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:10px 12px;display:flex;flex-direction:column;gap:6px}
.card.hidden-fb{background:#F1F1EE;border-style:dashed}
.card.sel{border-color:var(--red);box-shadow:0 0 0 1px var(--red)}
.card.confirm{background:var(--red-soft);border-color:var(--red)}
.top{display:flex;align-items:center;gap:8px}
.top input{width:16px;height:16px;accent-color:var(--red);padding:0}
.leaf{font-weight:700}
.time{margin-left:auto;color:var(--muted);font-size:12px}
.who{font-weight:600}
.meta{color:var(--muted);font-size:12px;word-break:break-all}
.text{white-space:pre-wrap;word-break:break-word;max-height:7.5em;overflow:auto;background:var(--bg);border-radius:8px;padding:6px 8px}
.text.empty{color:var(--muted)}
.chips{display:flex;flex-wrap:wrap;gap:6px}
.chip{font-size:12px;border-radius:99px;padding:2px 8px;background:var(--bg);border:1px solid var(--line)}
.chip.win{background:var(--gold-soft);border-color:var(--gold)}
.chip.off{background:#E4E4E0}
.chip.done{background:#E3F1E7;border-color:var(--green);color:var(--green)}
.acts{display:flex;gap:6px;justify-content:flex-end;flex-wrap:wrap;align-items:center}
.ask{color:var(--red);font-weight:700;margin-right:auto}
.empty-state{grid-column:1/-1;text-align:center;color:var(--muted);padding:40px 12px;background:var(--card);border:1px dashed var(--line);border-radius:12px}
.panel{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin-bottom:14px}
.panel h3{margin:0 0 4px;font-size:16px}
.panel .hint{margin:0 0 10px;color:var(--muted);font-size:13px}
.row{display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end}
label{display:flex;flex-direction:column;gap:4px;font-size:13px;color:var(--muted)}
table{width:100%;border-collapse:collapse}
th{text-align:left;font-size:12px;color:var(--muted);font-weight:600;padding:4px 6px;border-bottom:1px solid var(--line)}
td{padding:5px 6px;border-bottom:1px solid var(--line);vertical-align:middle}
td input{width:100%}
td.num input{width:80px;text-align:right}
.sum{margin-top:8px;font-size:13px}
.sum b{font-variant-numeric:tabular-nums}
.sum.bad{color:var(--red)}
.bar{display:flex;height:10px;border-radius:99px;overflow:hidden;background:#E9ECE6;margin-top:6px}
.bar span{display:block;height:100%}
.status{display:inline-flex;align-items:center;gap:6px;font-weight:700;padding:3px 10px;border-radius:99px}
.status.on{background:#E3F1E7;color:var(--green)}
.status.off{background:#EEE;color:var(--muted)}
.wtable td{font-size:13px}
.toast{position:fixed;left:50%;bottom:16px;transform:translateX(-50%);background:var(--ink);color:#fff;padding:8px 16px;border-radius:99px;font-size:13px;display:none;z-index:9}
.toast.on{display:block}
</style></head><body>
<header>
  <div class="tabs">
    <strong style="font-size:15px">回饋管理</strong>
    <span class="sp"></span>
    <span id="links"></span>
    <button id="refresh">重新整理</button>
  </div>
  <div class="stats" id="stats"><div class="stat"><small>讀取中…</small></div></div>
  <div class="tools">
    <input type="search" id="q" placeholder="搜尋學校、老師、Email、內容…">
    <select id="filter"><option value="all">全部回饋</option><option value="shown">公開中</option><option value="hidden">已隱藏</option><option value="photo">有照片</option></select>
  </div>
  <div class="bulk" id="bulk">
    <span id="bulkText"></span>
    <button class="danger" id="bulkDel">刪除勾選的回饋</button>
    <button id="bulkClear">取消勾選</button>
  </div>
</header>
<section><div class="grid" id="list"></div></section>
<div class="toast" id="toast"></div>
<script>
var TOKEN = '__TOKEN__';
var data = {rows: []}, selected = {}, confirming = null, busy = false, bulkConfirm = false;
function $(id){ return document.getElementById(id); }
function esc(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
function toast(msg){ var t=$('toast'); t.textContent=msg; t.className='toast on'; clearTimeout(t._h); t._h=setTimeout(function(){ t.className='toast'; }, 3500); }
function call(fn, args, ok){
  busy = true; renderBusy();
  var r = google.script.run.withSuccessHandler(function(res){ busy=false; renderBusy(); ok && ok(res); })
    .withFailureHandler(function(e){ busy=false; renderBusy(); render(); toast(e && e.message ? e.message : '操作沒有成功'); });
  r[fn].apply(r, [TOKEN].concat(args));
}
function renderBusy(){ Array.prototype.forEach.call(document.querySelectorAll('button[data-do],#bulkDel'), function(b){ b.disabled = busy; }); }
function load(){ call('adminData', [], function(d){ data=d; selected={}; confirming=null; bulkConfirm=false; render(); }); }
function key(r){ return r.row + '_' + r.ts; }

function render(){
  var rows = data.rows || [];
  var shown = rows.filter(function(r){ return !r.hidden; }).length;
  var photoN = rows.filter(function(r){ return r.photos.length; }).length;
  $('stats').innerHTML = '<div class="stat"><b>' + rows.length + '</b><small>片葉子</small></div>'
    + '<div class="stat"><b>' + shown + '</b><small>則公開中</small></div>'
    + '<div class="stat"><b>' + (rows.length - shown) + '</b><small>則已隱藏</small></div>'
    + '<div class="stat"><b>' + photoN + '</b><small>則附照片</small></div>';
  $('links').innerHTML = data.siteUrl ? '<a href="' + esc(data.siteUrl) + '" target="_blank">開啟正式網站</a>' : '';

  var q = $('q').value.trim().toLowerCase(), f = $('filter').value;
  var list = rows.slice().reverse().filter(function(r){
    if (f === 'shown' && r.hidden) return false;
    if (f === 'hidden' && !r.hidden) return false;
    if (f === 'photo' && !r.photos.length) return false;
    if (!q) return true;
    return [r.school, r.teacher, r.email, r.course, r.text].join(' ').toLowerCase().indexOf(q) >= 0;
  });
  var h = '';
  if (!rows.length) h = '<div class="empty-state">還沒有任何回饋。老師送出後，會依時間出現在這裡（最新的在最前面）。</div>';
  else if (!list.length) h = '<div class="empty-state">沒有符合條件的回饋。</div>';
  list.forEach(function(r){
    var k = key(r), isSel = !!selected[k], isConf = confirming === k;
    var chips = '';
    if (r.course) chips += '<span class="chip">' + esc(r.course) + '</span>';
    if (r.visit) chips += '<span class="chip">來園 ' + esc(r.visit) + '</span>';
    if (r.photos.length) chips += '<span class="chip">照片 ' + r.photos.length + ' 張</span>';
    if (r.hidden) chips += '<span class="chip off">內容已隱藏</span>';
    var photos = r.photos.map(function(u, i){ return '<a href="' + esc(u) + '" target="_blank">照片' + (i + 1) + '</a>'; }).join('　');
    var acts = isConf
      ? '<span class="ask">確定刪除這則回饋？</span><button class="danger" data-do="del" data-k="' + k + '">刪除</button><button data-do="cancel">取消</button>'
      : '<button data-do="hide" data-k="' + k + '">' + (r.hidden ? '恢復公開' : '隱藏內容') + '</button><button class="ghost-danger" data-do="ask" data-k="' + k + '">刪除</button>';
    h += '<div class="card' + (r.hidden ? ' hidden-fb' : '') + (isSel ? ' sel' : '') + (isConf ? ' confirm' : '') + '">'
      + '<div class="top"><input type="checkbox" data-do="sel" data-k="' + k + '"' + (isSel ? ' checked' : '') + ' aria-label="勾選第 ' + r.leaf + ' 片葉子"><span class="leaf">第 ' + r.leaf + ' 片葉子</span><span class="time">' + esc(r.time) + '</span></div>'
      + '<div class="who">' + esc(r.school) + '・' + esc(r.teacher) + '</div>'
      + '<div class="meta">' + esc(r.email) + (r.phone ? '・' + esc(r.phone) : '') + '</div>'
      + '<div class="text' + (r.text ? '' : ' empty') + '">' + (r.text ? esc(r.text) : '（沒有文字，只有照片）') + '</div>'
      + (photos ? '<div class="meta">' + photos + '</div>' : '')
      + '<div class="chips">' + chips + '</div>'
      + '<div class="acts">' + acts + '</div></div>';
  });
  $('list').innerHTML = h;
  var n = Object.keys(selected).length;
  $('bulk').className = 'bulk' + (n ? ' on' : '');
  $('bulkText').textContent = '已勾選 ' + n + ' 則。刪除後，後面的葉子編號會往前遞補，照片會移到雲端硬碟垃圾桶。';
  $('bulkDel').textContent = bulkConfirm ? '確定刪除 ' + n + ' 則' : '刪除勾選的回饋';
  renderBusy();
}

function rowByKey(k){ return (data.rows || []).filter(function(r){ return key(r) === k; })[0]; }
function doDelete(items){
  call('adminDelete', [items], function(res){
    toast('已刪除 ' + res.deleted + ' 則' + (res.photos ? '，' + res.photos + ' 張照片移到垃圾桶' : ''));
    load();
  });
}
$('list').addEventListener('click', function(e){
  var b = e.target.closest('[data-do]'); if (!b || busy) return;
  var d = b.getAttribute('data-do'), k = b.getAttribute('data-k'), r = k && rowByKey(k);
  if (d === 'sel') { if (b.checked) selected[k] = 1; else delete selected[k]; bulkConfirm = false; render(); return; }
  if (d === 'ask') { confirming = k; render(); return; }
  if (d === 'cancel') { confirming = null; render(); return; }
  if (d === 'del' && r) { doDelete([{row: r.row, ts: r.ts}]); return; }
  if (d === 'hide' && r) {
    call('adminSetHidden', [r.row, r.ts, !r.hidden], function(){ r.hidden = !r.hidden; render(); toast(r.hidden ? '已隱藏文字與照片，網站上只留下葉子' : '已恢復公開'); });
  }
});
$('bulkDel').onclick = function(){
  if (!bulkConfirm) { bulkConfirm = true; render(); return; }
  doDelete(Object.keys(selected).map(rowByKey).filter(Boolean).map(function(r){ return {row: r.row, ts: r.ts}; }));
};
$('bulkClear').onclick = function(){ selected = {}; bulkConfirm = false; render(); };
$('q').oninput = render; $('filter').onchange = render; $('refresh').onclick = load;
load();
</script></body></html>`;
