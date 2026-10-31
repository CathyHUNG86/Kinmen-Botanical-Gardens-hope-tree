/**
 * 金門植物園希望樹 — Google Apps Script 後端
 *
 * 試算表分頁：
 *   回饋：每一列是一則回饋（＝一片葉子）。同一位老師（以 Email 辨識）同一堂課最多 2 則，
 *         且文字與照片都不能和自己之前送出的重複。送出後立即長出葉子，內容也立即公開（不含姓名與聯絡方式）。
 *         若有不適合公開的內容，把「審核狀態」改成「不公開」即可隱藏文字，葉子仍保留。
 *   獎項：獎項名稱／總數量／剩餘數量／中獎機率(%)。機率加總可小於 100，其餘為未中獎。
 *   設定：活動開始日期、活動結束日期、每位教師抽獎次數。
 * 照片存在雲端硬碟「希望樹回饋照片」資料夾（僅您可見，不公開）。
 */

const SHEET_FEEDBACK = '回饋';
const SHEET_PRIZE = '獎項';
const SHEET_SETTING = '設定';
const HEADERS = ['送出時間', '葉子編號', '學校', '教師姓名', 'Email', '電話', '來園年月', '參與課程',
  '回饋內容', '照片連結', '同意公開', '審核狀態', '抽獎結果', '照片指紋'];
const COL = { school: 2, course: 7, visit: 6, email: 4, text: 8, photo: 9, isPublic: 10, status: 11, prize: 12, hash: 13 };
const MAX_PER_COURSE = 2; // 每位老師同一堂課程最多回饋次數

/** 第一次使用：在編輯器選 setup 後按「執行」 */
function setup() {
  const ss = SpreadsheetApp.getActive();
  if (/^未命名|^Untitled/.test(ss.getName())) ss.rename('金門植物園希望樹｜回饋資料');
  let fb = ss.getSheetByName(SHEET_FEEDBACK) || ss.insertSheet(SHEET_FEEDBACK);
  if (fb.getLastRow() === 0) { fb.appendRow(HEADERS); fb.setFrozenRows(1); }
  fb.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold'); // 舊版表頭也一併補上新欄位

  let pz = ss.getSheetByName(SHEET_PRIZE) || ss.insertSheet(SHEET_PRIZE);
  if (pz.getLastRow() === 0) {
    pz.appendRow(['獎項名稱', '總數量', '剩餘數量', '中獎機率(%)']);
    pz.appendRow(['（請改成實際獎品）獎品A', 20, 20, 15]);
    pz.appendRow(['（請改成實際獎品）獎品B', 40, 40, 25]);
    pz.setFrozenRows(1); pz.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  let st = ss.getSheetByName(SHEET_SETTING) || ss.insertSheet(SHEET_SETTING);
  if (st.getLastRow() === 0) {
    st.appendRow(['項目', '值']);
    st.appendRow(['活動開始日期', new Date(2026, 10, 1)]);
    st.appendRow(['活動結束日期', new Date(2026, 11, 31, 23, 59)]);
    st.appendRow(['每位教師抽獎次數', 1]);
    st.getRange(1, 1, 1, 2).setFontWeight('bold');
  }

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

function getSettings_() {
  const st = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTING);
  const map = {};
  st.getDataRange().getValues().slice(1).forEach(r => map[r[0]] = r[1]);
  const now = new Date();
  const start = map['活動開始日期'] ? new Date(map['活動開始日期']) : null;
  const end = map['活動結束日期'] ? new Date(map['活動結束日期']) : null;
  return {
    open: (!start || now >= start) && (!end || now <= end),
    drawLimit: Number(map['每位教師抽獎次數']) || 1
  };
}

/** 前端讀取整棵樹：只回傳公開需要的欄位，不含姓名與聯絡方式 */
function getTree() {
  const fb = SpreadsheetApp.getActive().getSheetByName(SHEET_FEEDBACK);
  const rows = fb.getLastRow() > 1 ? fb.getRange(2, 1, fb.getLastRow() - 1, HEADERS.length).getValues() : [];
  const leaves = rows.map((r, i) => ({
    n: i + 1,
    school: String(r[COL.school]),
    course: String(r[COL.course]),
    month: String(r[COL.visit]),
    text: r[COL.status] === '不公開' ? '' : String(r[COL.text]),
    hasPhoto: !!r[COL.photo]
  }));
  const s = getSettings_();
  return { leaves, campaign: { open: s.open, drawLimit: s.drawLimit } };
}

/** 收到回饋：存資料、存照片、抽獎，回傳葉子編號與抽獎結果 */
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
    const prize = draw_(fb, email);

    fb.appendRow([new Date(), leafNo, school, teacher, email, clean(p.phone, 20), clean(p.visit, 10), course,
      text, links.join('\n'), '是', '已公開', prize.record, hashes.join(',')]);

    return { leafNo, prize: { drawn: prize.drawn, won: prize.won, name: prize.name, message: prize.message } };
  } finally {
    lock.releaseLock();
  }
}

/** 抽獎：檢查活動期間與次數，依機率抽，並扣庫存 */
function draw_(fb, email) {
  const s = getSettings_();
  if (!s.open) return { drawn: false, record: '非抽獎期間', message: '目前不在抽獎期間，謝謝您的回饋，這片葉子已經長上希望樹。' };

  const prev = fb.getLastRow() > 1 ? fb.getRange(2, 1, fb.getLastRow() - 1, HEADERS.length).getValues() : [];
  const used = prev.filter(r => String(r[COL.email]).toLowerCase() === email &&
    r[COL.prize] && !['非抽獎期間', '已達抽獎次數'].includes(r[COL.prize])).length;
  if (used >= s.drawLimit) return { drawn: false, record: '已達抽獎次數', message: '您已參加過抽獎，這次的回饋一樣長成新葉子了，謝謝您！' };

  const pz = SpreadsheetApp.getActive().getSheetByName(SHEET_PRIZE);
  const list = pz.getLastRow() > 1 ? pz.getRange(2, 1, pz.getLastRow() - 1, 4).getValues() : [];
  let roll = Math.random() * 100, acc = 0;
  for (let i = 0; i < list.length; i++) {
    const [name, , left, pct] = list[i];
    if (!name || Number(left) <= 0) continue;
    acc += Number(pct) || 0;
    if (roll < acc) {
      pz.getRange(i + 2, 3).setValue(Number(left) - 1);
      return { drawn: true, won: true, name: String(name), record: '中獎：' + name };
    }
  }
  return { drawn: true, won: false, record: '未中獎' };
}

/** 比對用：去掉空白、標點與大小寫差異 */
function normText_(t) {
  return String(t || '').toLowerCase().replace(/[\s\p{P}\p{S}]/gu, '');
}

function md5_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
