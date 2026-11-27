/**
 * 一餃匠心｜訂單串接 Google 試算表（Google Apps Script）
 *
 * 使用方式（詳見 gas/README.md）：
 * 1. 在「新的空白 Google 試算表」開啟：擴充功能 → Apps Script，貼上本檔內容。
 * 2. 執行一次 setup()，完成授權與分頁建立。
 * 3.（選用）執行 testOrder()，確認會寫入一筆測試訂單。
 * 4. 部署 → 新增部署作業 → 網頁應用程式（執行身分：我；存取權：所有人）。
 *
 * 說明：此程式為「容器綁定」，不需要寫入試算表編號，也不含任何個人資料。
 */

const CONFIG = {
  TZ: 'Asia/Taipei',
  SHEETS: { orders: '訂單', items: '訂單明細', products: '商品' },
  STATUSES: ['新訂單', '製作中', '已完成', '已取消'],
  METHODS: ['內用', '外帶', '宅配'],
  MAX_LINES: 30,        // 單筆訂單最多幾種商品
  MAX_QTY: 99,          // 單項商品最多幾份
  DUP_SECONDS: 10,      // 同一電話幾秒內不可重複送單
  MAX_PER_MINUTE: 30    // 全站每分鐘最多接收幾筆
};

const ORDER_HEADERS = ['訂單編號', '下單時間', '姓名', '電話', '取餐方式', '配送地址', '商品件數', '總金額', '狀態'];
const ITEM_HEADERS = ['訂單編號', '商品編號', '商品名稱', '單價', '數量', '小計'];
const PRODUCT_HEADERS = ['商品編號', '商品名稱', '單價', '上架'];

// 初始商品（與網站菜單一致）。之後可直接在「商品」分頁修改價格與上架狀態。
const SEED_PRODUCTS = [
  ['p1', '招牌鮮肉鍋貼', 75, true],
  ['p2', '韭菜鮮肉鍋貼', 80, true],
  ['p3', '川香辣味鍋貼', 85, true],
  ['p4', '高麗菜水餃', 75, true],
  ['p5', '鮮蝦水餃', 110, true],
  ['p6', '酸辣湯', 45, true],
  ['p7', '無糖豆漿', 30, true],
  ['p8', '涼拌小黃瓜', 40, true],
  ['p9', '滷豆干', 45, true],
  ['g1', '招牌招財鍋貼禮盒', 399, true]
];

/* ============================== 初始化 ============================== */

/** 第一次使用請執行：建立 3 個分頁、欄位、下拉選單與格式（可重複執行，不會清掉資料）。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orders = ensureSheet_(ss, CONFIG.SHEETS.orders, ORDER_HEADERS);
  const items = ensureSheet_(ss, CONFIG.SHEETS.items, ITEM_HEADERS);
  const products = ensureSheet_(ss, CONFIG.SHEETS.products, PRODUCT_HEADERS);

  // 文字欄位設為「純文字」，避免電話開頭的 0 消失，也避免被當成公式
  [3, 4, 6].forEach(function (c) { orders.getRange(2, c, 2000, 1).setNumberFormat('@'); });
  orders.getRange(2, 1, 2000, 1).setNumberFormat('@');
  orders.getRange(2, 8, 2000, 1).setNumberFormat('"NT$" #,##0');
  items.getRange(2, 1, 5000, 2).setNumberFormat('@');
  items.getRange(2, 4, 5000, 1).setNumberFormat('"NT$" #,##0');
  items.getRange(2, 6, 5000, 1).setNumberFormat('"NT$" #,##0');
  products.getRange(2, 1, 500, 1).setNumberFormat('@');
  products.getRange(2, 3, 500, 1).setNumberFormat('#,##0');

  // 訂單狀態：下拉選單＋顏色
  const statusRange = orders.getRange(2, 9, 2000, 1);
  statusRange.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STATUSES, true).setAllowInvalid(false).build()
  );
  const colors = { '新訂單': '#fff2cc', '製作中': '#fce5cd', '已完成': '#d9ead3', '已取消': '#eeeeee' };
  orders.setConditionalFormatRules(Object.keys(colors).map(function (k) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(k).setBackground(colors[k]).setRanges([statusRange]).build();
  }));

  // 商品「上架」欄位：勾選框
  products.getRange(2, 4, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

  // 商品主檔：只有在「商品編號」欄完全空白時才放入初始商品
  // （不能用 getLastRow() 判斷：「上架」欄的空白勾選框也會被算進最後一列）
  const hasProducts = products.getRange(2, 1, 500, 1).getValues().some(function (r) { return String(r[0]).trim() !== ''; });
  if (!hasProducts) {
    products.getRange(2, 1, SEED_PRODUCTS.length, 4).setValues(SEED_PRODUCTS);
  }

  // 排序：訂單、訂單明細、商品；移除空白的預設分頁
  ['Sheet1', '工作表1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  // 固定欄寬（不要用自動調寬：剛建立時只有標題列，會把欄位縮得很窄）
  setWidths_(orders, [140, 165, 95, 115, 85, 220, 85, 100, 90]);
  setWidths_(items, [140, 90, 200, 90, 70, 100]);
  setWidths_(products, [100, 200, 90, 70]);
  orders.setFrozenColumns(1); // 凍結「訂單編號」欄，左右捲動時仍看得出是哪一筆
  SpreadsheetApp.flush();
  Logger.log('初始化完成：已建立「訂單」「訂單明細」「商品」三個分頁。');
}

function setWidths_(sheet, widths) {
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#392a20').setHorizontalAlignment('center');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ============================== 接收訂單 ============================== */

/** 網站結帳會 POST 到這裡（網址 = 部署後的網頁應用程式網址）。 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return json_(handleOrder_(body));
  } catch (err) {
    console.error(err && err.stack || err);
    return json_({ ok: false, code: 'SERVER_ERROR', message: '系統忙碌中，請稍後再試。' });
  }
}

/** 用瀏覽器開啟網頁應用程式網址時，只顯示服務狀態，不會洩漏任何資料。 */
function doGet() {
  return json_({ ok: true, service: 'yijiao-orders' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function handleOrder_(body) {
  // 機器人陷阱：隱藏欄位被填寫就假裝成功，但不寫入
  if (body && body.website) return { ok: true, orderId: 'YJ-OK', total: 0 };

  const productMap = readProducts_();
  const checked = validateOrder_(body, productMap);
  if (!checked.ok) return checked;
  const o = checked.order;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, code: 'BUSY', message: '目前訂單較多，請稍後再試。' };
  try {
    const cache = CacheService.getScriptCache();
    const now = new Date();

    // 全站每分鐘上限
    const rateKey = 'rate_' + Math.floor(now.getTime() / 60000);
    const rate = Number(cache.get(rateKey) || 0);
    if (rate >= CONFIG.MAX_PER_MINUTE) return { ok: false, code: 'RATE_LIMIT', message: '目前訂單較多，請稍後再試。' };

    // 同一電話短時間內不可重複送單（防連點）
    const dupKey = 'dup_' + o.phone.replace(/\D/g, '');
    if (cache.get(dupKey)) return { ok: false, code: 'DUPLICATE', message: '請勿重複送出，請稍候再試。' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderId = nextOrderId_(now);
    const time = Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');

    const oSheet = ss.getSheetByName(CONFIG.SHEETS.orders);
    oSheet.getRange(oSheet.getLastRow() + 1, 1, 1, ORDER_HEADERS.length).setValues([[
      orderId, time, guard_(o.name), guard_(o.phone), o.method, guard_(o.address), o.count, o.total, '新訂單'
    ]]);
    const iSheet = ss.getSheetByName(CONFIG.SHEETS.items);
    const rows = o.lines.map(function (l) { return [orderId, l.id, l.name, l.price, l.qty, l.price * l.qty]; });
    iSheet.getRange(iSheet.getLastRow() + 1, 1, rows.length, ITEM_HEADERS.length).setValues(rows);

    cache.put(rateKey, String(rate + 1), 120);
    cache.put(dupKey, '1', CONFIG.DUP_SECONDS);
    return { ok: true, orderId: orderId, total: o.total };
  } finally {
    lock.releaseLock();
  }
}

/** 讀取「商品」分頁：{ 編號: { name, price, active } } */
function readProducts_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.products);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
    const id = String(r[0]).trim(), price = Number(r[2]);
    if (id && isFinite(price) && price >= 0) map[id] = { name: String(r[1]), price: price, active: r[3] === true || String(r[3]).toUpperCase() === 'TRUE' };
  });
  return map;
}

/**
 * 驗證訂單並由伺服器重新計算金額（不信任網頁傳來的價格）。
 * 純運算函式，不碰試算表，方便測試。
 */
function validateOrder_(body, productMap) {
  const fail = function (code, message) { return { ok: false, code: code, message: message }; };
  if (!body || typeof body !== 'object') return fail('BAD_REQUEST', '訂單格式不正確。');

  const name = clean_(body.name, 30);
  const phone = clean_(body.phone, 18);
  const method = String(body.method || '');
  const address = clean_(body.address, 100);
  if (!name) return fail('BAD_NAME', '請填寫姓名。');
  if (!/^[0-9+() -]{8,18}$/.test(phone)) return fail('BAD_PHONE', '電話格式不正確。');
  if (CONFIG.METHODS.indexOf(method) < 0) return fail('BAD_METHOD', '取餐方式不正確。');
  if (method === '宅配' && !address) return fail('BAD_ADDRESS', '宅配請填寫配送地址。');

  const items = Array.isArray(body.items) ? body.items : [];
  if (!items.length) return fail('EMPTY', '購物車是空的。');
  if (items.length > CONFIG.MAX_LINES) return fail('TOO_MANY', '商品種類過多。');

  const merged = {};
  for (let i = 0; i < items.length; i++) {
    const id = String(items[i] && items[i].id || '');
    const qty = Number(items[i] && items[i].qty);
    if (!productMap[id] || !productMap[id].active) return fail('BAD_ITEM', '有商品已下架或不存在，請重新整理頁面。');
    if (!Number.isInteger(qty) || qty < 1 || qty > CONFIG.MAX_QTY) return fail('BAD_QTY', '商品數量不正確。');
    merged[id] = (merged[id] || 0) + qty;
    if (merged[id] > CONFIG.MAX_QTY) return fail('BAD_QTY', '商品數量不正確。');
  }

  const lines = Object.keys(merged).map(function (id) {
    return { id: id, name: productMap[id].name, price: productMap[id].price, qty: merged[id] };
  });
  const total = lines.reduce(function (s, l) { return s + l.price * l.qty; }, 0);
  const count = lines.reduce(function (s, l) { return s + l.qty; }, 0);

  // 網頁顯示的金額和後端不一致（例如剛改過價格）：請使用者重新整理
  if (body.total !== undefined && Number(body.total) !== total) {
    return fail('PRICE_MISMATCH', '商品價格已更新，請重新整理頁面後再下單。');
  }
  return { ok: true, order: { name: name, phone: phone, method: method, address: method === '宅配' ? address : '', lines: lines, total: total, count: count } };
}

/** 去除控制字元、限制長度 */
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
}

/** 防止「公式注入」：以 = + - @ 開頭的文字前面補上單引號，試算表會當成純文字 */
function guard_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** 產生訂單編號：YJ + 日期 + 當日流水號，例如 YJ20261127-001 */
function nextOrderId_(now) {
  const day = Utilities.formatDate(now, CONFIG.TZ, 'yyyyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = 'seq_' + day;
  const seq = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(seq));
  return 'YJ' + day + '-' + ('00' + seq).slice(-3);
}

/* ============================== 測試 ============================== */

/** 在編輯器執行：模擬網站送出一筆訂單，檢查是否寫入（會真的新增一筆，可手動刪除）。 */
function testOrder() {
  const map = readProducts_();
  const total = map.p1.price * 2 + map.p6.price;
  const res = handleOrder_({
    items: [{ id: 'p1', qty: 2 }, { id: 'p6', qty: 1 }],
    name: '測試顧客', phone: '0912345678', method: '外帶', address: '', total: total, website: ''
  });
  Logger.log(JSON.stringify(res));
  return res;
}
