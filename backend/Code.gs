/**
 * Paw-To-Go 毛孩快充站後端：Google 試算表 + Apps Script 網頁應用程式
 * 貼到試算表的「擴充功能 → Apps Script」，先執行一次 setup()，再部署為網頁應用程式。
 */

// 與前端 config.js 的 GOOGLE_CLIENT_ID 相同；留空則停用會員登入
const GOOGLE_CLIENT_ID = '';

const SHEET = { products: '商品', orders: '訂單', members: '會員' };
const PICKUPS = ['台南東區旗艦吧台取餐', '東區超充站機台取貨'];
const MEMBER_DISCOUNT = 0.1;
const MAX_ORDERS = 5000;

const PRODUCT_HEAD = ['商品編號', '銷售通路', '商品系列', '商品名稱', '核心配方/規格', '適用對象', '訴求', '定價', '預估食材成本', '圖示', '顏色', '上架'];
const ORDER_HEAD = ['訂單編號', '時間', '會員 Email', '會員姓名', '毛孩', '取貨方式', '品項明細', '小計', '折扣', '合計', '狀態'];
const MEMBER_HEAD = ['Google ID', 'Email', '姓名', '毛孩名字', '毛孩體重(kg)', '建立時間', '最後登入', '訂單數'];

// 成本欄留空：請直接在試算表「商品」工作表填入預估食材成本
const SEED = [
  ['P01', '自動販賣機', '單品', '汪汪電解水', '200ml 袋裝（低鈉等滲透壓）', '狗狗', '奔跑後補水', 35, '', 'bottle', '#4f9ec4'],
  ['P02', '自動販賣機', '單品', '純肉凍乾小酥包', '15g 夾鏈隨身包（純肉原粒）', '狗狗', '戶外互動獎勵', 50, '', 'pouch', '#b5533c'],
  ['P03', '自動販賣機', '單品', '舒壓鮮肉泥', '15g 單支條裝', '狗狗', '車內安撫', 30, '', 'sticks', '#c96b8a'],
  ['P04', '自動販賣機', '單品', '外出救急清潔組', '潔足濕巾 5 抽＋便袋 2 枚', '狗狗', '上車前擦腳', 35, '', 'pouch', '#5aa9a3'],
  ['P05', '自動販賣機', '組合包', '放電組', '電解水 1 包＋純肉凍乾 1 包', '狗狗', '現省 $10（原價 $85）', 75, '', 'box', '#e3a53c'],
  ['P06', '自動販賣機', '組合包', '上車組', '電解水 1 包＋救急清潔組 1 包', '狗狗', '現省 $11（原價 $70）', 59, '', 'box', '#6fae7a'],
  ['P07', '實體店', '主食單點', '鮮嫩雞肉能量碗', '150g（熟雞肉丁＋根莖時蔬）', '狗狗', '低敏低脂好消化', 129, '', 'bowl', '#e3a53c'],
  ['P08', '實體店', '主食單點', '原汁牛肉能量碗', '150g（熟牛絞肉＋根莖時蔬）', '狗狗', '活動量大補鐵', 129, '', 'bowl', '#b5533c'],
  ['P09', '實體店', '主食單點', '低敏櫻桃鴨能量碗', '150g（櫻桃鴨肉丁＋根莖時蔬）', '狗狗', '換季低敏', 129, '', 'bowl', '#b97a95'],
  ['P10', '實體店', '湯飲單點', '現溫大骨湯杯', '250ml 外帶溫熱杯', '狗狗', '挑嘴補水', 49, '', 'cup', '#d98b4a'],
  ['P11', '實體店', '湯飲加購', '現溫大骨湯杯（主餐加購）', '購買能量碗隨單加購優惠', '狗狗', '加購價', 39, '', 'cup', '#a98467'],
  ['P12', '實體店', '湯飲多日', '大骨湯家庭瓶', '500ml 冷藏隨行瓶', '狗狗', '分 3～4 天拌糧', 99, '', 'bottle', '#d98b4a'],
  ['P13', '實體店', '組合', '免開火輕便組', '能量碗任選 1 份＋舒壓肉泥 1 支', '狗狗', '現省 $14（原價 $159）', 145, '', 'box', '#e07a5f'],
  ['P14', '實體店', '組合', '雙犬週末加菜組', '能量碗任選 2 份＋骨湯家庭瓶 1 瓶', '狗狗', '現省 $38（原價 $357）', 319, '', 'box', '#5b9bd5']
];

/** 第一次使用時在編輯器手動執行：建立三張工作表並寫入 14 個商品。重複執行不會覆蓋既有資料。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const make = (name, head) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(head);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
    }
    return sh;
  };
  const p = make(SHEET.products, PRODUCT_HEAD);
  if (p.getLastRow() === 1) {
    p.getRange(2, 1, SEED.length, PRODUCT_HEAD.length).setValues(SEED.map(r => r.concat([true])));
    p.getRange(2, PRODUCT_HEAD.length, SEED.length, 1).insertCheckboxes();
  }
  make(SHEET.orders, ORDER_HEAD);
  make(SHEET.members, MEMBER_HEAD);
}

function doGet(e) {
  return handle_(() => {
    const action = (e && e.parameter && e.parameter.action) || 'products';
    if (action === 'products') return products_();
    throw new Error('未知的操作');
  });
}

function doPost(e) {
  return handle_(() => {
    const b = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (b.action === 'login') return login_(b);
    if (b.action === 'savePet') return savePet_(b);
    if (b.action === 'order') return order_(b);
    throw new Error('未知的操作');
  });
}

function handle_(fn) {
  let out;
  try {
    out = Object.assign({ ok: true }, fn());
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('找不到工作表「' + name + '」，請先執行 setup()');
  return sh;
}

/** 去除控制字元並避免試算表公式注入 */
function clean_(s, max) {
  return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, '').trim().replace(/^[=+\-@]+/, '').slice(0, max);
}

function readProducts_() {
  return sheet_(SHEET.products).getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '' && r[11] === true && Number(r[7]) > 0)
    .map(r => ({
      sku: String(r[0]), channel: String(r[1]), series: String(r[2]), name: String(r[3]), spec: String(r[4]),
      pet: String(r[5]), need: String(r[6]), price: Number(r[7]), cost: Number(r[8]) || 0,
      art: String(r[9]), color: String(r[10])
    }));
}

function products_() {
  const list = readProducts_();
  const costed = list.filter(p => p.cost > 0);
  const margin = costed.length ? costed.reduce((s, p) => s + (p.price - p.cost) / p.price, 0) / costed.length : 0;
  // 成本不對外提供，只回傳平均毛利率
  return { products: list.map(p => { const o = Object.assign({}, p); delete o.cost; return o; }), avgMargin: margin };
}

/** 向 Google 驗證登入憑證，回傳 {sub,email,name} */
function verify_(idToken) {
  if (!GOOGLE_CLIENT_ID) throw new Error('後端尚未設定 GOOGLE_CLIENT_ID');
  if (!idToken) throw new Error('請先登入');
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('登入已過期，請重新登入');
  const t = JSON.parse(res.getContentText());
  const issOk = t.iss === 'accounts.google.com' || t.iss === 'https://accounts.google.com';
  if (t.aud !== GOOGLE_CLIENT_ID || !issOk || Number(t.exp) * 1000 < Date.now()) throw new Error('登入驗證失敗');
  return { sub: String(t.sub), email: clean_(t.email, 120), name: clean_(t.name || t.email, 60) };
}

/** 找到或建立會員列，回傳列號（需在鎖內呼叫） */
function memberRow_(sh, u) {
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
  const i = ids.indexOf(u.sub);
  const now = new Date();
  if (i >= 0) {
    sh.getRange(i + 2, 2, 1, 2).setValues([[u.email, u.name]]);
    sh.getRange(i + 2, 7).setValue(now);
    return i + 2;
  }
  sh.appendRow([u.sub, u.email, u.name, '', '', now, now, 0]);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('@');
  return sh.getLastRow();
}

function memberOut_(sh, row) {
  const r = sh.getRange(row, 1, 1, MEMBER_HEAD.length).getValues()[0];
  return { email: r[1], name: r[2], petName: r[3], petWeight: r[4] === '' ? null : Number(r[4]), orders: Number(r[7]) || 0 };
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function login_(b) {
  const u = verify_(b.idToken);
  return withLock_(() => {
    const sh = sheet_(SHEET.members);
    return { member: memberOut_(sh, memberRow_(sh, u)) };
  });
}

function savePet_(b) {
  const u = verify_(b.idToken);
  const petName = clean_(b.petName, 20);
  const w = Number(b.petWeight);
  if (!petName) throw new Error('請填寫毛孩名字');
  if (!(w >= 1 && w <= 90)) throw new Error('體重請填 1 到 90 公斤');
  return withLock_(() => {
    const sh = sheet_(SHEET.members);
    const row = memberRow_(sh, u);
    sh.getRange(row, 4, 1, 2).setValues([[petName, Math.round(w * 10) / 10]]);
    return { member: memberOut_(sh, row) };
  });
}

function order_(b) {
  if (PICKUPS.indexOf(b.pickup) < 0) throw new Error('取貨方式不正確');
  if (!Array.isArray(b.items) || !b.items.length || b.items.length > 30) throw new Error('購物車內容不正確');
  const bySku = {};
  readProducts_().forEach(p => { bySku[p.sku] = p; });
  let subtotal = 0;
  const lines = b.items.map(it => {
    const p = bySku[String(it.sku)];
    const qty = Number(it.qty);
    if (!p) throw new Error('有商品已下架，請重新整理頁面');
    if (!Number.isInteger(qty) || qty < 1 || qty > 20) throw new Error('單一品項數量請在 1 到 20 之間');
    subtotal += p.price * qty;
    return p.name + ' ×' + qty;
  });
  // 價格一律以試算表為準，不採用前端傳來的金額
  const u = b.idToken ? verify_(b.idToken) : null;
  const discount = u ? Math.round(subtotal * MEMBER_DISCOUNT) : 0;
  const total = subtotal - discount;

  return withLock_(() => {
    const sh = sheet_(SHEET.orders);
    const n = sh.getLastRow();
    if (n > MAX_ORDERS) throw new Error('展示訂單已達上限');
    let pet = '';
    if (u) {
      const ms = sheet_(SHEET.members);
      const row = memberRow_(ms, u);
      const cell = ms.getRange(row, 8);
      cell.setValue((Number(cell.getValue()) || 0) + 1);
      pet = ms.getRange(row, 4).getValue();
    }
    const now = new Date();
    const orderId = 'PTG-' + Utilities.formatDate(now, 'Asia/Taipei', 'yyMMdd') + '-' + ('000' + n).slice(-4);
    sh.appendRow([orderId, now, u ? u.email : '訪客', u ? u.name : '', pet, b.pickup, lines.join('、'), subtotal, discount, total, '展示訂單']);
    return { orderId: orderId, subtotal: subtotal, discount: discount, total: total };
  });
}
