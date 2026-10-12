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

const SEED = [
  ['P01', '實體店', '營養湯飲', '黃金牛骨元氣水', '牛大骨、胡蘿蔔、蘋果', '狗狗', '補充元氣', 85, 30, 'cup', '#d98b4a'],
  ['P02', '實體店', '營養湯飲', '小魚干喵喵膠原飲', '丁香魚、雞軟骨、柴魚高湯', '貓咪', '皮毛養護', 80, 25, 'cup', '#5b9bd5'],
  ['P03', '實體店', '營養湯飲', '關節保養綠蓉湯', '綠唇貽貝、雞胸肉、花椰菜', '犬貓', '關節保健', 95, 35, 'cup', '#6fae7a'],
  ['P04', '實體店', '鮮食主餐', '嫩蒸土雞藜麥杯', '雞腿肉、紅藜麥、南瓜', '狗狗', '低脂健康', 120, 45, 'bowl', '#e3a53c'],
  ['P05', '實體店', '鮮食主餐', '鮮鮭牛腩雙拼能量碗', '鮭魚、牛腩、青豆', '狗狗', '高蛋白增肌', 145, 55, 'bowl', '#e07a5f'],
  ['P06', '實體店', '鮮食主餐', '慕斯鴨胸舒敏盒', '鴨胸肉、山藥、枸杞', '貓咪', '低敏易消化', 130, 45, 'bowl', '#b97a95'],
  ['P07', '實體店', '特色甜點', '汪汪雲朵羊奶霜', '羊奶、無糖優格、角豆粉', '狗狗', '腸胃保健', 90, 30, 'dessert', '#a98467'],
  ['P08', '實體店', '特色甜點', '肉泥珍珠啵啵杯', '雞肉泥、寒天珍珠、雞高湯', '貓咪', '補水解饞', 110, 35, 'dessert', '#c96b8a'],
  ['P09', '自動販賣機', '凍乾零食', '骰子牛凍乾原肉塊', '100% 原肉／50g', '犬貓', '高蛋白零食', 150, 60, 'pouch', '#b5533c'],
  ['P10', '自動販賣機', '鮮食包', '雞肉鮮蔬常溫餐包', '雞肉、蔬菜／100g', '狗狗', '外出便攜', 70, 25, 'pouch', '#7fb069'],
  ['P11', '自動販賣機', '肉泥條', '鮪魚干貝綜合肉泥條', '鮪魚、干貝／4 入', '貓咪', '誘食獎勵', 60, 20, 'sticks', '#4f9ec4'],
  ['P12', '自動販賣機', '機能水', '蔓越莓泌尿保健水', '蔓越莓萃取物／250ml', '貓咪', '泌尿道保健', 45, 15, 'bottle', '#c2456b'],
  ['P13', '自動販賣機', '凍乾零食', '爆蛋柳葉魚凍乾', '柳葉魚／40g', '犬貓', '豐富鈣質', 120, 45, 'pouch', '#d9a441'],
  ['P14', '自動販賣機', '潔牙骨', '綠茶薄荷潔牙棒', '綠茶粉、薄荷／5 入', '狗狗', '清新口氣', 80, 25, 'sticks', '#5aa9a3']
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
  const margin = list.length ? list.reduce((s, p) => s + (p.price - p.cost) / p.price, 0) / list.length : 0;
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
