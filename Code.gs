/**
 * ร้านกรูตำเอง — Backend (Google Apps Script) ใช้ Google Sheet เป็นฐานข้อมูล
 *
 * วิธีติดตั้ง
 *  1) เปิด Google Sheet ใหม่ > ส่วนขยาย > Apps Script > วางโค้ดนี้ทั้งไฟล์แทนของเดิม
 *  2) เลือกฟังก์ชัน setup แล้วกด Run (อนุญาตสิทธิ์ครั้งแรก)
 *  3) ไปที่ Project Settings > Script Properties แก้ ADMIN_PIN และ RIDER_PIN ให้เป็นรหัสของคุณเอง
 *  4) Deploy > New deployment > Web app
 *       Execute as: Me   |   Who has access: Anyone
 *  5) คัดลอก URL (ลงท้ายด้วย /exec) ไปใส่ตัวแปร API_URL ในไฟล์ HTML ทั้ง 4 หน้า
 *
 * เมื่อแก้โค้ดนี้ภายหลัง ต้อง Deploy > Manage deployments > Edit > New version ทุกครั้ง
 */

const TZ = 'Asia/Bangkok';

// หัวคอลัมน์ของแต่ละชีต (แถวแรก) — ลำดับตรงนี้คือลำดับคอลัมน์ในชีต
const SCHEMA = {
  Orders: ['id', 'timestamp', 'customer', 'phone', 'type', 'table', 'items', 'total',
           'note', 'status', 'payment', 'rating', 'riderNote', 'deliveryPhoto', 'deliveredAt'],
  Menu:   ['id', 'name', 'category', 'price', 'image'],
  Config: ['key', 'value'],
  Users:  ['phone', 'name', 'passHash', 'salt', 'address', 'lat', 'lng', 'createdAt']
};
// คอลัมน์ที่ต้องเก็บเป็นข้อความล้วน (กัน Sheets แปลงเบอร์โทรตัดเลข 0 หรือแปลงเวลาเป็นวันที่)
const TEXT_COLS = {
  Orders: ['id', 'timestamp', 'phone', 'deliveredAt'],
  Menu:   ['id'],
  Config: ['key', 'value'],
  Users:  ['phone', 'createdAt']
};

/* ---------------------------------------------------------------- setup */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  if (!props.getProperty('TOKEN_SECRET')) props.setProperty('TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!props.getProperty('ADMIN_PIN')) props.setProperty('ADMIN_PIN', '1234');
  if (!props.getProperty('RIDER_PIN')) props.setProperty('RIDER_PIN', '5678');

  Object.keys(SCHEMA).forEach(name => {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const head = SCHEMA[name];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#fde68a');
    sh.setFrozenRows(1);
    TEXT_COLS[name].forEach(c => {
      const idx = head.indexOf(c) + 1;
      sh.getRange(1, idx, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  });
  const first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);

  if (!readAll_('Config').length) {
    writeConfig_({
      shopName: 'ร้านส้มตำ ไก่ทอด หมูทอด ข้าวกระเพรา', shopPhone: '081-394-8240',
      bankName: 'กสิกรไทย', accName: 'ร้านส้มตำ', accNum: '123-4-56789-0',
      promptpayName: 'ร้านส้มตำ', promptpayNum: '081-394-8240', qrImage: '', logo: ''
    });
  }
  if (!readAll_('Menu').length) {
    [['ส้มตำไทย', 'ส้มตำ', 50, '🥗'], ['ส้มตำปูปลาร้า', 'ส้มตำ', 50, '🌶️'],
     ['ไก่ทอดหาดใหญ่', 'ของทอด/ทานเล่น', 45, '🍗'], ['ข้าวกระเพราหมูกรอบ', 'อาหารจานเดียว', 60, '🍛']
    ].forEach((m, i) => appendRow_('Menu', { id: String(Date.now() + i), name: m[0], category: m[1], price: m[2], image: m[3] }));
  }
  console.log('ตั้งค่าเสร็จแล้ว — อย่าลืมเปลี่ยน ADMIN_PIN / RIDER_PIN ใน Script Properties');
}

/* ------------------------------------------------------------ entrypoints */

// GET = ข้อมูลสาธารณะสำหรับหน้าลูกค้า (เมนู + ตั้งค่าร้าน) ไม่มีข้อมูลออเดอร์/ลูกค้า
function doGet() {
  try {
    return json_({ ok: true, menu: readAll_('Menu').map(menuOut_), config: readConfig_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = ACTIONS[p.action];
    if (!fn) throw new Error('ไม่รู้จักคำสั่ง: ' + p.action);
    // คำสั่งที่แก้ข้อมูลต้องต่อคิวผ่านล็อก กันออเดอร์/แถวชนกัน; คำสั่งอ่านอย่างเดียวไม่ต้องรอ
    if (!READ_ONLY[p.action] && !lock.tryLock(20000)) throw new Error('ระบบกำลังประมวลผล กรุณาลองใหม่อีกครั้ง');
    return json_({ ok: true, data: fn(p) });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

const READ_ONLY = { login: 1, orderStatus: 1, getOrders: 1, adminLogin: 1, riderLogin: 1 };

const ACTIONS = {
  // ---- ลูกค้า
  register: p => {
    const phone = phone_(p.phone);
    if (!p.name || !String(p.name).trim()) throw new Error('กรุณากรอกชื่อ');
    if (String(p.password || '').length < 4) throw new Error('รหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร');
    if (findRow_('Users', 'phone', phone)) throw new Error('เบอร์โทรศัพท์นี้ลงทะเบียนไปแล้ว');
    const salt = Utilities.getUuid();
    appendRow_('Users', {
      phone: phone, name: String(p.name).trim(), passHash: hashPw_(p.password, salt), salt: salt,
      address: p.address || '', lat: num_(p.lat), lng: num_(p.lng), createdAt: nowIso_()
    });
    return session_('customer', findRow_('Users', 'phone', phone).obj);
  },
  login: p => {
    const phone = phone_(p.phone);
    throttle_('login:' + phone);
    const u = findRow_('Users', 'phone', phone);
    if (!u || u.obj.passHash !== hashPw_(p.password, u.obj.salt)) throw new Error('เบอร์โทรศัพท์หรือรหัสผ่านไม่ถูกต้อง');
    clearThrottle_('login:' + phone);
    return session_('customer', u.obj);
  },
  updateUser: p => {
    const me = auth_(p.token, ['customer']);
    const row = findRow_('Users', 'phone', me.sub);
    if (!row) throw new Error('ไม่พบบัญชีผู้ใช้');
    const u = row.obj;
    const newPhone = p.user && p.user.phone ? phone_(p.user.phone) : u.phone;
    if (newPhone !== u.phone && findRow_('Users', 'phone', newPhone)) throw new Error('เบอร์โทรศัพท์นี้ถูกใช้แล้ว');
    const nu = p.user || {};
    u.phone = newPhone;
    if (nu.name) u.name = String(nu.name).trim();
    if (nu.address) u.address = nu.address;
    if (nu.coords) { u.lat = num_(nu.coords.lat); u.lng = num_(nu.coords.lng); }
    if (nu.password) {
      if (String(nu.password).length < 4) throw new Error('รหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร');
      u.salt = Utilities.getUuid();
      u.passHash = hashPw_(nu.password, u.salt);
    }
    writeRow_('Users', row.index, u);
    return session_('customer', u);
  },
  addOrder: p => {
    const me = auth_(p.token, ['customer']);
    const u = findRow_('Users', 'phone', me.sub);
    if (!u) throw new Error('ไม่พบบัญชีผู้ใช้ กรุณาเข้าสู่ระบบใหม่');
    const o = p.order || {};
    const menu = readAll_('Menu');
    // คำนวณราคาจากเมนูในชีตเอง ไม่เชื่อราคา/ยอดรวมที่ส่งมาจากเบราว์เซอร์
    const items = (o.items || []).map(x => {
      const m = menu.find(y => String(y.id) === String(x.id));
      if (!m) throw new Error('มีเมนูที่ไม่พร้อมขายในตะกร้า กรุณารีเฟรชหน้าเว็บแล้วสั่งใหม่');
      const qty = Math.floor(Number(x.qty));
      if (!(qty > 0 && qty <= 99)) throw new Error('จำนวนสินค้าไม่ถูกต้อง');
      return { id: String(m.id), name: m.name, price: Number(m.price), qty: qty };
    });
    if (!items.length) throw new Error('ยังไม่ได้เลือกเมนูอาหาร');
    const total = items.reduce((s, i) => s + i.price * i.qty, 0);
    const pay = o.payment === 'เก็บปลายทาง' ? 'เก็บปลายทาง' : 'รอชำระเงิน';
    const order = {
      id: nextOrderId_(), timestamp: nowIso_(),
      customer: u.obj.name + ' (' + u.obj.phone + ')', phone: u.obj.phone,
      type: o.type || 'ทานที่ร้าน', table: o.table || '-', items: JSON.stringify(items), total: total,
      note: String(o.note || '').slice(0, 1000), status: 'กำลังรอ', payment: pay,
      rating: '', riderNote: '', deliveryPhoto: '', deliveredAt: ''
    };
    appendRow_('Orders', order);
    return { id: order.id, total: total };
  },
  orderStatus: p => {                       // ลูกค้าเช็กสถานะออเดอร์ของตัวเอง
    const me = auth_(p.token, ['customer']);
    const r = findRow_('Orders', 'id', p.orderId);
    if (!r || r.obj.phone !== me.sub) throw new Error('ไม่พบออเดอร์');
    return { id: r.obj.id, status: r.obj.status, payment: r.obj.payment, rating: r.obj.rating };
  },
  notifyPaid: p => {
    const me = auth_(p.token, ['customer']);
    const r = findRow_('Orders', 'id', p.orderId);
    if (!r || r.obj.phone !== me.sub) throw new Error('ไม่พบออเดอร์');
    if (r.obj.payment === 'รอชำระเงิน') { r.obj.payment = 'แจ้งโอนแล้ว'; writeRow_('Orders', r.index, r.obj); }
    return true;
  },
  rateOrder: p => {
    const me = auth_(p.token, ['customer']);
    const r = findRow_('Orders', 'id', p.orderId);
    if (!r || r.obj.phone !== me.sub) throw new Error('ไม่พบออเดอร์');
    const stars = Math.round(Number(p.rating));
    if (!(stars >= 1 && stars <= 5)) throw new Error('คะแนนไม่ถูกต้อง');
    r.obj.rating = stars;
    writeRow_('Orders', r.index, r.obj);
    return true;
  },

  // ---- แอดมิน / ไรเดอร์
  adminLogin: p => {
    throttle_('adminLogin');
    if (String(p.pin || '') !== prop_('ADMIN_PIN')) throw new Error('รหัสผ่านไม่ถูกต้อง');
    clearThrottle_('adminLogin');
    return { token: makeToken_('admin', 'admin', 12 * 3600) };
  },
  riderLogin: p => {
    throttle_('riderLogin');
    if (String(p.pin || '') !== prop_('RIDER_PIN')) throw new Error('รหัสผ่านไม่ถูกต้อง');
    clearThrottle_('riderLogin');
    return { token: makeToken_('rider', 'rider', 12 * 3600) };
  },
  getOrders: p => {
    const me = auth_(p.token, ['admin', 'rider']);
    let list = readAll_('Orders').map(orderOut_);
    if (me.role === 'rider') {               // ไรเดอร์เห็นเฉพาะงานส่งที่ถึงขั้น "กำลังขนส่ง"
      list = list.filter(o => o.type === 'จัดส่งถึงบ้าน' && o.status === 'กำลังขนส่ง');
    }
    return { orders: list.slice(-1000) };
  },
  updateStatus: p => {
    auth_(p.token, ['admin']);
    const r = findRow_('Orders', 'id', p.orderId);
    if (!r) throw new Error('ไม่พบออเดอร์');
    if (p.status) r.obj.status = String(p.status);
    if (p.payment) r.obj.payment = String(p.payment);
    writeRow_('Orders', r.index, r.obj);
    return true;
  },
  riderComplete: p => {
    auth_(p.token, ['rider', 'admin']);
    const r = findRow_('Orders', 'id', p.orderId);
    if (!r) throw new Error('ไม่พบออเดอร์');
    if (p.photoBase64) r.obj.deliveryPhoto = saveImage_(p.photoBase64, 'delivery-' + r.obj.id);
    r.obj.riderNote = String(p.riderNote || '').slice(0, 500);
    r.obj.status = 'เสร็จสิ้น';
    r.obj.deliveredAt = nowIso_();
    if (r.obj.payment === 'เก็บปลายทาง') r.obj.payment = 'ชำระเงินแล้ว';
    writeRow_('Orders', r.index, r.obj);
    return true;
  },
  deleteOrders: p => {
    auth_(p.token, ['admin']);
    const ids = (p.orderIds || []).map(String);
    const sh = sheet_('Orders');
    const rows = readAll_('Orders');
    for (let i = rows.length - 1; i >= 0; i--) {     // ลบจากล่างขึ้นบน กัน index เลื่อน
      if (ids.indexOf(String(rows[i].id)) >= 0) sh.deleteRow(i + 2);
    }
    return true;
  },
  updateConfig: p => {
    auth_(p.token, ['admin']);
    const cfg = Object.assign({}, p.config || {});
    ['logo', 'qrImage'].forEach(k => {
      if (cfg[k] && String(cfg[k]).indexOf('data:') === 0) cfg[k] = saveImage_(cfg[k], k);
    });
    writeConfig_(cfg);
    return readConfig_();
  },
  updateMenu: p => {
    auth_(p.token, ['admin']);
    const list = (p.menu || []).map(m => ({
      id: String(m.id), name: String(m.name || '').trim(), category: m.category || '',
      price: Number(m.price) || 0,
      image: m.image && String(m.image).indexOf('data:') === 0 ? saveImage_(m.image, 'menu-' + m.id) : (m.image || '')
    }));
    const sh = sheet_('Menu');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, SCHEMA.Menu.length).clearContent();
    if (list.length) {
      sh.getRange(2, 1, list.length, SCHEMA.Menu.length)
        .setValues(list.map(m => SCHEMA.Menu.map(c => m[c])));
    }
    return list;                              // ส่งกลับให้หน้าแอดมินแทนที่รูป base64 ด้วย URL
  }
};

/* ------------------------------------------------------ sheet helpers */

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function sheet_(name) {
  const sh = SpreadsheetApp.openById(prop_('SPREADSHEET_ID')).getSheetByName(name);
  if (!sh) throw new Error('ยังไม่ได้รัน setup() — ไม่พบชีต ' + name);
  return sh;
}
function readAll_(name) {
  const sh = sheet_(name), head = SCHEMA[name];
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues()
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => { const o = {}; head.forEach((c, i) => o[c] = r[i]); return o; });
}
function rowOf_(name, obj) { return SCHEMA[name].map(c => obj[c] === undefined || obj[c] === null ? '' : obj[c]); }
function appendRow_(name, obj) {
  const sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, 1, SCHEMA[name].length).setValues([rowOf_(name, obj)]);
}
function findRow_(name, col, val) {
  const rows = readAll_(name);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][col]) === String(val)) return { index: i + 2, obj: rows[i] };
  }
  return null;
}
function writeRow_(name, index, obj) {
  sheet_(name).getRange(index, 1, 1, SCHEMA[name].length).setValues([rowOf_(name, obj)]);
}
function readConfig_() {
  const c = {};
  readAll_('Config').forEach(r => c[r.key] = r.value);
  return c;
}
function writeConfig_(cfg) {
  const cur = readConfig_();
  Object.keys(cfg).forEach(k => cur[k] = cfg[k] === null || cfg[k] === undefined ? '' : String(cfg[k]));
  const sh = sheet_('Config');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const keys = Object.keys(cur);
  if (keys.length) sh.getRange(2, 1, keys.length, 2).setValues(keys.map(k => [k, cur[k]]));
}
function menuOut_(m) { return { id: String(m.id), name: m.name, category: m.category, price: Number(m.price), image: m.image }; }
function orderOut_(o) {
  let items = [];
  try { items = JSON.parse(o.items || '[]'); } catch (e) {}
  return {
    id: String(o.id), timestamp: String(o.timestamp), customer: o.customer, phone: String(o.phone), type: o.type,
    table: o.table, items: items, total: Number(o.total), note: o.note, status: o.status, payment: o.payment,
    rating: o.rating, riderNote: o.riderNote, deliveryPhoto: o.deliveryPhoto, deliveredAt: String(o.deliveredAt || '')
  };
}
function nextOrderId_() {                    // ORD-ปีเดือนวัน-ลำดับของวัน เช่น ORD-261110-007
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(new Date(), TZ, 'yyMMdd');
  const last = (props.getProperty('ORDER_SEQ') || '').split(':');
  const seq = last[0] === day ? Number(last[1]) + 1 : 1;
  props.setProperty('ORDER_SEQ', day + ':' + seq);
  return 'ORD-' + day + '-' + ('00' + seq).slice(-3);
}

/* ------------------------------------------------------ utility / security */

function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function nowIso_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"); }
function num_(v) { const n = Number(v); return isFinite(n) && v !== '' && v !== null ? n : ''; }
function phone_(v) {
  let s = String(v || '').replace(/[\s-]/g, '');
  if (s.indexOf('+66') === 0) s = '0' + s.slice(3);
  if (!/^0\d{8,9}$/.test(s)) throw new Error('กรุณากรอกเบอร์โทรศัพท์ให้ถูกต้อง');
  return s;
}
function hex_(bytes) { return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join(''); }
function hashPw_(pw, salt) {
  let h = String(salt) + String(pw);
  for (let i = 0; i < 500; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  return h;
}
function makeToken_(role, sub, ttlSec) {
  const body = Utilities.base64EncodeWebSafe(JSON.stringify({ role: role, sub: sub, exp: Date.now() + ttlSec * 1000 }));
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(body, prop_('TOKEN_SECRET')));
  return body + '.' + sig;
}
function auth_(token, roles) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2) throw new Error('กรุณาเข้าสู่ระบบ');
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(parts[0], prop_('TOKEN_SECRET')));
  if (sig !== parts[1]) throw new Error('กรุณาเข้าสู่ระบบ');
  let d;
  try { d = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString()); }
  catch (e) { throw new Error('กรุณาเข้าสู่ระบบ'); }
  if (d.exp < Date.now()) throw new Error('หมดอายุการเข้าสู่ระบบ กรุณาเข้าใหม่');
  if (roles.indexOf(d.role) < 0) throw new Error('ไม่มีสิทธิ์ดำเนินการ');
  return d;
}
function session_(role, u) {
  return {
    token: makeToken_(role, u.phone, 30 * 24 * 3600),
    user: { name: u.name, phone: String(u.phone), address: u.address,
            coords: (u.lat !== '' && u.lng !== '') ? { lat: Number(u.lat), lng: Number(u.lng) } : null }
  };
}
// กันเดารหัส: ผิดเกิน 8 ครั้งใน 15 นาที ล็อกชั่วคราว
function throttle_(key) {
  const cache = CacheService.getScriptCache(), k = 'thr:' + key;
  const n = Number(cache.get(k) || 0);
  if (n >= 8) throw new Error('ลองผิดหลายครั้งเกินไป กรุณารอ 15 นาทีแล้วลองใหม่');
  cache.put(k, String(n + 1), 900);
}
function clearThrottle_(key) { CacheService.getScriptCache().remove('thr:' + key); }

// เก็บรูปลง Google Drive (เซลล์ใน Sheet จุข้อความได้ไม่เกิน 50,000 ตัวอักษร จึงเก็บรูป base64 ตรง ๆ ไม่ได้)
function saveImage_(dataUrl, prefix) {
  const m = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!m) throw new Error('รูปแบบรูปภาพไม่ถูกต้อง');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 4 * 1024 * 1024) throw new Error('รูปใหญ่เกิน 4MB');
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg').replace(/\+.*/, '');
  const name = prefix + '-' + hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes)).slice(0, 12) + '.' + ext;
  const folders = DriveApp.getFoldersByName('Krutumeng-Images');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Krutumeng-Images');
  const same = folder.getFilesByName(name);
  const file = same.hasNext() ? same.next() : folder.createFile(Utilities.newBlob(bytes, m[1], name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}
