// ตั้งค่ากลางของทุกหน้า — แก้ URL ที่นี่ที่เดียว
// นำ URL ของ Web App (Apps Script > Deploy > Web app ลงท้ายด้วย /exec) มาวางแทน
const API_URL = "https://script.google.com/macros/s/AKfycbxrJeY6drhz5RVjS9lMhyXQK91enUgEZraEgIQuJkSmv6QsTdZSM-FZHErbAdkJb4Q/exec";

// เรียกคำสั่งไปที่ Google Sheet (ผ่าน Apps Script)
// ส่งเป็น text/plain เพื่อไม่ให้เบราว์เซอร์ติด CORS preflight และอ่านผลลัพธ์กลับมาได้
async function apiCall(action, payload) {
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(Object.assign({ action: action }, payload || {}))
  });
  let j;
  try { j = await res.json(); } catch (e) { throw new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่สำเร็จ'); }
  if (!j.ok) throw new Error(j.error || 'เกิดข้อผิดพลาด');
  return j.data;
}

// โหลดเมนู + ตั้งค่าร้าน (ข้อมูลสาธารณะ)
async function loadPublicData() {
  const res = await fetch(API_URL, { cache: 'no-store' });
  const j = await res.json();
  if (j.ok === false) throw new Error(j.error || 'โหลดข้อมูลไม่สำเร็จ');
  return j;
}
