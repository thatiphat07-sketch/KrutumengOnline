# ร้านกรูตำเอง — ระบบสั่งอาหาร + POS (ฐานข้อมูล Google Sheet)

ข้อมูลทั้งหมด (เมนู, ตั้งค่าร้าน, ลูกค้า, ออเดอร์) เก็บใน Google Sheet ผ่าน Google Apps Script

## ไฟล์
- `Code.gs`      ฝั่งหลังบ้าน วางใน Apps Script ของ Google Sheet
- `config.js`    ที่เดียวที่ต้องใส่ API_URL (ใช้ร่วมกันทุกหน้า)
- `index.html` `customer.html` `admin.html` `rider.html`  หน้าเว็บ

## ติดตั้ง
1. สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script > วาง `Code.gs` ทั้งไฟล์
2. เลือกฟังก์ชัน `setup` > Run (กดอนุญาตสิทธิ์) — จะสร้างแท็บ Orders / Menu / Config / Users
3. Project Settings > Script Properties > แก้ `ADMIN_PIN` และ `RIDER_PIN` (ค่าเริ่มต้น 1234 / 5678 — ต้องเปลี่ยน)
4. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone
5. คัดลอก URL ที่ลงท้าย `/exec` ไปใส่ `API_URL` ใน `config.js`
6. อัปโหลดไฟล์เว็บขึ้นโฮสต์ (GitHub Pages ฯลฯ)

แก้ Code.gs ภายหลัง: Deploy > Manage deployments > Edit > New version (URL เดิมจะยังใช้ได้)
