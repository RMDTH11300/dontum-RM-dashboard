# DRMS 1.0 — โรงพยาบาลดอนตูม

ระบบใหม่ถูกแยกไว้ในโฟลเดอร์ `drms-v6` และไม่แก้ Dashboard เดิม

## ลิงก์หลังอัปโหลด
- ระบบใหม่: `/dontum-RM-dashboard/drms-v6/`
- Admin: `/dontum-RM-dashboard/drms-v6/admin.html`

## การทำงาน
ระบบอ่านข้อมูลจากโฟลเดอร์ `../data/` ของ Dashboard เดิม จึงไม่ต้องคัดลอกข้อมูลซ้ำ

## วิธีอัปเดตข้อมูลจาก Admin
1. เปิด `admin.html`
2. เลือกปีและกดโหลดข้อมูล
3. เพิ่ม/แก้ไข/ลบ
4. กดส่งออก JSON
5. อัปโหลดไฟล์ที่ได้ไปทับในโฟลเดอร์ `data` ของ Repository

> GitHub Pages เป็น Static Site หน้า Admin จึงไม่สามารถบันทึกกลับ Repository โดยตรง
