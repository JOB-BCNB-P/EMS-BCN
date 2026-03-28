# ระบบบริหารจัดการวิชาการ (Google Sheet Edition)

ระบบบริหารจัดการวิชาการ — ดึงข้อมูลจาก Google Sheet แบบ **อ่านอย่างเดียว**  
Deploy บน GitHub Pages ได้ทันที ไม่ต้องมี Backend

---

## วิธีใช้งาน

### 1. สร้าง Google Sheet

1. ไปที่ [Google Sheets](https://sheets.google.com) → สร้าง Spreadsheet ใหม่
2. สร้าง **Tab (แผ่นงาน)** ตามชื่อต่อไปนี้ (ชื่อต้องตรงทุกตัวอักษร):

| Tab | ข้อมูล |
|-----|--------|
| `student` | ข้อมูลนักศึกษา |
| `teacher` | ข้อมูลอาจารย์ |
| `subject` | รายวิชาที่เปิดสอน |
| `schedule` | ตารางเรียน/ตารางสอบ |
| `grade` | ผลการเรียน |
| `eng_result` | ผลสอบภาษาอังกฤษ |
| `leave` | ข้อมูลการลา |
| `evaluation` | ผลประเมินอาจารย์ |
| `tracking` | ติดตามรายละเอียดรายวิชา |
| `announcement` | ประกาศ/แจ้งเตือน |
| `user` | บัญชีผู้ใช้งาน (login) |

3. **แถวแรก** ของแต่ละ Tab = หัวคอลัมน์ (ดูจาก CSV template)
4. แถวที่ 2 เป็นต้นไป = ข้อมูลจริง

### 2. ตั้ง Share เป็น Public

1. คลิก **Share** (มุมขวาบน)
2. ใน "General access" เปลี่ยนเป็น **"Anyone with the link"**
3. สิทธิ์เป็น **"Viewer"**
4. คลิก **Done**
5. คัดลอก URL

### 3. Deploy บน GitHub Pages

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```

Settings → Pages → Source: `main` → Save

### 4. เปิดเว็บ & วาง URL

1. เปิด `https://<your-username>.github.io/<repo-name>/`
2. วาง URL ของ Google Sheet → คลิก **"เชื่อมต่อ"**
3. Login ด้วยข้อมูลจาก Tab `user`

---

## โครงสร้างไฟล์

```
├── index.html      # หน้าหลัก
├── styles.css      # CSS
├── app.js          # Application logic (read-only)
├── gsheet-db.js    # Google Sheet data layer
└── README.md
```

---

## หัวคอลัมน์ของแต่ละ Tab

### Tab: `user` (บัญชี login)
```
name | role | email | password | national_id | responsible_year
```
- `role`: `admin` / `teacher` / `classTeacher` / `student`
- admin ใช้ `password` (6 หลัก)
- teacher/classTeacher ใช้ `email` + `password`
- student ใช้ `national_id` (13 หลัก) จาก Tab student

### Tab: `student`
```
name | student_id | status | phone | email | parent_name | parent_phone | advisor | year_level | room | national_id
```

### Tab: `teacher`
```
name | position | department | phone | email | responsible_year
```

### Tab: `subject`
```
subject_name | coordinator | year_level | room | credits | semester | academic_year
```

### Tab: `schedule`
```
subject_name | schedule_date | schedule_time | schedule_type | room | year_level
```

### Tab: `grade`
```
name | student_id | subject_name | grade | credits | semester | academic_year
```

### Tab: `eng_result`
```
name | eng_score | eng_type | eng_status
```

### Tab: `leave`
```
name | subject_name | leave_hours | leave_percent | semester | academic_year | leave_date | leave_type
```

### Tab: `announcement`
```
announcement_title | announcement_content | announcement_date | event_type
```

### Tab: `evaluation`
```
subject_name | name | eval_topic | eval_score | semester | academic_year
```

### Tab: `tracking`
```
subject_name | theory_practice | year_level | room | semester | coordinator | class_teacher_check | academic_propose | deputy_sign | approved_date
```

---

## การอัปเดตข้อมูล

- แก้ไขข้อมูลใน Google Sheet โดยตรง
- เว็บจะ auto-refresh ทุก 60 วินาที
- หรือกดปุ่ม 🔄 ที่ header เพื่อรีเฟรชทันที

---

## หมายเหตุ

- ระบบเป็น **read-only** — ไม่สามารถเพิ่ม/แก้ไข/ลบข้อมูลผ่านหน้าเว็บ
- ทุกการแก้ไขทำใน Google Sheet → เว็บดึงข้อมูลมาแสดงอัตโนมัติ
- ไม่ต้องมี API Key, Firebase, หรือ Backend ใดๆ
- รองรับ Google Sheet ที่เป็น Public (Anyone with link → Viewer)
