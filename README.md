# ClickUp Public-Share → Excel Exporter

ดึงข้อมูล **ทุก Task** จากลิงค์ ClickUp Public Share (`sharing.clickup.com`)
ออกมาเป็นไฟล์ Excel — ครบทุก Field และ Comment ไม่ใช่แค่คอลัมน์ที่เห็นในตาราง
**ไม่ต้องใช้ API token หรือ login**

## วิธีทำงาน

ระบบเรียก ClickUp public share API ตรงๆ (ตัวเดียวกับที่หน้าเว็บแชร์ใช้):

1. `id.app.clickup.com/shard/v1/handshake/{workspace}` → หา API server (frontdoor) ของ workspace
2. `{frontdoor}/view/v1/{ws}/public/view/{viewId}?token=…` → รายการ task id ทั้งหมดในทุก group
3. `{frontdoor}/view/v1/{ws}/public/view/{viewId}/task/{id}?token=…` → รายละเอียดเต็มของแต่ละ task
   (description, comments, subtasks, checklists, attachments, tags, assignees, custom fields …)

เรียกแบบขนาน (หลาย thread) เพื่อความเร็ว แล้วแปลงเป็น Excel หลาย sheet

## ติดตั้ง

```powershell
cd clickup-scraper
py -3.13 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## ใช้งาน (Web UI)

```powershell
.venv\Scripts\python.exe app.py
```

เปิดเบราว์เซอร์ไปที่ <http://127.0.0.1:5000> → วางลิงค์ → กด **Export เป็น Excel**
(ปุ่ม **ดูตัวอย่าง** จะบอกจำนวน task/comment ที่จะได้ก่อนดาวน์โหลด)

## สร้างไฟล์ .exe (รันบนเครื่องอื่นที่ไม่มี Python)

ดับเบิลคลิก `build_exe.bat` หรือสั่ง:

```powershell
.venv\Scripts\python.exe -m PyInstaller --noconfirm --onefile `
  --name clickup-scraper --add-data "templates;templates" `
  --collect-submodules openpyxl --hidden-import openpyxl.cell._writer app.py
```

จะได้ไฟล์เดียว **`dist\clickup-scraper.exe`** (~13 MB) — ก๊อปไฟล์นี้ไปวางบน
Windows เครื่องไหนก็ได้แล้วดับเบิลคลิก เปิดเบราว์เซอร์ให้อัตโนมัติที่
<http://127.0.0.1:5000> (ถ้า port 5000 ไม่ว่างจะเลือก port อื่นให้เอง)
ปิดโปรแกรมโดยปิดหน้าต่าง console สีดำ

> ไม่ต้องติดตั้ง Python หรืออะไรเพิ่มบนเครื่องปลายทาง — ทุกอย่างรวมอยู่ในไฟล์ .exe แล้ว

## ใช้งานแบบ command line

```powershell
# พิมพ์สรุปออกจอ
.venv\Scripts\python.exe scraper.py "<public-link>"

# สร้างไฟล์ Excel โดยตรง
.venv\Scripts\python.exe -c "import scraper,exporter; r=scraper.scrape('<public-link>'); open('out.xlsx','wb').write(exporter.build_workbook(r))"
```

## ไฟล์ Excel ที่ได้

| Sheet | เนื้อหา |
|-------|---------|
| **Tasks** | 1 แถวต่อ 1 task — ID, ชื่อ, Status, Priority, Assignees, Tags, วันที่ทั้งหมด, ผู้สร้าง, Time estimate, จำนวน comment/attachment/subtask, **Description เต็ม**, Parent, List/Project, URL และ **Custom Fields** (ถ้ามี) |
| **Comments** | 1 แถวต่อ 1 comment — task, ผู้เขียน, วันที่, ข้อความเต็ม, จำนวน reply, reactions |
| **Subtasks** | subtask ทั้งหมด พร้อม status / assignee / due date |
| **Checklists** | รายการ checklist item พร้อมสถานะติ๊ก |
| **Attachments** | ไฟล์แนบทั้งหมด พร้อมชนิด / ขนาด / ลิงค์ดาวน์โหลด |

วันเวลาทั้งหมดแปลงเป็นเขตเวลาไทย (UTC+7) แล้ว

## โครงสร้างโค้ด

```
clickup-scraper/
├── scraper.py        # parse ลิงค์ + เรียก API + ดึงทุก task (core)
├── exporter.py       # แปลงผลลัพธ์เป็น Excel หลาย sheet
├── app.py            # Flask web UI (/, /preview, /export)
├── templates/
│   └── index.html    # หน้าเว็บใส่ลิงค์ + ปุ่ม Export
├── requirements.txt
└── README.md
```

## หมายเหตุ / ข้อจำกัด

- ลิงค์ต้องเป็นแบบ **Public Share ของ View** (มีหลาย task) เช่น `.../<ws>/gr/h/<viewId>/<token>`
- ถ้าเจ้าของปิดการแชร์หรือลิงค์หมดอายุ ระบบจะแจ้ง error
- Field ที่จะดึงได้ขึ้นกับสิ่งที่เจ้าของเปิดเผย (`public_fields`) — โดยปกติเปิด
  content, comments, customFields, subtasks, tags, checklists, attachments
- เป็นการอ่านข้อมูลสาธารณะอย่างเดียว (read-only) ไม่มีการแก้ไขข้อมูลใน ClickUp
```
