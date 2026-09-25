# check-edi-AGN

เครื่องมือตรวจสอบเอกสารขาเข้าเรือก่อนจัดทำใบขนสินค้าขาเข้า (Thailand ocean import) —
เทียบ **MANIFEST** (รายงานจากสายเรือ, .xls/.xlsx) กับ **ENTER** (เอกสารยืนยัน/แก้ไขจาก
ลูกค้า, .pdf) ทีละ B/L ย่อย แล้วสร้างรายงาน `EDI.xlsx` แบบไฮไลต์สีบอกจุดที่ต้องแก้ก่อน
ยื่นใบขน

## ตรวจอะไรบ้าง

เทียบระหว่าง ENTER กับ MANIFEST:

- CONSIGNEE, CONTAINER NO., STATUS (CY/LCL/LCL-CFS)
- TOTAL PACKAGE, ชนิดบรรจุภัณฑ์ (PACKAGING)
- GROSS WEIGHT, MEASUREMENT
- MARKS, DESCRIPTION
- REEFER TEMP, DG (CLASS/UN), TRANSIT/TRANSHIPMENT
- ยอดรวมทั้งฉบับ เทียบกับยอดที่แต่ละเอกสารระบุเอง (GRAND TOTAL/PACKAGES-KGS-CBM)
- B/L ที่มีใน ENTER แต่ตกหล่นจาก MANIFEST (เสี่ยงหลุดจากใบขน)

ตรวจภายใน MANIFEST เอง (ไม่ต้องพึ่ง ENTER ยกเว้นที่ระบุ) ตามกฎที่กำหนดไว้ตายตัว:

1. **SHED NO.** — USED ENGINE/LAOS ต้องเป็น 0126/0124 (เฉพาะ DISCHARGE BANGKOK),
   DG ต้องเป็น 2826 (เฉพาะ LAEM CHABANG), THLKR/BMT/SCT/UNITHAI ต้องเป็น
   0332/0110/0302/0113
2. **ประเทศปลายทาง** กรณี TRANSIT/TRANSSHIPMENT ต้องตรงตารางรหัสประเทศที่กำหนด (เช่น
   LAOS→LL, MYANMAR→MM, CHINA→CN, …)
3. **CARGO MOVEMENT** — TRANSIT ไปที่อื่นต้องเป็น `(7-TRANSIT)` เสมอ ยกเว้นไปลาวที่ต้อง
   ตรงกับค่าที่ ENTER ระบุเท่านั้น
4. **TAX ID vs NOTIFY** — ชื่อใน TAX ID ต้องตรงกับ N:/NOTIFY PARTY ของ MANIFEST
   ยกเว้น B/L ที่มี TRANSIT/TRANSHIPMENT จะยอมรับชื่อที่ ENTER ระบุแยกไว้เองด้วย

## ติดตั้ง

ต้องมี Python 3.9+

```bash
pip install -r requirements.txt
```

## วิธีใช้

วางไฟล์ MANIFEST (.xls/.xlsx) และ ENTER (.pdf) ไว้ในโฟลเดอร์เดียวกัน (หรือใน
โฟลเดอร์ย่อยชื่อ `input/`) แล้วรัน:

```bash
python scripts/build_edi.py
```

สคริปต์จะหาไฟล์ `.xls`/`.xlsx` และ `.pdf` ให้อัตโนมัติ (เลือกไฟล์ที่ชื่อขึ้นต้นด้วย
`MANIFEST`/`ENTER` ก่อน ถ้าไม่เจอจะใช้ไฟล์ประเภทนั้นไฟล์เดียวที่มีในโฟลเดอร์) แล้วสร้าง
`EDI.xlsx` ไว้ที่เดียวกับไฟล์ MANIFEST

หรือระบุพาธเองและเลือกโฟลเดอร์ผลลัพธ์:

```bash
python scripts/build_edi.py --manifest input/MANIFEST.xls --enter input/ENTER.pdf --outdir input
```

สคริปต์จะพิมพ์สรุปจำนวน B/L, จุดที่ต้องตรวจ, และรายการปัญหาแต่ละ B/L ออกทาง terminal
ด้วย ก่อนเปิดไฟล์ Excel ดูรายละเอียด

## อ่านผลลัพธ์ EDI.xlsx

- **สีแดง** ⚠ = ไม่ตรงกันแน่นอน ต้องแก้ก่อนยื่นใบขน
- **สีส้ม** = ต้องให้คนตรวจสอบยืนยัน (เช่น ดึงข้อมูลด้วยความมั่นใจต่ำ, หรือ CARGO
  MOVEMENT ไปลาวที่หา ENTER มาเทียบไม่เจอ)
- **สีเหลือง** = B/L นี้ไม่มีเอกสาร ENTER ให้เทียบ
- **สีเขียว** = ผ่าน (เฉพาะคอลัมน์สรุปผลเร่งด่วน คอลัมน์ A)
- คอลัมน์ A (สรุปผลเร่งด่วน) และคอลัมน์สุดท้าย (หมายเหตุ) สรุปทุกจุดที่ผิดของ B/L นั้น
  ไว้ให้อ่านรวดเดียว ไม่ต้องไล่ดูทีละคอลัมน์

## รองรับเลย์เอาต์ไหนบ้าง

- **MANIFEST.xls**: รายงานอิสระ (free-form) 1 บล็อกต่อ 1 B/L ย่อย — คอลัมน์ marks/
  container ถูกเดาตำแหน่งอัตโนมัติจากเนื้อหาจริง (ไม่ต้องแก้โค้ดเองถ้าไฟล์สายเรืออื่น
  จัดคอลัมน์เยื้องไปนิดหน่อย)
- **ENTER.pdf**: รองรับ 4 เลย์เอาต์ที่รู้จัก (auto-detect) ครอบคลุมฟอร์มจากหลาย
  forwarder/สายเรือ รวมถึงตารางแบบมีเส้นคั่นจริง (x0-based column) และแบบ B/L NO. เป็น
  anchor ข้อความ — ถ้าเจอ ENTER รูปแบบใหม่ที่ยังไม่ตรง ให้ดูคอมเมนต์ใน
  `scripts/build_edi.py` (แต่ละ `parse_enter_layout_*()`) ประกอบ

## หมายเหตุ

รูปแบบรายงานนี้ปรับจากการใช้งานจริงหลายครั้งกับหลายสายเรือ/ลูกค้า — สี, ลำดับคอลัมน์,
และกฎการตัดสินแต่ละจุดเป็นข้อตกลงที่ยืนยันแล้ว ไม่ใช่ค่าที่เดาเอง

## ใช้งานผ่านเว็บ (ในเครื่อง/วงแลน)

```bash
pip install -r requirements.txt
python webapp/app.py
```

เปิด http://localhost:5000 (เครื่องอื่นในวงแลนเปิด `http://<IP เครื่องนี้>:5000`) → เลือกไฟล์ MANIFEST + ENTER
ใส่ SHED NO. แล้วกด **ตรวจสอบ** จะเห็นสรุปจุดผิด + ตารางสีเหมือน EDI.xlsx และปุ่มดาวน์โหลด EDI.xlsx

SHED NO.: ทุก B/L ต้องตรงกับเลขที่ใส่ (ตรง = `-`) ยกเว้น LAOS ใช้กฎเดิม 0124
(บรรทัดคำสั่ง: `python scripts/build_edi.py --shed 0141`)
