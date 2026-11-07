# ผู้ช่วยเขียนงานวิจัยทางการพยาบาล
**Nursing Research Writer** · พัฒนาโดย Pannathat Chanpanit, Sisaket

เว็บแอปช่วยวางโครง เขียน ตรวจ และส่งออกงานวิจัยทางการพยาบาล 5 บท เป็นไฟล์ Word ภาษาไทย

🔗 **ใช้งาน:** https://pannathatc-cell.github.io/nursing-research-writer/

## ความสามารถหลัก
- เขียนรายบทและรายหัวข้อ พร้อมผู้ช่วย AI (Claude, OpenAI, Gemini หรือผู้ให้บริการที่เข้ากันได้กับ Claude/OpenAI)
- ตรวจสอบเอกสารอ้างอิง APA 7th edition: รายการท้ายงานเขียน สถานะยืนยันแล้ว/ยังไม่ยืนยัน/ไม่พบแหล่ง และตรวจกับ Crossref, OpenAlex, ThaiJO
- แชตบอต "น้องค้นคว้า" ช่วยสืบค้นและตอบคำถาม สั่งด้วยเสียงได้
- เรียบเรียงและพาราเฟรส: เรียบเรียงใหม่ 8 รูปแบบ ตรวจไวยากรณ์ แปลไทย-อังกฤษ สรุปความ และตรวจความซ้ำกับต้นฉบับ
- แหล่งสืบค้น: CINAHL และ EBSCO (มหาวิทยาลัย), ProQuest, ThaiLIS (TDC), PubMed, ScienceDirect, Scopus, ThaiJO, Google Scholar
- สถิติการวิจัย คำนวณขนาดตัวอย่าง จริยธรรมการวิจัย โหมดผู้ทรงคุณวุฒิ
- ส่งออก Word (TH Sarabun) และบันทึกลง Google Drive / Google Sheets

## โครงสร้าง
| ไฟล์ | หน้าที่ |
|---|---|
| `index.html` | ตัวโปรแกรมทั้งหมด (ไฟล์เดียว) |
| `apps-script/Code.gs` | โค้ด Google Apps Script สำหรับบันทึกโครงการลง Google Drive และ Google Sheets |

## ความเป็นส่วนตัว
- ข้อมูลโครงการเก็บในเบราว์เซอร์ของผู้ใช้แต่ละเครื่อง ไม่ได้ส่งขึ้น GitHub
- มี API key ทดลองใช้ (อายุ 30 วัน) ตั้งไว้ในโปรแกรม ผู้ใช้เปลี่ยนเป็น key ของตนเองได้ที่เมนูตั้งค่าผู้ช่วย AI
- โปรแกรมเชื่อมกับ Google Drive ของผู้พัฒนาไว้แล้ว ผู้ใช้กดบันทึกโครงการและไฟล์ Word ลง Drive ได้ที่เมนู "บันทึกลง Google Drive" (ผู้ใช้ทุกคนเห็นรายชื่อโครงการที่สำรองไว้ใน Drive นี้)
- ก่อนใช้ Apps Script ให้ตั้งค่า `TOKEN` ใน `apps-script/Code.gs` เป็นรหัสลับของตัวเอง

## การบันทึกลง Google Drive
1. เปิด Google Sheets ปลายทาง > ส่วนขยาย > Apps Script แล้ววางโค้ดจาก `apps-script/Code.gs` และตั้ง `TOKEN`
2. เรียกใช้ฟังก์ชัน `setup` หนึ่งครั้ง แล้วอนุญาตสิทธิ์
3. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone)
4. ในโปรแกรม เมนู "บันทึกลง Google Drive" ใส่ Web app URL และ Token แล้วกดทดสอบ

---
© 2026 Pannathat Chanpanit, Faculty of Nursing, Sisaket Rajabhat University
