**Project Overview: Daily Foreign Exchange Rates via API**
กระบวนการนี้ทำหน้าที่ดึงข้อมูลอัตราแลกเปลี่ยนเช่นเดิม แต่เปลี่ยนมาเชื่อมต่อกับ API Endpoint โดยตรง (ผ่าน Cloudflare Worker) ซึ่งช่วยให้บอททำงานได้เร็วขึ้น ลดความเสี่ยงจาก UI เว็บไซต์เปลี่ยนแปลง และเพิ่มความเสถียรในการทำงาน

1. Workflow Structure (Flowchart)
โครงสร้างหลักใน Main.xaml ยังคงใช้ Flowchart ในการควบคุมทิศทาง แต่เปลี่ยน Logic ภายในส่วนการดึงข้อมูล:

Initial Configuration: อ่านไฟล์ Process_Config.json เพื่อตั้งค่าตัวแปร เช่น ช่วงวันที่, สกุลเงินที่ต้องการ และความเร็วในการหน่วงเวลา (Delay)

Environment Preparation: สร้าง Folder และเรียก ClearFiles.xaml เพื่อล้างข้อมูลเก่า

BOT API Process: ขั้นตอนหลักในการดึงข้อมูลผ่านระบบ Network:

Date Parsing: แปลงรูปแบบวันที่จาก Config ให้เป็นรูปแบบ yyyy-MM-dd ตามข้อกำหนดของ API

HTTP Request: ใช้ ui:HttpClient ส่ง GET Request ไปยัง API พร้อมแนบ Parameter (Start Period, End Period, Currency)

JSON Extraction: รับผลลัพธ์ในรูปแบบ JSON และใช้ Invoke Code (Newtonsoft.Json) เพื่อแปลงข้อมูลดิบให้กลายเป็น DataTable

Excel Reporting: นำข้อมูลจาก API มาทำการ Merging กับรายการวันที่ทั้งหมดในเดือนนั้นๆ เพื่อสร้างรายงานที่สมบูรณ์

Notification & Cleanup: ส่งผลลัพธ์ผ่าน SendMail.xaml และลบไฟล์ CSV/XLSX ชั่วคราวออก

2. Key Component Differences & Upgrades
A. API Integration (HTTP Client)
แทนที่จะต้องเลื่อนหาปุ่มหรือเลือกวันที่ในหน้าเว็บ บอทตัวนี้ใช้วิธีเรียกใช้งาน API:

Endpoint: https://uipath-cloudflare-exchangerate...

Efficiency: ลดการใช้ทรัพยากรเครื่อง (CPU/RAM) เนื่องจากไม่ต้องเปิด Browser

Reliability: ไม่ได้รับผลกระทบจากโฆษณา, Pop-up หรือการเปลี่ยนแปลงของหน้าจอเว็บไซต์

B. JSON Data Processing
มีการใช้ Library Newtonsoft.Json ภายใน Invoke Code เพื่อเจาะลึกเข้าไปใน Object JSON:

ดึงข้อมูลจาก result -> data -> data_detail

Mapping ค่า buying_sight, buying_transfer และ selling เข้าสู่คอลัมน์ของ DataTable โดยตรง

C. Data Merging with LINQ
ในส่วนการเขียน Excel บอทมีการใช้ LINQ (Group Join) เพื่อรวมข้อมูล:

สร้างตารางที่มี "วันที่ครบทุกวัน" (รวมวันเสาร์-อาทิตย์)

นำข้อมูลจาก API มา Matching ตามวันที่

หากวันไหนไม่มีข้อมูล (วันหยุดราชการ/เสาร์-อาทิตย์) บอทจะเติม Remark อัตโนมัติว่า "is not working day"

3. Error Handling & Resilience
Screenshot on Fail: หาก API ล่มหรือการประมวลผลผิดพลาด ระบบจะ Take Screenshot หน้าจอขณะนั้นทันทีเพื่อการตรวจสอบ

SMTP Fallback: ใน SendMail.xaml ยังคงรักษาความสามารถในการส่งอีเมลผ่าน SMTP ในกรณีที่เครื่องปลายทางไม่มี Microsoft Outlook
