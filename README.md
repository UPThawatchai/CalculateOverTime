# CalculateOverTime
เมื่อผมขี้เกียจคำนวนเวลาที่ผม ทำงานล่วงเวลา เลยเขียนโปรแกรม Java เพื่อช่วยคำนวน การทำงานของโปรแกรมคร่าวๆ อ่านไฟล์ xlsx เลือกเอาแต่ column ที่ทำงานเกิน 18:00 มาคำนวนโดยใช้ Lib Apacha POI
- ต้องใช้ File xlsx 
- ต้องมีข้อมูลที่ Column C, Row 9 เท่านั้น เพราะมันเป็น Template ของ Timesheet ของบริษัทผมแต่สามารถประยุกต์ได้
- ข้อมูลที่ใส่ต้องเป็น Format Datetime ex. 18/03/2020 18:09:00
- jar ที่ใช้หาโหลดได้ใน Internet
 - commons-collections4-4.1 .jar
 - commons-io-2.5.jar
 - core-3.2.1.jar
 - poi-3.17.jar
 - poi-excelant-3.17.jar
 - poi-ooxml-3.17.jar
 - poi-ooxml-schemas-3.17.jar
 - xmlbeans-2.6.0.jar
