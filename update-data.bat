@echo off
chcp 65001 >nul
REM อัปเดตข้อมูล: วางไฟล์ Excel ใหม่ใน data-source\ แล้วดับเบิลคลิกไฟล์นี้
node scripts\encrypt-data.mjs || goto :err
git add public/data/data.enc
git commit -m "update data" || echo (ไม่มีการเปลี่ยนแปลง)
git push || goto :err
echo.
echo สำเร็จ - รอ 1-2 นาทีให้ GitHub Pages อัปเดต
pause & exit /b
:err
echo เกิดข้อผิดพลาด
pause
