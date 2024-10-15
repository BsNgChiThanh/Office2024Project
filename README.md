# Kích hoạt Office 2024 Project Pro
- ![image](https://github.com/user-attachments/assets/892ab962-1334-4126-9b74-42be48da0f04)
- ![image](https://github.com/BsNgChiThanh/Lich-phong-kham/assets/82578024/d575f08f-29b1-4848-83b0-fb5e88dcb50c)
- Source cài đặt:
  - Office 2024 Project Pro [tại đây](https://raw.githubusercontent.com/BsNgChiThanh/Office2024Project/IMP/OfficeProjectPro2024.exe) hoặc [tại đây](https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectPro2024Retail&platform=x64&language=de-de&version=O16GA)
  - Office 2024 Project Std [tại đây](https://raw.githubusercontent.com/BsNgChiThanh/Office2024Project/IMP/OfficeProjectStd2024.exe) hoặc [tại đây](https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectStd2024Retail&platform=x64&language=de-de&version=O16GA)

Có nhiều cách kích hoạt, xong tôi chỉ ra 2 cách kích hoạt điển hình nhất:

# CÁCH 1:
- [Dùng MAS Tools](https://github.com/BsNgChiThanh/MAS-TOOL)
- Các bạn chạy **PowerShell** bằng quyền **Run as Administrator**, rồi dán câu lệnh sau đâu vào:
  ```php
  irm https://raw.githubusercontent.com/BsNgChiThanh/MAS-TOOL/IMP/MAS.ps1 | iex
  ```
- Làm theo chỉ dẫn của cửa sổ hiện lên.
- Done!

## CÁCH 2
- ĐẦU TIÊN KÍCH HOẠT BẰNG KEY KMS 180 NGÀY:
  - Chạy **cmd** bằng quyền **Run as Administrator** rồi dán đoạn lệnh này vào, nhấn enter:
  ```php
  set v=16
  if exist "%ProgramFiles%\Microsoft Office\Office%v%\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office%v%"
  if exist "%ProgramFiles(x86)%\Microsoft Office\Office%v%\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office%v%"
  cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms" >nul
  (for /f %%x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
  (for /f %%x in ('dir /b ..\root\Licenses16\projectprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
  (for /f %%x in ('dir /b ..\root\Licenses16\projectpro2024vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
  cscript ospp.vbs /setprt:1688
  cscript ospp.vbs /inpkey:D9GTG-NP7DV-T6JP3-B6B62-JB89R
  cscript ospp.vbs /inpkey:GNJ6P-Y4RBM-C32WW-2VJKJ-MTHKK
  cscript ospp.vbs /sethst:107.175.77.7
  cscript ospp.vbs /sethst:172.16.0.2
  cscript ospp.vbs /act
  ```
  - ![image](https://github.com/user-attachments/assets/0d2a21a0-38da-44e6-b053-637254f5d70f)
- SAU ĐÓ TẠO TÁC VỤ GIA HẠN OFFICE HÀNG TUẦN:
  - Chạy Windows PowerShell bằng quyền Run as Administrator rồi dán đoạn lệnh này vào, nhấn enter:
    
  ```PHP
  irm https://raw.githubusercontent.com/BsNgChiThanh/Kich-hoat-Office/KichHoatOffice/GiaHanKichHoat.ps1 | iex
  ```
  - ![image](https://github.com/user-attachments/assets/c61d847b-f874-4549-92af-f49985044f7e)
- DONE!
