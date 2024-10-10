# Kích hoạt Office 2024 Project
- ![image](https://github.com/user-attachments/assets/892ab962-1334-4126-9b74-42be48da0f04)
- ![image](https://github.com/BsNgChiThanh/Lich-phong-kham/assets/82578024/d575f08f-29b1-4848-83b0-fb5e88dcb50c)
- Source cài đặt:
- Office 2024 Project Pro [tại đây](https://raw.githubusercontent.com/BsNgChiThanh/Office2024Project/IMP/OfficeProjectPro2024.exe) hoặc [tại đây](https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectPro2024Retail&platform=x64&language=de-de&version=O16GA)
- Office 2024 Project Std [tại đây](https://raw.githubusercontent.com/BsNgChiThanh/Office2024Project/IMP/OfficeProjectStd2024.exe) hoặc [tại đây](https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectStd2024Retail&platform=x64&language=de-de&version=O16GA)

## CÁCH 1:
- Dùng MAS Tools
- https://github.com/BsNgChiThanh/MAS-TOOL

## CÁCH 2

```php
for %a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%a")
if exist "%ProgramFiles% (x86)\Microsoft Office\Office1%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%a"))
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms" >nul
(for /f %%x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\projectprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\projectpro2024vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:D9GTG-NP7DV-T6JP3-B6B62-JB89R
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act
```
đang cập nhật...
