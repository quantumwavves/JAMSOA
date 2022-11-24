@echo off
:ADMIN
openfiles >nul 2>nul ||(
echo CreateObject^("Shell.Application"^).ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
"%temp%\getadmin.vbs" >nul 2>&1
goto:eof
)
del /f /q "%temp%\getadmin.vbs" >nul 2>nul

for /f "tokens=6 delims=[]. " %%G in ('ver') do set win=%%G

setlocal

pushd "%~dp0"
Title Office 2019 Retail to Volume License Converter

rem SET OfficePath=%ProgramFiles%\Microsoft Office
SET OfficePath=%ProgramFiles(x86)%\Microsoft Office
if not exist "%OfficePath%\root\Licenses16" SET OfficePath=%ProgramFiles(x86)%\Microsoft Office
if not exist "%OfficePath%\root\Licenses16" (
    echo Could not find the license files for Office 2019!
    goto :eof
)

echo Press Enter to start VL-Conversion...
echo.
echo.
cd /D "%SystemRoot%\System32"

if %win% GEQ 9200 (
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\ProPlus2019VL_KMS_Client_AE-ppd.xrm-ms"
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\ProPlus2019VL_KMS_Client_AE-ul.xrm-ms"
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\ProPlus2019VL_KMS_Client_AE-ul-oob.xrm-ms"


    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-bridge-office.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-root.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-root-bridge-test.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-stil.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-ul.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-ul-oob.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\pkeyconfig-office.xrm-ms
)
 if %win% LSS 9200 (
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\ProPlus2019VL_KMS_Client_AE-ppd.xrm-ms"
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\ProPlus2019VL_KMS_Client_AE-ul.xrm-ms"
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\ProPlus2019VL_KMS_Client_AE-ul-oob.xrm-ms"


    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-bridge-office.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-root.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-root-bridge-test.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-stil.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-ul.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-ul-oob.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\pkeyconfig-office.xrm-ms
)

cscript "%OfficePath%\Office16\ospp.vbs" /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript "%OfficePath%\Office16\ospp.vbs" /act

echo.
echo Retail to Volume License conversion finished.
echo.
