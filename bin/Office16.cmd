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

set OPPKEY=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
set PPKEY=YG9NW-3K39V-2T3HJ-93F3Q-G83KT
set VPKEY=PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
set S4BKEY=869NQ-FJ69K-466HW-QYCP2-DDBV6

pushd "%~dp0"
Title Office 2016 Retail to Volume License Converter

SET OfficePath=%ProgramFiles%\Microsoft Office
if not exist "%OfficePath%\root\Licenses16" SET OfficePath=%ProgramFiles(x86)%\Microsoft Office
if not exist "%OfficePath%\root\Licenses16" (
    echo Could not find the license files for Office 2016!
    goto :eof
)

echo Press Enter to start VL-Conversion...
echo.
echo.
cd /D "%SystemRoot%\System32"

if %win% GEQ 9200 (
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms"
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms"
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms"

    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms"
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms"
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms"

    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\VisioProVL_KMS_Client-ppd.xrm-ms"
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\VisioProVL_KMS_Client-ul-oob.xrm-ms"
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\VisioProVL_KMS_Client-ul.xrm-ms"

    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\SkypeforBusinessVL_KMS_Client-ppd.xrm-ms"
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\SkypeforBusinessVL_KMS_Client-ul.xrm-ms"
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\SkypeforBusinessVL_KMS_Client-ul-oob.xrm-ms"

    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-bridge-office.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-root.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-root-bridge-test.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-stil.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-ul.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\client-issuance-ul-oob.xrm-ms
    cscript slmgr.vbs /ilc "%OfficePath%\root\Licenses16\pkeyconfig-office.xrm-ms
)
 if %win% LSS 9200 (
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms"
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms"
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms"

    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms"
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms"
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms"

    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\VisioProVL_KMS_Client-ppd.xrm-ms"
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\VisioProVL_KMS_Client-ul-oob.xrm-ms"
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\VisioProVL_KMS_Client-ul.xrm-ms"0

    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\SkypeforBusinessVL_KMS_Client-ppd.xrm-ms"
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\SkypeforBusinessVL_KMS_Client-ul.xrm-ms"
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\SkypeforBusinessVL_KMS_Client-ul-oob.xrm-ms"

    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-bridge-office.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-root.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-root-bridge-test.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-stil.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-ul.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\client-issuance-ul-oob.xrm-ms
    cscript "%OfficePath%\Office16\ospp.vbs" /inslic:"%OfficePath%\root\Licenses16\pkeyconfig-office.xrm-ms
)

for %%a in (%OPPKEY% %PPKEY% %VPKEY% %S4BKEY%) do cscript "%OfficePath%\Office16\ospp.vbs" /inpkey:%%a
cscript "%OfficePath%\Office16\ospp.vbs" /act

echo.
echo Retail to Volume License conversion finished.
echo.

