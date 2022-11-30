#JAMSOA v 1.0 by QuantumWavves
$Folder = 'C:\Program Files\Microsoft Office\Office16\'
if (Test-Path -Path $Folder) {
    Set-Location 'C:\Program Files\Microsoft Office\Office16\'
} else {
    Set-Location 'C:\Program Files (x86)\Microsoft Office\Office16'
}
$WINVERSION=(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
$TMSV = cmd.exe /c "cscript ospp.vbs /dstatus" | Select-String 'License Name' 
$TMSV = $TMSV.ToString() -replace "License Name", ""
$TMSV = $TMSV.ToString() -replace "edition", ""
$TMSV = $TMSV.ToString() -replace " ", ""
$TMSV = $TMSV.ToString() -replace ":", ""
$MSOV = $TMSV.ToString() -replace ",", ""

function AMSOVD {
    Write-Output $MSOV
    if($MSOV -eq "Office21Office21ProPlus2021R_Grace"){MSOPP2021}
    if($MSOV -eq "Office19Office19ProPlus2019R_Grace"){
        if ($WINVERSION -eq "Windows 10 Enterprise LTSC 2021") {
            RETTOVLKMS2019LTSC
            MSOPP2019
        }
        else {
            MSOPP2019
        }
    }
    if($MSOV -eq "Office16Office16ProPlusR_Grace"){
        if ($WINVERSION -eq "Windows 10 Enterprise LTSC 2021") {
            RETTOVLKMS2016LTSC
            MSOPP2016
        }
        else {
            MSOPP2016
        }
    }
}
function MSOPP2021{
    cmd.exe /c "for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL*.xrm-ms') do cscript ospp.vbs //B /inslic:'..\root\Licenses16%x'"
    cmd.exe /c "cscript ospp.vbs /setprt:1688"
    cmd.exe /c "cscript ospp.vbs /unpkey:6F7TH >nul"
    cmd.exe /c "cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
    cmd.exe /c "cscript ospp.vbs /sethst:s8.uk.to"
    cmd.exe /c "cscript ospp.vbs /act"
}
function MSOPP2019 {
    cmd.exe /c "for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:'..\root\Licenses16\%x'"
    cmd.exe /c "cscript ospp.vbs /setprt:1688"
    cmd.exe /c "cscript ospp.vbs /unpkey:6MWKP >nul"
    cmd.exe /c "cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
    cmd.exe /c "cscript ospp.vbs /sethst:s8.uk.to"
    cmd.exe /c "cscript ospp.vbs /act"
}
function MSOPP2016 {
    cmd.exe /c "for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:'..\root\Licenses16%x'"
    cmd.exe /c "cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
    cmd.exe /c "cscript ospp.vbs /unpkey:BTDRB >nul"
    cmd.exe /c "cscript ospp.vbs /unpkey:KHGM9 >nul"
    cmd.exe /c "cscript ospp.vbs /unpkey:CPQVG >nul"
    cmd.exe /c "cscript ospp.vbs /sethst:kms8.msguides.com"
    cmd.exe /c "cscript ospp.vbs /setprt:1688"
    cmd.exe /c "cscript ospp.vbs /act"
}
function RETTOVLKMS2019LTSC{
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/quantumwavves/JAMSOA/master/bin/Office19.cmd" -OutFile "$env:temp\2019.cmd"
    cmd.exe /c "$env:temp\2019.cmd"
    Get-ChildItem -Path "$env:temp" *.* -Recurse | Remove-Item -Force -Recurse
}
function RETTOVLKMS2016LTSC{
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/quantumwavves/JAMSOA/master/bin/Office16.cmd" -OutFile "$env:temp\2019.cmd"
    cmd.exe /c "$env:temp\2016.cmd"
    Get-ChildItem -Path "$env:temp" *.* -Recurse | Remove-Item -Force -Recurse
}
AMSOVD