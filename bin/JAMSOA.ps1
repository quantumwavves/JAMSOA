#JAMSOA v 1.0 by QuantumWavves
$Folder = 'C:\Program Files\Microsoft Office\Office16\'
if (Test-Path -Path $Folder) {
    Set-Location 'C:\Program Files\Microsoft Office\Office16\'
} else {
    Set-Location 'C:\Program Files (x86)\Microsoft Office\Office16'
}

$TMSV = cmd.exe /c "cscript ospp.vbs /dstatus" | Select-String 'License Name' 
$TMSV = $TMSV.ToString() -replace "License Name", ""
$TMSV = $TMSV.ToString() -replace "edition", ""
$TMSV = $TMSV.ToString() -replace " ", ""
$TMSV = $TMSV.ToString() -replace ":", ""
$MSOV = $TMSV.ToString() -replace ",", ""

function AMSOVD {
    Write-Output $MSOV
    if($MSOV -eq "Office21Office21ProPlus2021R_Grace"){MSOPP2021}
    if($MSOV -eq "Office19Office19ProPlus2019R_Grace"){MSOPP2019}
    if($MSOV -eq "Office16Office16ProPlusR_Grace"){MSOPP2016}
}

function MSOPP2021{
    cmd.exe /c "for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL*.xrm-ms') do cscript ospp.vbs //B /inslic:'..\root\Licenses16%x'"
    cmd.exe /c "cscript ospp.vbs /setprt:1688"
    cmd.exe /c "cscript ospp.vbs /unpkey:6F7TH >nul"
    cmd.exe /c "cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
    cmd.exe /c "cscript ospp.vbs /sethst:s8.uk.to"
    cmd.exe /c "cscript ospp.vbs /act"
}

function MSOPP2019{
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/quantumwavves/JAMSOA/master/bin/Office19.cmd" -OutFile "$env:temp\2019.cmd"
    cmd.exe /c "$env:temp\2019.cmd"
    Get-ChildItem -Path "$env:temp" *.* -Recurse | Remove-Item -Force -Recurse
}

function MSOPP2016{
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/quantumwavves/JAMSOA/master/bin/Office16.cmd" -OutFile "$env:temp\2019.cmd"
    cmd.exe /c "$env:temp\2016.cmd"
    Get-ChildItem -Path "$env:temp" *.* -Recurse | Remove-Item -Force -Recurse
}

AMSOVD