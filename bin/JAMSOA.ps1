#JAMSOA v 1.0 by QuantumWavves
Set-Location 'C:\Program Files (x86)\Microsoft Office\Office16'

$TMSV = cmd.exe /c "cscript ospp.vbs /dstatus" | Select-String 'License Name' 
$TMSV = $TMSV.ToString() -replace "License Name", ""
$TMSV = $TMSV.ToString() -replace "edition", ""
$TMSV = $TMSV.ToString() -replace " ", ""
$TMSV = $TMSV.ToString() -replace ":", ""
$MSOV = $TMSV.ToString() -replace ",", ""

function AMSOVD {
    echo $MSOV
    if($MSOV -eq "Office21Office21ProPlus2021R_Grace"){MSOPP2021}
    if($MSOV -eq "Office19Office19ProPlus2019R_Grace"){MSOPP2019}
    if($MSOV -eq "Office16Office16ProPlusR_Grace"){MSOPP2016}
}

function MSOPP2021{
    Set-Location 'C:\Program Files\Microsoft Office\Office16\'
    cmd.exe /c "for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL*.xrm-ms') do cscript ospp.vbs //B /inslic:'..\root\Licenses16%x'"
    cmd.exe /c "cscript ospp.vbs /setprt:1688"
    cmd.exe /c "cscript ospp.vbs /unpkey:6F7TH >nul"
    cmd.exe /c "cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
    cmd.exe /c "cscript ospp.vbs /sethst:s8.uk.to"
    cmd.exe /c "cscript ospp.vbs /act"
}

function MSOPP2019{
    cmd.exe /c "for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs //B /inslic:'..\root\Licenses16\%x'"
    cmd.exe /c "cscript ospp.vbs /setprt:1688"
    cmd.exe /c "cscript ospp.vbs /unpkey:6MWKP >nul"
    cmd.exe /c "cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP "
    cmd.exe /c "cscript ospp.vbs /sethst:kms8.msguides.com"
    cmd.exe /c "cscript ospp.vbs /act"
}

function MSOPP2016{
    cmd.exe /c "for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs //B /inslic:'..\root\Licenses16%x'"
    cmd.exe /c "cscript ospp.vbs /setprt:1688"
    cmd.exe /c "cscript ospp.vbs /unpkey:WFG99 >nul"
    cmd.exe /c "cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
    cmd.exe /c "cscript ospp.vbs /sethst:kms8.msguides.com"
    cmd.exe /c "cscript ospp.vbs /act"
}

AMSOVD