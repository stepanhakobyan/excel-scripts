if (Test-Path .\src1) {
    Remove-Item .\src1
}
if (Test-Path .\custom-ribbon-controls.xlsm) {
    Copy-Item .\custom-ribbon-controls.xlsm output.zip
    Expand-Archive -Path .\output.zip -DestinationPath .\src1
    Remove-Item output.zip
    Remove-Item .\src 
    Rename-Item .\src1 .\src
}