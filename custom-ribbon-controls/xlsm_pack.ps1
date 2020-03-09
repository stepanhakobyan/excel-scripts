if (Test-Path .\custom-ribbon-controls.xlsm) {
    Remove-Item .\custom-ribbon-controls.xlsm
}
if (Test-Path .\output.zip) {
    Remove-Item .\output.zip
}
Compress-Archive -Path ./src/* -DestinationPath output.zip
Rename-Item .\output.zip .\custom-ribbon-controls.xlsm