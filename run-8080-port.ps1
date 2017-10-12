$host.ui.RawUI.WindowTitle = "论文检测"
Set-Location $PSScriptRoot
[Environment]::CurrentDirectory=$PSScriptRoot
Write-Host "工作目录: $PWD"
java --% -jar -Dfile.encoding=UTF-8 -Dserver.port=8080 PD.jar
pause