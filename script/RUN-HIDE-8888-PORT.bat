@echo off &title ÂÛÎÄ¼ì²â
if "%1" == "h" goto begin 
mshta vbscript:createobject("wscript.shell").run("%~nx0 h",0)(window.close)&&exit 
:begin 
start /b java -jar -Dfile.encoding=UTF-8 -Dserver.port=8888 PD.jar >log.txt