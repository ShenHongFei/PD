@echo off &title 论文检测
if "%1" == "h" goto begin 
mshta vbscript:createobject("wscript.shell").run("%~nx0 h",0)(window.close)&&exit 
:begin 
start /b java -jar -Dfile.encoding=UTF-8 -Dserver.port=8888 PD.jar >log.txt