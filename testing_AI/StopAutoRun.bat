@echo off
ren C:\testing_AI\AutoRun.bat AutoRun_.bat
del \F C:\testing_AI\logs\wait.txt
Goto :End
:End
taskkill /FI "WINDOWTITLE ne END" /IM cmd.exe /F /T
