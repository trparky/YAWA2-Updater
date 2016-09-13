MSBuild.exe "YAWA2 Updater.sln" /noconsolelogger /t:Rebuild /p:Configuration=Debug
MSBuild.exe "YAWA2 Updater.sln" /noconsolelogger /t:Rebuild /p:Configuration=Release
cd "WinApp.ini Updater\bin\Release"
"Package This Program!.bat"