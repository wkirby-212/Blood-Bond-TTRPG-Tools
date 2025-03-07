@echo off
echo Creating DialogSpellMaker shortcut on your desktop...

:: Get the current directory where this batch file is located
set "CURRENT_DIR=%~dp0"
set "PS_SCRIPT=%CURRENT_DIR%DialogSpellMaker.ps1"
set "ICON_FILE=%CURRENT_DIR%full_logo.ico"
set "DESKTOP_DIR=%USERPROFILE%\Desktop"
set "SHORTCUT_NAME=DialogSpellMaker.lnk"

:: Create a temporary VBScript to make the shortcut
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%TEMP%\CreateShortcut.vbs"
echo sLinkFile = "%DESKTOP_DIR%\%SHORTCUT_NAME%" >> "%TEMP%\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%TEMP%\CreateShortcut.vbs"
echo oLink.TargetPath = "powershell.exe" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.Arguments = "-ExecutionPolicy Bypass -File ""%PS_SCRIPT%""" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.WorkingDirectory = "%CURRENT_DIR%" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.IconLocation = "%ICON_FILE%" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.Save >> "%TEMP%\CreateShortcut.vbs"

:: Run the VBScript to create the shortcut
cscript //nologo "%TEMP%\CreateShortcut.vbs"

:: Delete the temporary VBScript
del "%TEMP%\CreateShortcut.vbs"

echo.
echo Shortcut created successfully on your desktop!
echo.
pause
