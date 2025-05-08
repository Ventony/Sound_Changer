@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
cls

:: ------------------------------------------------------------

:: 속성 창 열기
echo.
echo  =================================
echo   사운드 속성 창을 열고 있습니다.
echo  =================================
echo.
control mmsys.cpl
timeout /t 2 > nul
cls

:: ------------------------------------------------------------

:: 사용자 입력 받기
echo.
echo  =================================
echo     [사운드 장치 생성 도우미]
echo  =================================
echo  사용방법.txt를 참고하며 실행해주세요.
echo.
set /p deviceName=생성할 사운드 장치 이름을 입력하세요: 

:: 파일 이름에 사용할 안전한 이름 (공백을 밑줄로 대체)
set "safeName=%deviceName: =_%"

:: ------------------------------------------------------------

:: .vbs 파일 생성
echo Set WshShell = CreateObject("WScript.Shell") > "%safeName%.vbs"
echo WshShell.Run WshShell.CurrentDirectory ^& "\%safeName%.cmd", 0 >> "%safeName%.vbs"
echo Set WshShell = Nothing >> "%safeName%.vbs"
echo MsgBox "It's Change for %deviceName%" >> "%safeName%.vbs"

:: .cmd 파일 생성
echo nircmd.exe setdefaultsounddevice "%deviceName%" 1 > "%safeName%.cmd"
echo nircmd.exe setdefaultsounddevice "%deviceName%" 2 >> "%safeName%.cmd"


:: 바로가기 만들 VBS 파일 경로
set "targetFile=%CD%\%safeName%.vbs"

:: 바탕화면 바로가기 파일 경로
set "shortcutPath=%USERPROFILE%\Desktop\%safeName%.vbs.lnk"

:: VBS 스크립트로 바로가기 생성
echo Set WshShell = CreateObject("WScript.Shell") > createShortcut.vbs
echo Set shortcut = WshShell.CreateShortcut("%shortcutPath%") >> createShortcut.vbs
echo shortcut.TargetPath = "%targetFile%" >> createShortcut.vbs
echo shortcut.WorkingDirectory = "%CD%" >> createShortcut.vbs
echo shortcut.Save >> createShortcut.vbs

:: VBS 스크립트 실행
cscript //nologo createShortcut.vbs

:: VBS 파일 삭제
del createShortcut.vbs

:: ------------------------------------------------------------

:: 완료 메시지
cls
echo.
echo  =================================
echo        생성이 완료되었습니다.
echo  =================================
echo.
echo.
echo.
echo 5초 뒤에 이 창이 닫힙니다...
timeout /t 5 > nul
exit
