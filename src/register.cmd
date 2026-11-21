@echo off
rem 管理员身份运行；脚本所在目录即 DLL 目录
set DLL_DIR=%~dp0
set DLL_NAME=BetterWPP.dll
set PROGID=BetterWPP.Connect
set CLSID={11111111-1111-1111-1111-111111111111}

rem 1) COM 注册（写入 HKCR\CLSID，指向 DLL）
regasm /codebase "%DLL_DIR%%DLL_NAME%"

rem 2) 向 PowerPoint 登记 COM 加载项（HKCU，免管理员）；通用键 + 16.0 版本键都写
for %%K in (
  "HKCU\Software\Microsoft\Office\PowerPoint\Addins\%PROGID%"
  "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Addins\%PROGID%"
) do (
  reg add %%K /v FriendlyName /t REG_SZ /d "BetterWPP" /f
  reg add %%K /v Description /t REG_SZ /d "BetterWPP COM add-in" /f
  reg add %%K /v LoadBehavior /t REG_DWORD /d 3 /f
  reg add %%K /v CommandLineSafe /t REG_DWORD /d 0 /f
)

echo Registered %PROGID%. Restart PowerPoint.
pause
