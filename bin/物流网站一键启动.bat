@echo off
set "PROFILE=%USERPROFILE%\.config\edge-logistics"

:: 检测 9222 端口是否已被占用（说明物流 Edge 已在运行）
netstat -an | find ":9222" | find "LISTENING" >nul
if %errorlevel% equ 0 (
    echo Logistics Edge is already running. Close it first before restarting.
    pause
    exit /b
)

if not exist "%PROFILE%" mkdir "%PROFILE%"
rd /s /q "%PROFILE%\Default\Sessions" 2>nul
start "" msedge --remote-debugging-port=9222 --user-data-dir="%PROFILE%" --no-first-run --no-default-browser-check http://nzhexp.nextsls.com/tms/wos/shipment?page=1^&pageSize=30^&activeTab=ready http://sfgjdl.nextsls.com/wos https://www.17track.net/zh-cn http://www.360vipwuliu.com/ https://yplogistics.com/?hmsr=wechat^&hmpl=^&hmcu=^&hmkw=^&hmci= http://smtgyl.nextsls.com/wos http://39.108.216.104:5001/#/tracking http://xmsdwl.nextsls.com/tracking/app#/tracking
pause
