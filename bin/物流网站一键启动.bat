@echo off
set "EDGE=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
set "PROFILE=%USERPROFILE%\.config\edge-logistics"

if not exist "%EDGE%" (
    echo Edge not found at: %EDGE%
    pause
    exit /b 1
)

if not exist "%PROFILE%" (
    mkdir "%PROFILE%"
    echo First launch: opening 8 logistics websites...
    start "" "%EDGE%" --remote-debugging-port=9222 --user-data-dir="%PROFILE%" --no-first-run --no-default-browser-check http://nzhexp.nextsls.com/tms/wos/shipment?page=1^&pageSize=30^&activeTab=ready http://sfgjdl.nextsls.com/wos https://www.17track.net/zh-cn http://www.360vipwuliu.com/ https://yplogistics.com/?hmsr=wechat^&hmpl=^&hmcu=^&hmkw=^&hmci= http://smtgyl.nextsls.com/wos http://39.108.216.104:5001/#/tracking http://xmsdwl.nextsls.com/tracking/app#/tracking
) else (
    echo Restoring previous session...
    start "" "%EDGE%" --remote-debugging-port=9222 --user-data-dir="%PROFILE%" --no-first-run --no-default-browser-check
)

echo Done.
pause
