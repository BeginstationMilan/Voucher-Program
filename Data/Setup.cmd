@echo off
setlocal

echo Do you have Python 3.11 installed? (Y/N)
set /p have_python=

if /i "%have_python%"=="Y" (
    echo Installing required modules using pip...
    pip install gspread oauth2client img2pdf pillow pystrich pywin32 tqdm
) else (
    echo Python 3.11 is required to run this script. Would you like to install it? (Y/N)
    set /p install_python=
    if /i "%install_python%"=="Y" (
        echo Opening Microsoft Store...
        start ms-windows-store://pdp/?productid=9NRWMJP3717K
        echo.
        echo Please install Python 3.11 from the Microsoft Store.
        echo After installation, press any key to continue...
        pause > nul
        echo Installing required modules using pip...
        pip install gspread colorama img2pdf oauth2client Pillow pywin32 qrcode python-barcode==0.13.1 pypiwin32


    ) else (
        echo Exiting...
        exit /b
    )
)

pause
