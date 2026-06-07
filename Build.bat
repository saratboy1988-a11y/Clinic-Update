@echo off
chcp 65001 >nul
title Clinic Management System Builder
echo ============================================================
echo        BUILDING CLINIC MANAGEMENT SYSTEM TO EXE
echo ============================================================
echo.

:: ១. ពិនិត្យ និងដំឡើង PyInstaller ប្រសិនបើមិនទាន់មាន
echo [*] Checking for PyInstaller...
pip install pyinstaller --quiet
if errorlevel 1 (
    echo [ERROR] Failed to install PyInstaller!
    pause
    exit /b 1
)

:: ២. លុប Folder ចាស់ៗដែលធ្លាប់ Build មុនចេញ
echo [*] Cleaning old build files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ClinicManager.spec del /f /q ClinicManager.spec

:: ៣. បង្កើត settings.ini ជាមួយ UTF-8 BOM (បើមិនទាន់មាន)
echo [*] Ensuring settings.ini has UTF-8 BOM...
python -c "import codecs; f=open('settings.ini', 'rb'); content=f.read(); f.close(); f=open('settings.ini', 'wb'); f.write(codecs.BOM_UTF8+content if not content.startswith(codecs.BOM_UTF8) else content); f.close()" 2>nul || echo Warning: Could not add BOM to settings.ini

:: ៣. ចាប់ផ្តើម Build ទៅជា EXE
:: --noconsole: មិនបង្ហាញផ្ទាំង CMD ខ្មៅពេលបើកកម្មវិធី
:: --onedir: រក្សា DLLs ក្នុង folder កម្មវិធី ដើម្បីជៀសវាងបញ្ហា _MEI Temp extraction
:: --name: កំណត់ឈ្មោះកម្មវិធី
:: --add-data: បញ្ចូល files ថែម (semicolon សម្រាប់ Windows)
echo [*] Compiling main.py to application folder (Please wait, this may take a few minutes)...
pyinstaller --noconsole --onedir --name "ClinicManager" ^
--hidden-import=PyQt5.QtCore ^
--hidden-import=PyQt5.QtGui ^
--hidden-import=PyQt5.QtWidgets ^
--hidden-import=matplotlib.backends.backend_qt5agg ^
--hidden-import=sqlite3 ^
--hidden-import=hashlib ^
--hidden-import=logging ^
--hidden-import=configparser ^
--hidden-import=openpyxl ^
--hidden-import=reportlab ^
--hidden-import=docx ^
--hidden-import=matplotlib.pyplot ^
--hidden-import=encodings.utf_8_sig ^
--hidden-import=certifi ^
--add-binary "C:\Users\PC\AppData\Local\Programs\Python\Python310\vcruntime140.dll;." ^
--add-binary "C:\Users\PC\AppData\Local\Programs\Python\Python310\vcruntime140_1.dll;." ^
--add-data "settings.ini;." ^
--add-data "template.xlsx;." ^
--add-data "license_manager.py;." ^
--add-data "db.py;." ^
--add-data "excel_handler.py;." ^
--add-data "report_handler.py;." ^
--add-data "ui_utils.py;." ^
--add-data "widgets.py;." ^
--add-data "constants.py;." ^
--add-data "license_server_config.json;." ^
main.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed! Please check the error messages above.
    echo ============================================================
    pause
    exit /b 1
)

:: ៤. ពិនិត្យមើលថា EXE ត្រូវបានបង្កើតជោគជ័យឬនៅ
if not exist dist\ClinicManager\ClinicManager.exe (
    echo.
    echo [ERROR] ClinicManager.exe was not created!
    echo ============================================================
    pause
    exit /b 1
)

:: ៥. ចម្លងឯកសារ Assets ទៅក្នុង Folder dist (សម្រាប់ development)
echo [*] Copying necessary files to dist folder...
copy /y settings.ini dist\ClinicManager\ >nul
copy /y template.xlsx dist\ClinicManager\ >nul
copy /y license_manager.py dist\ClinicManager\ >nul
copy /y db.py dist\ClinicManager\ >nul
copy /y excel_handler.py dist\ClinicManager\ >nul
copy /y report_handler.py dist\ClinicManager\ >nul
copy /y ui_utils.py dist\ClinicManager\ >nul
copy /y widgets.py dist\ClinicManager\ >nul
copy /y constants.py dist\ClinicManager\ >nul
copy /y license_server_config.json dist\ClinicManager\ >nul

echo.
echo ============================================================
echo   SUCCESS: Build completed! Check the 'dist' folder.
echo ============================================================
echo.
echo Files created:
echo   - dist\ClinicManager\ClinicManager.exe (Main executable)
echo.
echo IMPORTANT: The following files are already embedded in the EXE:
echo   - settings.ini
echo   - template.xlsx
echo   - All Python modules
echo.
echo OPTIONAL: Copy clinic.db to dist\ if you want to keep existing data
echo ============================================================
echo.
pause
