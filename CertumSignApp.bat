@echo off

:: === SIGNIERUNGSKONFIGURATION ===
set SIGNTOOL="C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe"
set CERT_DIR="C:\Zertifikate\Certum"
set PFX="%CERT_DIR%\EnricoSadlowski_2025.pfx"
set PASSWORD=300675s70931_#
set EXE="Output\Zeiterfassung.exe"
set TIMEURL=http://timestamp.digicert.com

:: === BACKUP DES ZERTIFIKATS ===
echo Erstelle verschlüsseltes Backup des PFX...
set BACKUP_DIR="%USERPROFILE%\Documents\Zertifikat_Backup"
mkdir %BACKUP_DIR%
copy /Y %PFX% %BACKUP_DIR%\EnricoSadlowski_2025_backup.pfx >nul

:: Optional: in passwortgeschütztes ZIP packen (nur mit 7-Zip)
if exist "C:\Program Files\7-Zip\7z.exe" (
    "C:\Program Files\7-Zip\7z.exe" a -p%PASSWORD% "%BACKUP_DIR%\EnricoSadlowski_2025_backup.zip" "%PFX%"
)

:: === SIGNIERUNG STARTET ===
echo Signiere: %EXE%
%SIGNTOOL% sign /f %PFX% /p %PASSWORD% /tr %TIMEURL% /td sha256 /fd sha256 %EXE%

if %errorlevel% neq 0 (
    echo FEHLER: Signierung fehlgeschlagen.
    pause
    exit /b %errorlevel%
) else (
    echo Erfolgreich signiert: %EXE%
)

pause