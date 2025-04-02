@ECHO OFF

signtool sign /f "EnricoSadlowski_2024-10-25_12-36-34.pfx" /p "300675s70931_#" /fd SHA256 /t http://timestamp.digicert.com "Zeiterfassung.exe"
PAUSE