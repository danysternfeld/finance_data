REM Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
python make_version_file.py
powershell "del -r -force  build"
pyinstaller --clean -F .\make_performance_xl.py
del *.exe
move .\dist\make_performance_xl.exe .
del *.zip
powershell -Command "Compress-Archive  -path  *exe ,excel_files , .\Oct_2026\  -destinationpath analysis.zip"