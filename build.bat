@echo off
REM === Activate Virtual Environment and Build PDF Letterhead Merger ===

call .venv\Scripts\activate.bat

echo Cleaning old build files...
rmdir /s /q "_output"
mkdir "_output"

echo Running PyInstaller...
python -m PyInstaller --name "PDF Letterhead Merger" ^
                      --onefile ^
                      --windowed ^
                      --icon="icon.ico" ^
                      --add-data "..\\icon.ico;." ^
                      --distpath "./_output" ^
                      --workpath "./_output/build" ^
                      --specpath "./_output" ^
                      --clean ^
                      main.py

echo.
echo === Build Complete ===
echo Executable is located in: _output\PDF Letterhead Merger.exe
pause
