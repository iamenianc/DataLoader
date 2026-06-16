@echo off
REM ===========================================================================
REM  Builds a standalone Windows ExcelStage that needs NO .NET install to run.
REM
REM  It produces a FOLDER (zipped) rather than a single .exe, because the SQL
REM  Server networking library must sit next to the .exe as a real file - a
REM  single-file bundle crashes when it connects to SQL Server.
REM
REM  You only need the .NET 8 SDK installed to BUILD it once:
REM      winget install Microsoft.DotNet.SDK.8
REM  Then double-click this file (or run it from a command prompt).
REM ===========================================================================

echo Building ExcelStage ...
echo.

dotnet publish ExcelStage -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false -o dist\ExcelStage

if not exist "dist\ExcelStage\ExcelStage.exe" (
    echo.
    echo Build did not produce dist\ExcelStage\ExcelStage.exe - check the messages above.
    echo.
    pause
    exit /b 1
)

echo.
echo Zipping ...
powershell -NoProfile -Command "Compress-Archive -Path 'dist\ExcelStage\*' -DestinationPath 'dist\ExcelStage-windows.zip' -Force"

echo.
echo Done!
echo   Folder : %CD%\dist\ExcelStage
echo   Zip    : %CD%\dist\ExcelStage-windows.zip
echo.
echo Give the ZIP to users. They unzip it and run ExcelStage.exe inside - no install needed.
echo.
pause
