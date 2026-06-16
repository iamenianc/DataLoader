@echo off
REM ===========================================================================
REM  Builds a standalone Windows ExcelStage.exe (no .NET install needed to run).
REM  You only need the .NET 8 SDK installed to BUILD it once:
REM      winget install Microsoft.DotNet.SDK.8
REM  Then double-click this file (or run it from a command prompt).
REM  The finished program will be in the "dist" folder: dist\ExcelStage.exe
REM ===========================================================================

echo Building ExcelStage.exe ...
echo.

dotnet publish ExcelStage -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o dist

echo.
if exist "dist\ExcelStage.exe" (
    echo Done!  Your program is here:  %CD%\dist\ExcelStage.exe
    echo Copy that single file to any Windows PC and run it - no install needed.
) else (
    echo Build did not produce dist\ExcelStage.exe - check the messages above.
)

echo.
pause
