@echo off
set "PROJECT_DIR=Arcxis Cad Tools - Brics - Mechanical"
set "PROJECT_FILE=%PROJECT_DIR%\Arcxis Cad Tools - Brics - Mechanical.vbproj"
set "SOLUTION_FILE=Arcxis Cad Tools - Brics - Mechanical.sln"

echo ========================================
echo Clearing NuGet Cache and Restoring Packages
echo ========================================
echo.
echo WARNING: Make sure Visual Studio is CLOSED before running this script!
echo.
pause

echo.
echo Step 1: Clearing NuGet cache...
echo ----------------------------------------
dotnet nuget locals all --clear
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to clear NuGet cache. Make sure Visual Studio is closed!
    pause
    exit /b 1
)

echo.
echo Step 2: Deleting bin and obj folders...
echo ----------------------------------------
if exist "%PROJECT_DIR%\bin\" (
    echo Deleting bin folder...
    rmdir /s /q "%PROJECT_DIR%\bin"
)

if exist "%PROJECT_DIR%\obj\" (
    echo Deleting obj folder...
    rmdir /s /q "%PROJECT_DIR%\obj"
)

echo.
echo Step 3: Restoring NuGet packages...
echo ----------------------------------------
dotnet restore "%SOLUTION_FILE%"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to restore NuGet packages!
    pause
    exit /b 1
)

echo.
echo Step 4: Building the project...
echo ----------------------------------------
dotnet build "%SOLUTION_FILE%" --configuration Debug -p:Platform=x64
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo SUCCESS! Cache cleared and project rebuilt.
echo ========================================
echo.
echo You can now reopen Visual Studio.
pause
