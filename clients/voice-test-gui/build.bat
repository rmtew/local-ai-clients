@echo off
REM Build voice-test-gui.exe (ASR via HTTP to local-ai-server)
REM Can be run from any terminal and any directory

setlocal EnableDelayedExpansion
cd /d "%~dp0"

REM Auto-setup MSVC environment if not already configured
where cl.exe >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Setting up MSVC environment...
    set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
    if not exist "!VSWHERE!" (
        echo ERROR: vswhere.exe not found. Install Visual Studio with C++ workload.
        exit /b 1
    )
    for /f "usebackq tokens=*" %%i in (`"!VSWHERE!" -latest -property installationPath`) do set "VSINSTALL=%%i"
    if not defined VSINSTALL (
        echo ERROR: Could not find Visual Studio installation
        exit /b 1
    )
    call "!VSINSTALL!\VC\Auxiliary\Build\vcvarsall.bat" x64 >nul 2>&1
    where cl.exe >nul 2>&1
    if !ERRORLEVEL! NEQ 0 (
        echo ERROR: Failed to initialize MSVC environment
        exit /b 1
    )
)

REM Paths relative to this build script (clients\voice-test-gui\)
set REPO_ROOT=..\..
set BUILD_DIR=%REPO_ROOT%\build\voice-test-gui
set BIN_DIR=%REPO_ROOT%\bin
set SHARED_DIR=%REPO_ROOT%\shared

REM Create output directories
if not exist "%BUILD_DIR%" mkdir "%BUILD_DIR%"
if not exist "%BIN_DIR%" mkdir "%BIN_DIR%"

REM Compile shared asr_client
echo Compiling asr_client...
cl /nologo /W3 /Od /Zi /DDEBUG /I"%SHARED_DIR%" /c "%SHARED_DIR%\asr_client.c" /Fo:"%BUILD_DIR%\asr_client.obj"
if %ERRORLEVEL% NEQ 0 (
    echo asr_client compilation failed.
    exit /b 1
)

REM Compile GUI
echo Compiling GUI (debug)...
cl /nologo /W3 /Od /Zi /DDEBUG /I"%SHARED_DIR%" /c src\main.c /Fo:"%BUILD_DIR%\main.obj"
if %ERRORLEVEL% NEQ 0 (
    echo GUI compilation failed.
    exit /b 1
)

REM Link
echo Linking...
link /nologo /DEBUG /SUBSYSTEM:WINDOWS /OUT:"%BIN_DIR%\voice-test-gui.exe" "%BUILD_DIR%\main.obj" "%BUILD_DIR%\asr_client.obj" mfplat.lib mf.lib mfreadwrite.lib mfuuid.lib ole32.lib comctl32.lib user32.lib gdi32.lib winmm.lib winhttp.lib psapi.lib advapi32.lib

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build complete: %BIN_DIR%\voice-test-gui.exe
) else (
    echo.
    echo Build failed.
    exit /b 1
)
