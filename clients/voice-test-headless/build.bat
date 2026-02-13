@echo off
REM Build voice-test-headless.exe - offline transcription test harness (HTTP to local-ai-server)
setlocal EnableDelayedExpansion
cd /d "%~dp0"

where cl.exe >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
    for /f "usebackq tokens=*" %%i in (`"!VSWHERE!" -latest -property installationPath`) do set "VSINSTALL=%%i"
    call "!VSINSTALL!\VC\Auxiliary\Build\vcvarsall.bat" x64 >nul 2>&1
)

REM Paths relative to this build script (clients\voice-test-headless\)
set REPO_ROOT=..\..
set BUILD_DIR=%REPO_ROOT%\build\voice-test-headless
set BIN_DIR=%REPO_ROOT%\bin
set SHARED_DIR=%REPO_ROOT%\shared

if not exist "%BUILD_DIR%" mkdir "%BUILD_DIR%"
if not exist "%BIN_DIR%" mkdir "%BIN_DIR%"

echo Compiling asr_client...
cl /nologo /W3 /Od /Zi /DDEBUG /I"%SHARED_DIR%" /c "%SHARED_DIR%\asr_client.c" /Fo:"%BUILD_DIR%\asr_client.obj"
if %ERRORLEVEL% NEQ 0 exit /b 1

echo Compiling headless test...
cl /nologo /W3 /Od /Zi /I"%SHARED_DIR%" /c src\main.c /Fo:"%BUILD_DIR%\main.obj"
if %ERRORLEVEL% NEQ 0 exit /b 1

echo Linking...
link /nologo /DEBUG /SUBSYSTEM:CONSOLE /OUT:"%BIN_DIR%\voice-test-headless.exe" "%BUILD_DIR%\main.obj" "%BUILD_DIR%\asr_client.obj" winhttp.lib
if %ERRORLEVEL% NEQ 0 exit /b 1

echo Build complete: %BIN_DIR%\voice-test-headless.exe
