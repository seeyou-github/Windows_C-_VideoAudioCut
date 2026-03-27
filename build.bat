@echo off
setlocal

echo ===== Build VideoAudioCut =====

if not exist build mkdir build

set CXX=g++
set RC=windres
set CXXFLAGS=-std=c++17 -DUNICODE -D_UNICODE -Wall -Wextra
set LDFLAGS=-municode -mwindows -lgdi32 -lcomctl32 -lshell32 -lole32 -loleaut32 -luuid -ldwmapi -luxtheme

%CXX% -c src\main.cpp %CXXFLAGS% -o build\main.o
if errorlevel 1 goto error

%CXX% -c src\MainWindow.cpp %CXXFLAGS% -o build\MainWindow.o
if errorlevel 1 goto error

%CXX% -c src\Config.cpp %CXXFLAGS% -o build\Config.o
if errorlevel 1 goto error

%CXX% -c src\FFmpegRunner.cpp %CXXFLAGS% -o build\FFmpegRunner.o
if errorlevel 1 goto error

%CXX% -c src\ProcessUtils.cpp %CXXFLAGS% -o build\ProcessUtils.o
if errorlevel 1 goto error

%CXX% -c src\StringResources.cpp %CXXFLAGS% -o build\StringResources.o
if errorlevel 1 goto error

%RC% src\res\resource.rc -O coff -o build\resource.o
if errorlevel 1 goto error

echo Linking...
%CXX% build\main.o build\MainWindow.o build\Config.o build\FFmpegRunner.o build\ProcessUtils.o build\StringResources.o build\resource.o -o build\VideoAudioCut.exe %LDFLAGS%
if errorlevel 1 goto error

echo ===== Build succeeded =====
echo Output: build\VideoAudioCut.exe
goto end

:error
echo ===== Build failed =====
exit /b 1

:end
endlocal
