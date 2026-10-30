@echo off
rem Rebuilds xlpolish.jar in the repository root from java\src (needs a JDK 9 or newer).
setlocal
set HERE=%~dp0
set BUILD=%TEMP%\xlpolish_build
if exist "%BUILD%" rmdir /s /q "%BUILD%"
mkdir "%BUILD%"
javac --release 8 -Xlint:all -Xlint:-options -d "%BUILD%" -sourcepath "%HERE%stub" "%HERE%src\xlpolish\*.java" || exit /b 1
rmdir /s /q "%BUILD%\com"
jar cf "%HERE%..\xlpolish.jar" -C "%BUILD%" xlpolish || exit /b 1
rmdir /s /q "%BUILD%"
echo built %HERE%..\xlpolish.jar
