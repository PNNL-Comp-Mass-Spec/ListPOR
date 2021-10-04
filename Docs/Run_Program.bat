@echo off

set ExePath=ListPOR.exe

if exist %ExePath% goto DoWork
if exist ..\%ExePath% set ExePath=..\%ExePath% && goto DoWork
if exist ..\Bin\%ExePath% set ExePath=..\Bin\%ExePath% && goto DoWork
if exist ..\Bin\Debug\%ExePath% set ExePath=..\Bin\Debug\%ExePath% && goto DoWork

echo Executable not found: %ExePath%
goto Done

:DoWork
echo.
echo Procesing with %ExePath%
echo.

%ExePath% TestData.txt /conf:95 TestData_filtered_95pct.txt
%ExePath% TestData.txt /conf:97 TestData_filtered_97pct.txt
%ExePath% TestData.txt /conf:99 TestData_filtered_99pct.txt

:Done

pause
