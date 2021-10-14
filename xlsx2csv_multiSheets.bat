@echo off & title Converting  files from Xlsx to Csv

call :MakeVBS "%~0"

echo #--------*--------*--------*--------*--------*--------*--------*--------*--------#
echo *        Start converting xlsx files to csv format
::echo *            %date:~,4% - %date:~5,2% - %date:~8,2%  %time:~0,2% : %time:~3,2% : %time:~6,2%
echo #--------*--------*--------*--------*--------*--------*--------*--------*--------#>>xlsx2csv_multiSheets_log.txt
echo *        Start converting xlsx files to csv format>> xlsx2csv_multiSheets_log.txt

for /r %%a in (*.xlsx) do (
    ::cls
    call get_time.bat
    call get_time.bat >>xlsx2csv_multiSheets_log.txt
    echo *            Converting:"%%~a" 
    echo *            Converting:"%%~a" >>xlsx2csv_multiSheets_log.txt
    Xlsx2Csv.vbs "%%~a"
)
::cls
call get_time.bat
call get_time.bat >>xlsx2csv_multiSheets_log.txt
::echo *            %date:~,4% - %date:~5,2% - %date:~8,2%  %time:~0,2% : %time:~3,2% : %time:~6,2%
::echo *            %date:~,4% - %date:~5,2% - %date:~8,2%  %time:~0,2% : %time:~3,2% : %time:~6,2%>>xlsx2csv_multiSheets_log.txt
echo *        Finished converting all xlsx files in  current path to csv format
echo *        Finished converting all xlsx files in current path to csv format >> xlsx2csv_multiSheets_log.txt
del Xlsx2Csv.vbs
pause
exit


:MakeVBS
for /f "tokens=1 delims=[]" %%a in ('find /n "::Xlsx2Csv::" "%~1"') do set HH=%%~a
more +%HH% "%~1">Xlsx2Csv.vbs
goto :eof

:: VBS code

::Xlsx2Csv::
const xlCSV = 6
Set fso=CreateObject("Scripting.FileSystemObject")
xlsxPath = WScript.Arguments(0)
csvPath = fso.GetFile(xlsxPath).ParentFolder.Path 
Set objExcel = CreateObject("Excel.Application")
Set objWorkBook = objExcel.WorkBooks.Open(xlsxPath)
objExcel.DisplayAlerts = FALSE
objExcel.Visible = FALSE
Dim index
index=1
For Each objWorksheet in objWorkBook.Worksheets
objWorksheet.SaveAs csvPath & "\" & fso.GetBaseName(xlsxPath)&"_sheet" &index& ".csv" , xlCSV
objExcel.ActiveWorkBook.Saved=True
index=index+1
Next
objExcel.Quit