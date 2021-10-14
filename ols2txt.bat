@echo off & title Converting  files from ols to txt

call :MakeVBS "%~0"

echo #--------*--------*--------*--------*--------*--------*--------*--------*--------#
echo *        Start converting ols files to txt format
::echo *            %date:~,4% - %date:~5,2% - %date:~8,2%  %time:~0,2% : %time:~3,2% : %time:~6,2%
echo #--------*--------*--------*--------*--------*--------*--------*--------*--------#>>ols2txt_log.txt
echo *        Start converting ols files to txt format>> ols2txt_log.txt

for /r %%a in (*.ols) do (
    ::cls
    call get_time.bat
    call get_time.bat >>ols2txt_log.txt
    echo *            Converting:"%%~a" 
    echo *            Converting:"%%~a" >>ols2txt_log.txt
    
    ols2txt.vbs "%%~a"

)
::cls
call get_time.bat
call get_time.bat >>ols2txt_log.txt

:: Comment
::objWorksheet.SaveAs csvPath & "\" & fso.GetBaseName(xlsxPath)&"_sheet" &index& ".csv" , xlCSV
::objWorksheet.SaveAs csvPath & "\" & fso.GetBaseName(xlsxPath)& ".csv" , xlCSV

::echo *            %date:~,4% - %date:~5,2% - %date:~8,2%  %time:~0,2% : %time:~3,2% : %time:~6,2%
::echo *            %date:~,4% - %date:~5,2% - %date:~8,2%  %time:~0,2% : %time:~3,2% : %time:~6,2%>>ols2txt_log.txt
echo *        Finished converting all ols files in  current path to txt format
echo *        Finished converting all ols files in current path to txt format >> ols2txt_log.txt
del ols2txt.vbs
pause
exit



:MakeVBS
for /f "tokens=1 delims=[]" %%a in ('find /n "::ols2txt::" "%~1"') do set HH=%%~a
more +%HH% "%~1">ols2txt.vbs
goto :eof

:: VBS code

::ols2txt::
Set fso=CreateObject("Scripting.FileSystemObject")
olsPath = WScript.Arguments(0)
txtPath = fso.GetFile(olsPath).ParentFolder.Path 
set openfile = fso.opentextfile(olsPath,1,true)
Dim index
For index=1 to 7
openfile.Skipline()
Next
set testfile = fso.createtextfile(txtPath & "\" & fso.GetBaseName(olsPath)& ".txt" ,true)

testfile.writeline(openfile.readall)

testfile.close

