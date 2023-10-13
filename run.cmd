@set DataPath=z:\Save\TSS\Reporting\MBR
@set PyPath=..\..\Template\Scripts
@set /p month=<Data-IN\month.txt
@set /p year=<Data-IN\year.txt

@if "%year%" == ""  (
	echo Year nicht gesetzt!
	goto ende
)

@python "%PyPath%\Export.py" 
@if not ERRORLEVEL 0 goto ERROR 

rem "%DataPath%\%year%\%month%" %month% %year%

@python "%PyPath%\PrepareTemplate.py" "%DataPath%\%year%\%month%" %month% %year%
@if not ERRORLEVEL 0 goto ERROR 

@python "%PyPath%\ProcessData.py" "%DataPath%\%year%\%month%" %month% %year%
@if not ERRORLEVEL 0 goto ERROR

@python "%PyPath%\BuildCharts.py" "%DataPath%\%year%\%month%" %month% %year%
@if not ERRORLEVEL 0 goto ERROR

@python "%PyPath%\UpdatePPT.py" "%DataPath%\%year%\%month%" %month% %year%
@if not ERRORLEVEL 0 goto ERROR

@move /Y *.csv tmp
@move /Y *.svg tmp
@move /Y *.png tmp

mkdir \\di-daten\verwaltung\GL\TSS\MBR\%year%\%month%\
xcopy * \\di-daten\verwaltung\GL\TSS\MBR\%year%\%month%\  /S
del TTE-MBR-%year%-%month%_leer.pptx

@echo ---------- Script fertig ----------
@goto END


:ERROR
@echo ********** ERROR  Script abgebrochen ***********


:END
pause