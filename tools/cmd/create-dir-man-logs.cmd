@echo off
@rem //****************************************************
@rem //* @procedure proc-Begin
@rem //****************************************************
:proc-Begin
	setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
	if "%~1" == "" (
		echo Usage:%~nx0 [data-file] [status-file] [src-dir]
		set vRc=1
		goto :proc-End
	)

	if "%~2" == "" (
		echo Usage:%~nx0 [data-file] [status-file] [src-dir]
		set vRc=2
		goto :proc-End
	)

	if "%~3" == "" (
		echo Usage:%~nx0 [data-file] [status-file] [src-dir]
		set vRc=3
		goto :proc-End
	)

@rem --	if not exist %~dpnx1\DVD_RTAV (
@rem --		call :proc-RunLog "not exist [DVD_RTAV]!! [src-dir:%~1]"
@rem --		goto :proc-End
@rem --	)
	
	set vDataFile=%~1
	set vStatusFile=%~2
	set vSrcDir=%~dpnx3

@rem //****************************************************
@rem //* @procedure proc-Main
@rem //****************************************************
:proc-Main
	set vDataHome=C:\Users\skybl\Documents\mydata\db\csv

	if not exist %vDataHome%\%vDataFile% (
		echo id,src_path,created> %vDataHome%\%vDataFile%
		echo id,ended> %vDataHome%\%vStatusFile%
		set vId=1
	) else (
		for /f "tokens=1 delims=," %%i in (%vDataHome%\%vDataFile%) do (
			set vIdMax=%%i
		)
		
		set /a vId = !!vIdMax!! + 1
	)

	set vDateTime=%date: =0% %time: =0%
	
	set vCsvData=%vId%,"%vSrcDir%","%vDateTime%"
	
	echo %vCsvData%>> %vDataHome%\%vDataFile%

@rem //****************************************************
@rem //* @procedure proc-End
@rem //****************************************************
:proc-End
	pause
	if defined vRc (
		exit /b %vRc%
	)
	endlocal
	goto :eof

@rem //****************************************************
@rem //* @procedure proc-RunLog [arg1]
@rem //* @arg1 msg-log
@rem //****************************************************
:proc-RunLog
	set vDate=%date: =0%
	set vTime=%time: =0%
	echo %~1,"%vDate% %vTime%","%~2"
	goto :eof

