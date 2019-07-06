@echo off
@rem //****************************************************
@rem //* @procedure proc-Begin
@rem //****************************************************
:proc-Begin
	setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
	@rem -- import profile
	call %~dp0profile.cmd
	if "%~1" == "" (
		echo Usage:%~nx0 [src-dir]
		set vRc=1
		goto :proc-End
	)

@rem --	if not exist %~dpnx1\DVD_RTAV (
@rem --		call :proc-RunLog "not exist [DVD_RTAV]!! [src-dir:%~1]"
@rem --		goto :proc-End
@rem --	)
	
	@rem -- get drive name
	set vCmdHome=%~dp0
	set vCmd=get-vol.cmd
	set vDriveName=%~d1

	call %vCmdHome%%vCmd% %vDriveName% vDriveName_ vVolName_ vVolSer_
	
	set vRc=%errorlevel%
	if %vRc% neq 0 (
		call :proc-RunLog "E" "command execute failer!! [cmd=%vCmd% rc=%vRc%]"
		goto :proc-End
	)
	
	set vSrcDir=%~dpnx1

@rem //****************************************************
@rem //* @procedure proc-Main
@rem //****************************************************
:proc-Main
@rem --	set vDataHome=C:\Users\skybl\Documents\mydata\db\csv
	set vDataFile=vr_edit_logs.csv

	if not exist %vCsvHome%\%vDataFile% (
		echo id,vol_name,src_path,created> %vDataHome%\%vDataFile%
		set vId=1
	) else (
		for /f "tokens=1 delims=," %%i in (%vCsvHome%%vDataFile%) do (
			set vIdMax=%%i
		)
		
		set /a vId = !!vIdMax!! + 1
	)

	set vDateTime=%date: =0% %time: =0%
	
	set vCsvData=%vId%,"%vVolName_%","%vSrcDir%","%vDateTime%"
	
	echo %vCsvData%>> %vCsvHome%%vDataFile%

	if not "%~2" == "" (
		shift /1
		goto :proc-Begin
	)

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

