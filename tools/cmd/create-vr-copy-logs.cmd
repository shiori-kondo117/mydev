@echo off
@rem //****************************************************
@rem //* @procedure proc-Begin
@rem //****************************************************
	setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
:proc-Begin
	@rem -- import profile
	call %~dp0profile.cmd
	set vCmdHome=%~dp0
	if "%~1" == "" (
		echo Usage:%~nx0 [src-dir]
		set vRc=1
		goto :proc-End
	)

@rem --	if not exist %~dpnx1\DVD_RTAV (
@rem --		call :proc-RunLog "E" "not exist [DVD_RTAV]!! [src-dir:%~1]"
@rem --		goto :proc-End
@rem --	)
	
	set vDestDir=%~dpnx1
	set /p vSrcDrive=please drive-name[{K:^|R:^|L:}]:

@rem //****************************************************
@rem //* @procedure proc-Main
@rem //****************************************************
:proc-Main
@rem --	set vDataHome=C:\Users\skybl\Documents\mydata\db\csv
	set vDataFile=vr_copy_logs.csv

	if not exist %vCsvHome%\%vDataFile% (
		echo id,src_drive,vol_name,dest_dir,created> %vDataHome%%vDataFile%
		set vId=1
	) else (
		for /f "tokens=1 delims=," %%i in (%vCsvHome%\%vDataFile%) do (
			set vIdMax=%%i
		)
		
		set /a vId = !!vIdMax!! + 1
	)

	set vCmd=get-vol.cmd
	call %vCmdHome%%vCmd% %~d1 vDriveName_ vVolName_ vVolSer_
	
	if %errorlevel% neq 0 (
		call :proc-RunLog "E" "commadn execute failler!! [rc=%errorlevel%]"
		set vRc=1
		goto :proc-End
	)

	set vDateTime=%date: =0% %time: =0%
	
	set vCsvData=%vId%,"%vSrcDrive%","%vVolName_%","%vDestDir%","%vDateTime%"
	
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

