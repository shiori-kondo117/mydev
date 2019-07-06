@echo off
@rem //****************************************************
@rem //* @procedure proc-Begin
@rem //****************************************************
:proc-Begin
	set vRc=

	if "%~1" == "" (
		call :proc-RunLog "E" "Usage:%~nx0 [dest-dir(create directory)]"
		set vRc=1
		goto :proc-End
	)
	
	set vDriveName=%~d1
	
	if not exist %vDriveName% (
		call :proc-RunLog "E" "get [drive-name] failler!! [src-dir=%~1]"
		set vRc=2
		goto :proc-End
	)
	
	set vCmdHome=%~dp0
	set vCmd=get-vol.cmd

	call %vCmdHome%%vCmd% %vDriveName% "vDriveName_" "vVolName_" "vVolSer_"
	
	if %errorlevel% neq 0 (
		call :proc-RunLog "E" "command execute failler!! [cmd=%vCmd%]"
		set vRc=3
		goto :proc-End
	)
	
	set vDestDir=%~dp1\

	set vCmd=get-id-data-hd-file-path.cmd
	
	call %vCmdHome%%vCmd% %vVolName_% vDataFile
	
	if %errorlevel% neq 0 (
		call :proc-RunLog "E" "get file path [data-id-path] failler!! [media-type=%vMediaType%]"
		set vRc=4
		goto :proc-End
	)
	
	set vDataId_=""

	for /f "tokens=1,2 delims==" %%i in (%vDataFile%) do (
		if "%%i" == "id" (
			set vDataId_=%%j
		)
	)
	
	if %vDataId_% == "" (
		set vRc=5
		call :proc-RunLog "E" "Illigal(id=xxx) data file!! [data-file=%vDataFile%]"
		goto :proc-End
	)
@rem //****************************************************
@rem //* @procedure proc-Main
@rem //****************************************************
:proc-Main
	set /a vDataId_+=1
	set vDataId=%vDataId_%
	
	@rem -- padding zeros
	if %vDataId_% lss 10 (
		set vDataId_=0000%vDataId_%
	) else if %vDataId_% lss 100 (
		set vDataId_=000%vDataId_%
	) else if %vDataId_% lss 1000 (
		set vDataId_=00%vDataId_%
	) else if %vDataId_% lss 10000 (
		set vDataId_=0%vDataId_%
	)
	
	@rem -- create directory
	mkdir %vDestDir%%vDataId_%
	
	@rem -- data id file update
	set vDataLine=id=%vDataId%
	echo %vDataLine% >%vDataFile%
	
@rem //****************************************************
@rem //* @procedure proc-End
@rem //****************************************************
:proc-End
	pause
	if defined vRc (
		exit /b %vRc%
	)
	goto :eof

@rem //****************************************************
@rem //* @procedure proc-RunLog [arg1] [arg2]
@rem //* @arg1 msg-type([I]nformation|[E]rror^|[W]arning|[F]atal)
@rem //* @arg2 msg-log
@rem //****************************************************
:proc-RunLog
	set vDate=%date: =0%
	set vTime=%time: =0%
	echo %~1,"%vDate% %vTime%","%~2"
	goto :eof

