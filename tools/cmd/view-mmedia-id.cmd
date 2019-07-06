@echo off
@rem //****************************************************
@rem //* @procedure proc-Begin
@rem //****************************************************
:proc-Begin
	set vRc=
	set /p vMediaType="select media-type?(bd|bd-re|cd|cd-rw|cd-rw|dvd|dvd-rw|dvd+rw|dvd-ram|cfm|mmc|ms|hd|sd|ssd|vhd):"
	
	echo %vMediaType%|findstr "bd bd-re cd cd-rw dvd dvd-rw dvd+rw dvd-ram cfm mmc ms hd sd ssd vhd"
	
	if %errorlevel% neq 0 (
		call :proc-RunLog "E" "illigal media-type !! [media-type=%vMediaType%]"
		set vRc=1
		goto :proc-End
	)
	
	set vCmdHome=%~dp0
	set vCmd=get-id-data-file-path.cmd
	
	call %vCmdHome%%vCmd% %vMediaType% vDataFile
	
	if %errorlevel% neq 0 (
		call :proc-RunLog "E" "get file path [data-id-path] failler!! [media-type=%vMediaType%]"
		set vRc=2
		goto :proc-End
	)
	
	set vDataId_=""

	for /f "tokens=1,2 delims==" %%i in (%vDataFile%) do (
		if "%%i" == "id" (
			set vDataId_=%%j
		)
	)
	
	if %vDataId_% == "" (
		set vRc=1
		call :proc-RunLog "E" "get [data-id] failler!!"
		goto :proc-End
	)
		
@rem //****************************************************
@rem //* @procedure proc-Main
@rem //****************************************************
:proc-Main
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
	
	@rem -- view data id
	echo MM-DATA-ID:%vDataId_%
	
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

