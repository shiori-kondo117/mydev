@echo off
:proc-Init
	@rem -- import profile
	call %~dp0profile.cmd
	if "%~1" == "" (
		echo Usage:%~nx0 [src-dir]
		exit /b 1
	)
	
	set vCmdHome=%~dp0

	set vDestDir=%vCsvHome%
	set vDestFile=mmedia_file_lists.csv
	set vDestPath=%vDestDir%%vDestFile%
	
	if not exist "%vDestPath%" (
		set vCsvHead="id,file_name,vol_name,vol_ser,path,attributes_map,created,file_size,file_ext"
		echo %vCsvHead:"=% > "%vDestPath%"
	)
	
	set vCmd=get-vol.cmd
	set vDriveName=%~d1
	call %vCmdHome%%vCmd% %vDriveName% vDriveName_ vVolName_ vVolSer_
	
	if %errorlevel% neq 0 (
		echo "E" "command failer!! [cmd=%vCmd% rc=%errorlevel%]"
		exit /b 2
	)
:proc-Main
	call :proc-RunLog "I" "command starting... [cmd=%~nx0]"
	setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

	set vCmd=get-id-data-last.cmd
	
	call %vCmdHome%%vCmd% %vDestPath% vIdLast_

	set /a vId=%vIdLast_%+1
	for /r "%~d1\" %%i in (*.*) do (
		echo !vId!,"%%~nxi","%vVolName_%","%vVolSer_%","%%~dpi","%%~ai","%%~ti","%%~zi","%%~xi">>"%vDestPath%"
		set /a vId+=1
	)

	call :proc-RunLog "I" "command ended [cmd=%~nx0]"
:proc-Term
	endlocal
	
	pause
	goto :eof

@rem //*****************************************************
@rem //* @procedure :proc-RunLog
@rem //*****************************************************
:proc-RunLog
		set vDateTime=%date: =0% %time: =0%
		echo [%~1],%vDateTime%,%~2

		goto :eof

