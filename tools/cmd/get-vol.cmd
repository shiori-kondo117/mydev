@echo off
:proc-Init
	if "%~1" == "" (
		call :proc-Usage %~nx0
		exit /b 1
	)

	if "%~2" == "" (
		call :proc-Usage %~nx0
		exit /b 2
	)

	if "%~3" == "" (
		call :proc-Usage %~nx0
		exit /b 3
	)

	if "%~4" == "" (
		call :proc-Usage %~nx0
		exit /b 4
	)

:proc-Main
	set vCmd=vol %~d1
	for /f "usebackq tokens=1,2,3,5" %%i in (`%vCmd%`) do (
		if "%%~i" == "ドライブ" (
			set vDriveName=%%j
			set vVolName=%%l
		) else if "%%~i" == "ボリューム" (
			set vVolSer=%%k
		)
	)

	if %errorlevel% neq 0 (
		call :proc-RunLog "E" "command failer!! [cmd=%vCmd% rc=%errorlevel%]"
		exit /b 5
	)

:proc-Term
	set %~2=%vDriveName%
	set %~3=%vVolName%
	set %~4=%vVolSer%

	goto :eof

@rem //*****************************************************
@rem //* @procedure :proc-Usage
@rem //*****************************************************
:proc-Usage
	echo Usage:%~1 [drive-name] [var-driev-name] [var-vol-name] [var-vol-ser]
	
	goto :eof

@rem //*****************************************************
@rem //* @procedure :proc-RunLog
@rem //*****************************************************
:proc-RunLog
		set vDateTime=%date: =0% %time: =0%
		echo [%~1],%vDateTime%,%~2

		goto :eof

