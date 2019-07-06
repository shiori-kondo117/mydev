:proc-Init
	if "%~1" == "" (
		call :proc-Usage %~nx0
		exit /b 1
	)
	
	if "%~2" == "" (
		call :proc-Usage %~nx0
		exit /b 2
	)
	
	set vIdFilePath=%~1

:proc-Main	
	for /f "tokens=1 delims=," %%i in (%vIdFilePath%) do (
		set vIdLast=%%i
	)
	
	if %errorlevel% neq 0 (
		echo "E" "csv file read failer!! [file=%vDestPath%]"
		exit /b 3
	)

:proc-Term
	set %~2=%vIdLast%

	goto :eof

:proc-Usage
		echo Usage:%~1 [id-file-path] [var-id-last]

		goto :eof
