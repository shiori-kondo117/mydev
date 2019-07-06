@rem //*****************************************************
@rem //* @procedure proc-Begin
@rem //*****************************************************
:proc-Begin
	@rem -- set configration
	set vAppHome=C:\Program Files (x86)\CPRMDecrypter
	set vApp=cprmgetkey.exe
	set vTmpFile=%tmp%\%~n0_tmp.log
	set vKeyCprmkey="ContentsKey Base64:"
	
	@rem -- arguments validation
	if "%~1" == "" (
		echo Usage:%~nx0 [drive-name] [var-cprm-key]
		exit /b 1
	)

	if "%~2" == "" (
		echo Usage:%~nx0 [drive-name] [var-cprm-key]
		exit /b 2
	)

	@rem -- set src-drive-name
	set vDriveName=%~d1

@rem //*****************************************************
@rem //* @procedure proc-Begin
@rem //*****************************************************
:proc-Main
	"%vAppHome%\%vApp%" %vDriveName% >2 >%vTmpFile%

	type %vTmpFile%|findstr %vKeyCprmkey% 2>&1 > nul

	if %errorlevel% neq 0 (
		echo app execute failer!! [app=%vApp% rc=%errorlevel%]
		exit /b 3
	)
	for /f "tokens=1,2 delims=:" %%i in (%vTmpFile%) do (
		if "%%i" == "ContentsKey Base64" (
			set vCprmKey=%%j
		)
	)

	if %errorlevel% neq 0 (
		echo [E],%date: =0% %time: =0%,command failer!! [cmd=%vCmd% rc=%errorlevel%]
	)

	@rem -- set cprm-key and trim  blank
	set %2=%vCprmKey: =%

@rem //*****************************************************
@rem //* @procedure proc-End
@rem //*****************************************************
:proc-End
	goto :eof
