@echo off
@rem //****************************************************
@rem //* @procedure proc-Begin
@rem //****************************************************
:proc-Begin
	set vRc=
	
	@rem // is null argument 
	if "%~1" == "" (
		call :proc-RunLog "E" "Usage:%~nx0 [id-data-file] [var-id-data]"
		set vRc=1
		goto :proc-End
	)

	@rem is exist file
	if not exist "%~1" (
		call :proc-RunLog "E" "file not exist [file=%~1]"
		set vRc=2
		goto :proc-End
	)

	set vDataFile=%~1
	set vVarIdData=%~2
@rem //****************************************************
@rem //* @procedure proc-Main
@rem //****************************************************
:proc-Main
	set vDataId_=""

	for /f "tokens=1,2 delims==" %%i in (%vDataFile%) do (
		if "%%i" == "id" (
			set vDataId_=%%j
		)
	)
	
	if %vDataId_% == "" (
		set vRc=3
		call :proc-RunLog "E" "get [data-id] failure!! [file=%vDataFile%]"
		goto :proc-End
	)

	@rem id data increment
	set /a vDataId_+=1
	
	@rem set next id data to variable-name 
	set %vVarIdData%=%vDataId_%

@rem //****************************************************
@rem //* @procedure proc-End
@rem //****************************************************
:proc-End
	@rem --pause
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

