@echo off
@rem //****************************************************
@rem //* @procedure proc-Begin
@rem //****************************************************
:proc-Begin
	@rem -- import profile
	call %~dp0profile.cmd
	set vRc=
	
	if "%~1" == "" (
		echo Usage:%~nx0 [hd-vol] [var-file-path]
		set vRc=1
		goto :proc-End
	)

	if "%~2" == "" (
		echo Usage:%~nx0 [hd-vol] [var-file-path]
		set vRc=2
		goto :proc-End
	)

	@rem -- set arguments
	set vHdVol=%~1
	set vVarDataFilePath=%~2

@rem //****************************************************
@rem //* @procedure proc-Main
@rem //****************************************************
:proc-Main
	set vDataFile_=""

@rem --	set vDataHome=J:\mydata
	set vDataFile=%vCsvHome%id_data_hd_files.csv

	set vCmd=findstr /i %vHdVol% %vDataFile%

	for /f "usebackq tokens=1,2 delims=," %%i in (`%vCmd%`) do (
		set vDataFile_=%%j
	)
	
	if %vDataFile_% == "" (
		set vRc=3
		call :proc-RunLog "E" "search  [data-id-hd-file] not found!! [hd-vol=%vHdVol%]"
		goto :proc-End
	)

@rem --	set vDataHome=L:\mydata\db\text
	set %vVarDataFilePath%=%vTextHome%%vDataFile_%
	
@rem //****************************************************
@rem //* @procedure proc-End
@rem //****************************************************
:proc-End
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

