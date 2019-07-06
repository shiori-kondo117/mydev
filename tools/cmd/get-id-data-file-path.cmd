@echo off
@rem //****************************************************
@rem //* @procedure proc-Begin
@rem //****************************************************
:proc-Begin
	set vRc=

	@rem // import profile
	call %~dp0profile.cmd

	if "%~1" == "" (
		echo Usage:%~nx0 [media-type] [var-file-path]
		set vRc=1
		goto :proc-End
	)

	if "%~2" == "" (
		echo Usage:%~nx0 [media-type] [var-file-path]
		set vRc=2
		goto :proc-End
	)

	@rem -- set arguments
	set vMediaType=%~1
	set vVarDataFilePath=%~2

@rem //****************************************************
@rem //* @procedure proc-Main
@rem //****************************************************
:proc-Main
	set vDataFile_=""

@rem --	set vDataHome=C:\Users\skybl\Documents\mydata\db\csv
	set vDataFile=%vDataHome%\db\csv\id_data_files.csv

	for /f "skip=1 tokens=2,3 delims=," %%i in (%vDataFile%) do (
		if /i "%%i" == "%~1" (
			set vDataFile_=%%j
		)
	)
	
	if %vDataFile_% == "" (
		set vRc=2
		call :proc-RunLog "E" "get [data-id-file] is search not found!! [medai-type=%vMediaType%]"
		goto :proc-End
	)

@rem --	set vDataHome=C:\Users\skybl\Documents\mydata\db\text
	set %vVarDataFilePath%=%vDataHome%\db\text\%vDataFile_%
	
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

