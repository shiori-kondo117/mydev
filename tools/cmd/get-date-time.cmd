@rem -- Initial Process Start
:proc-Init
	if "%~1" == "" (
		call :proc-Usage
		exit /b 1
	)

	if "%~2" == "" (
		call :proc-Usage
		exit /b 2
	)
@rem -- Main Process Start
:proc-Main
	set vDate=%date: =0%
	set vTime=%time: =0%

	set %~1=%vDate%
	set %~2=%vTime%
	
@rem -- Termination Process Start
:proc-Term
	goto :eof

@rem //*****************************************************
@rem //* @procedure :proc-Usage
@rem //*****************************************************
:proc-Usage
		echo Usage:%~nx0 [var-date] [var-time]

		goto :eof

