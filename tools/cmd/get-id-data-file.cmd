	if %errorlevel% neq 0 (
		proc-RunLog "E" "command failer!! [cmd=%vCmd% rc=%errorlevel%]"
		exit /b 3
	)

	
@rem //*****************************************************
@rem //* @procedure :proc-RunLog
@rem //*****************************************************
:proc-RunLog
		set vDateTime=%date: =0% %time: =0%
		echo [%~1],%vDateTime%,%~2

		goto :eof

