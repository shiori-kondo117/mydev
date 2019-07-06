@rem --@echo off
if "%~1" == "" (
	echo "Usage:%~nx0 [src-dir] [dest-dir]"
	exit /b 1
)

set /p vDestDir="dest-dir:"
@rem --if "%~2" == "" (
@rem --	echo "Usage:%~nx0 [src-dir] [dest-dir]"
@rem --	exit /b 2
@rem --)
@rem --
if not exist "%~1" (
	echo "directory not exist [src-dir=%1]"
	exit /b 3
)

if not exist "%vDestDir%" (
	echo "directory not exist [dest-dir=%vDestDir%]"
	exit /b 4
)

echo "%~a1"|findstr ^d

if not errorlevel 0 (
	echo "not directory [src-dir=%1]"
	exit /b 5
)

for /d %%x in ("%vDestDir%") do @echo %%~ax|findstr ^d

if not errorlevel 0 (
	echo "not directory [dest-dir=%vDestDir%]"
	exit /b 6
)

set vSrcDir="%~1"
@rem --set vDestDir="%~2"

@rem --set vCmd=%temp%\copy-rename-files_.cmd

@rem --call :Proc-Copy-Rename > %vCmd%
call :Proc-Copy-Rename "V"

@rem --type %vCmd%

set /p vYesNo="Run Copy?...[y/n] "

if not "%vYesNo%" == "y" (
	goto :Proc-End
)

@rem --call %vCmd%
call :Proc-Copy-Rename "R"

if not errorlevel 0 (
	echo "run copy command failer!! [cmd=%vCmd%]
	exit /7
)

:Proc-End
	del %vCmd%
	pause
	goto :eof

:Proc-Copy-Rename
	for /r %vSrcDir% %%x in (*.*) do (
		for /f "tokens=1,2,3 delims=_" %%a in ("%%~nxx") do (
			set vCmdLine=copy "%%~dpnxx" %vDestDir%\"%%b_%%a_%%c"
			echo %vCmdLine%
			if "%~1" == "R" %vCmdLine%
		)
	)

	goto :eof
