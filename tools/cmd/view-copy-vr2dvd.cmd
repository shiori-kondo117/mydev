<<<<<<< HEAD
@echo off
:Begin-Proc
	setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

	set myRc=
	set myMsg=

	if "%~1" == "" (
		set myMsg="Usage:%~nx0 [src-dir]"
		set myRc=1
		goto End-Proc
	)

	if not exist "%~1" (
		set myMsg="src-dir not exist !! [src-dir=%~1]"
		set myRc=2
		goto End-Proc
	)

	set /p myDestDir="Copy To Destination Directory[S:\MOVIE-FILES]:"

	if not exist "%myDestDir%" (
		set myMsg="dest-dir not exist !! [dest-dir=%myDestDir%]"
		set myRc=3
		goto End-Proc
	)

	set mySrcDir=%~1

:Main-Proc
	for %%x in ("%mySrcDir%\*.*") do (
		if not exist "s:\MOVIE-FILES\%%~nxx" (
			echo copy "%%~dpnxx" "%myDestDir%"
		)
	)

:End-Proc
	if defined myMsg (
		echo %myMsg:"=%
	)
	pause
	if defined myRc (
		echo %myMsg:"=%
		exit /b %myRc%
	)
=======
@echo off
:Begin-Proc
	setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

	set myRc=
	set myMsg=

	if "%~1" == "" (
		set myMsg="Usage:%~nx0 [src-dir]"
		set myRc=1
		goto End-Proc
	)

	if not exist "%~1" (
		set myMsg="src-dir not exist !! [src-dir=%~1]"
		set myRc=2
		goto End-Proc
	)

	set /p myDestDir="Copy To Destination Directory[S:\MOVIE-FILES]:"

	if not exist "%myDestDir%" (
		set myMsg="dest-dir not exist !! [dest-dir=%myDestDir%]"
		set myRc=3
		goto End-Proc
	)

	set mySrcDir=%~1

:Main-Proc
	for %%x in ("%mySrcDir%\*.*") do (
		if not exist "s:\MOVIE-FILES\%%~nxx" (
			echo copy "%%~dpnxx" "%myDestDir%"
		)
	)

:End-Proc
	if defined myMsg (
		echo %myMsg:"=%
	)
	pause
	if defined myRc (
		echo %myMsg:"=%
		exit /b %myRc%
	)
>>>>>>> a70cdd52703c729acb7c415b8be99cf84b44e2f7
