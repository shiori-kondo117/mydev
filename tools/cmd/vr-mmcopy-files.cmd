if "%~1" == "" (
	echo "Usage:%~nx0 [src-drive] [dest-drive]"
	exit /b 1
)

if not exist "%~1" (
	echo "src-drive not exist!! [src-drive=%~1]
	exit /b 2
)

set /p vDestDrive="dest-drive[K:]:"

if not exist "%vDestDrive%" (
	echo "dest-drive not exist!! [src-drive=%vDestDrive%]
	exit /b 3
)

set /p vFileExt="file-ext[(MP4|DIVX|WMV|AVI|RMVB|MPG)]:"

echo %vFileExt%^|findstr "MP4 DIVX WMV AVI RMVB MPG" 1>null

if errorlevel 1 (
	echo "illigal file extention!! [file-ext=%vFileExt%]
	exit /b 4
)

set vDestDir=%vDestDrive%\%vFileExt%-FILES

if not exist "%vDestDir%" (
	mkdir %vDestDir%
)

set vSrcDir="%~d1\"


for /r %vSrcDir% %%x in (*.%vFileExt%) do (
	set vFile=%%~nxx
	if /i "%vFile:~0,1%" == "V" (
		for /f "tokens=1,2,3 delims=_." %%a in ("%%~nxx") do (
			if "%%~c" == "" (
				@echo "copy "%%~dpnxx" "%vDestDir%\%%b_%%a%%~xx"
			) else (
				@echo "copy "%%~dpnxx" "%vDestDir%\%%b_%%a%_%%c%~xx"
			)
		)
	) else (
		@echo "copy "%%~dpnxx" "%vDestDir%\"
	)
)

