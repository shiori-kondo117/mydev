@echo off
set vRc=0
set vVrInfoFile=VR_MANGR.IFO

if "%~1" == "" (
	echo Usage:%~nx0 [src-dir]
	set vRc=1
	goto :proc-End
)

if not "%~nx1" == "%vVrInfoFile%" (
	echo not exist VR_MANGR.IFO!! [src-file=%~dpnx1]
	set vRc=2
	goto :proc-End
)

set vHomeDir=%~dp1
set vOldFile=%vHomeDir%%vVrInfoFile%
set vNewFile=_%vVrInfoFile%
set vVrInfoBackupFile=%vHomeDir%VR_MANGR.BUP

ren %vOldFile% %vNewFile%
copy %vVrInfoBackupFile% %vHomeDir%%vVrInfoFile%

:proc-End
	pause
	if defined vRc (
		exit /b %vRc
	)
