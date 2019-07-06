Option Explicit

Private Function getDriveAvalilableSpace( _
	byval p_strDriveLetter _
	)
	'-- var Object
	Dim objFSO
	Dim objDrive
	
	'-- var Integer
	Dim intAvailableSpace
	
	'-- var String
	
	intAvailableSpace  = 0
	
	getDriveAvalilableSpace = 0
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Set objDrive = objFSO.GetDrive(objFSO.GetDriveName(p_strDriveLetter))
	
	intAvailableSpace = objDrive.AvailableSpace
	
	getDriveAvalilableSpace = intAvailableSpace
	
	Set objFSO = Nothing
End Function

