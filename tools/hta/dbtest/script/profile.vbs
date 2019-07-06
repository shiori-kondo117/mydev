Option Explicit

'//*********************************************************
'// global @constant definition
'//*********************************************************
Const FUNC_EVENT_PREFIX = "evtActionCommand_"

'//*********************************************************
'// global @variable definition
'//*********************************************************
'-- var String
Dim g_strHomeDir

'-- var Object
Dim g_objFso

'//*********************************************************
'//* @procedure window_onload
'//*********************************************************
Private Sub window_onload()
	'-- var String
	Dim strCmdLine
	
	window.resizeTo 1024, 768
	
	Set g_objFso = CreateObject("Scripting.FileSystemObject")
	
	
	strCmdLine = Replace(oHTA.commandLine,"""","")
	g_strHomeDir = g_objFso.GetParentFolderName(strCmdLine) & "\..\"
	
'--	msgbox g_strHomeDir
End Sub
