'//**********************************************************
'// Version	: 0.3.1
'//**********************************************************
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001
Const DC_CSVFILE_PATH = "D:\home\kazu\MyData\Database\"
Const ForReading = 1, ForWriting = 2, ForAppending = 8

On Error Resume Next

Private Sub cmdResisterData()
	Dim strFileType
	Dim strWQ

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strPathtoTextFile = DC_CSVFILE_PATH

	objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	          "Data Source=" & strPathtoTextFile & ";" & _
	          "Extended Properties=""text;HDR=YES;FMT=Delimited"""

	strSQL = "SELECT * FROM cprmkey.csv " & _
		"WHERE �t�H���_�ԍ� = '" & F_FOLDER_NAME.value & "'" & vbCrLf

	DEBUG_VIEW.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	If (objRecordset.RecordCount > 0) Then
		Msgbox "�t�H���_�ԍ� " & F_FOLDER_NAME.value & _
			" �́A���ɓo�^�ς݂ł��B(cnt=" & objRecordset.RecordCount & ")"
		Exit Sub
	End If

	objRecordSet.AddNew
	objRecordSet("�t�H���_�ԍ�") = F_FOLDER_NAME.value
	objRecordSet("CPRM�L�[") = F_CPRMKEY.value
	objRecordSet.Update

	Msgbox("����ɓo�^����܂����B")
'//	ID_FRM_CONTENT.document.body.innerHTML = "����ɓo�^����܂����B"
'//	RESULT.innerHTML = "�f�[�^���o�^����܂����B"

	objRecordSet.Close()
	objConnection.Close()

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Sub cmdListData()
	Dim strFileType
	Dim strWQ

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strPathtoTextFile = DC_CSVFILE_PATH

	objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	          "Data Source=" & strPathtoTextFile & ";" & _
	          "Extended Properties=""text;HDR=YES;FMT=Delimited"""

	strSQL = "SELECT * FROM cprmkey.csv " & _
		"WHERE �t�H���_�ԍ� = '" & F_FOLDER_NAME.value & "'"

	DEBUG_VIEW.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	If (objRecordset.EOF) Then
		Msgbox "�f�[�^�͑��݂��܂���B"
		Exit Sub
	End If

	objRecordset.MoveFirst

	strHTML = "<table border=""1"" cellspacing=""0"">" & vbCrLf
	strHTML = strHTML & "<tr style=""background-color:skyblue;"">"
	For Each objField in objRecordset.Fields
		strHTML = strHTML & "<th>" & objField.Name & "</th>" & vbNewLine
	Next
	strHTML = strHTML & "</tr>"

	Do Until (objRecordset.EOF)
		strHTML = strHTML & "<tr>"
		For Each objField in objRecordset.Fields
			strHTML = strHTML & "<td>" & objField.Value & "</td>" & vbNewLine
		Next
		strHTML = strHTML & "</tr>"

	    objRecordset.MoveNext
	Loop

	strHTML = strHTML & "</table>"

'//	ID_FRM_CONTENT.document.body.innerHTML = strHTML
	RESULT.innerHTML = strHTML

	objRecordSet.Close()
	objConnection.Close()

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub

Private Sub cmdClearData()
	F_FOLDER_NAME.value = ""
	F_CPRMKEY.value = ""
End Sub
'// CPRM-KEY���擾����B
Private Sub getCprmKey()
	Dim objWsh
	Dim objFSO
	Dim objTS, strTS
 	Dim oRe, oMatch, oMatches

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objWsh = CreateObject("WScript.Shell")

	Call objWsh.Run("%ComSpec% /C cprmgetkey.bat > cprmgetkey.log", , True)

	Set objTS = objFSO.OpenTextFile("cprmgetkey.log")

	strTS = objTS.ReadAll()

	objTS.close()

  	Set oRe = New RegExp
  	oRe.Pattern = "ContentsKey Base64: (.+)"

	' Matches �R���N�V�������擾���܂��B
  	Set oMatches = oRe.Execute(strTS)

	If (isObject(oMatches)) Then
		' Matches �R���N�V�����̍ŏ��̍��ڂ��擾���܂��B
	  	Set oMatch = oMatches(0)
		
		' �T�u�}�b�`�̕������擾���܂��B
		'//RESULT.value = strTS & vbCrLf
		F_CPRM_KEY.value = oMatch.SubMatches(0)
'//		DEBUG_VIEW.innerText = oMatch.SubMatches(0)
	Else
		Msgbox "�f�B�X�N�G���[���������܂����B"
		Exit Sub
	End If

	Set objTS = Nothing
	Set objFSO = Nothing
End Sub
Private Sub body_Load

End Sub

Private Sub Window_onload()
	Call window.resizeTo(320, 240)
End Sub

