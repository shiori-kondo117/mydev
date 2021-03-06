'//**********************************************************
'// Version	: 0.3
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
		"WHERE フォルダ番号 = '" & F_FOLDER_NAME.value & "'" & vbCrLf

	DEBUG_VIEW.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	If (objRecordset.RecordCount > 0) Then
		Msgbox "フォルダ番号 " & F_FOLDER_NAME.value & _
			" は、既に登録済みです。(cnt=" & objRecordset.RecordCount & ")"
		Exit Sub
	End If

	objRecordSet.AddNew
	objRecordSet("フォルダ番号") = F_FOLDER_NAME.value
	objRecordSet("CPRMキー") = F_CPRMKEY.value
	objRecordSet.Update

	Msgbox("正常に登録されました。")
'//	ID_FRM_CONTENT.document.body.innerHTML = "正常に登録されました。"
'//	RESULT.innerHTML = "データが登録されました。"

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
		"WHERE フォルダ番号 = '" & F_FOLDER_NAME.value & "'"

	DEBUG_VIEW.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	If (objRecordset.EOF) Then
		Msgbox "データは存在しません。"
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

Private Sub body_Load

End Sub

Private Sub Window_onload()
	Call window.resizeTo(320, 240)
End Sub

