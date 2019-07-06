Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001
Const DC_CSVFILE_PATH = "D:\temp\"
Const ForReading = 1, ForWriting = 2, ForAppending = 8

On Error Resume Next

Private Sub cmdResisterData()
	Dim strHTML

	strHTML = "<button onclick=""cmdTest_Click"">テスト</button>" & vbNewLine
'//	Msgbox("正常に登録されました。")
	ID_FRM_CONTENT.document.body.innerHTML = strHTML
End Sub
Private Sub cmdTest_Click()
	Msgbox("テストボタンが押されました。")
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

	strSQL = "SELECT * FROM cprmkey.csv "

	DEBUG_VIEW.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

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

	ID_FRM_CONTENT.document.body.innerHTML = strHTML
'//	RESULT.innerHTML = strHTML

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
	Call window.resizeTo(640, 480)
End Sub

