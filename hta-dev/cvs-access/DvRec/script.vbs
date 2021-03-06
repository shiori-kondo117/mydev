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

	strSQL = "SELECT * FROM DvRec.csv "

	DEBUG_VIEW.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	objRecordSet.AddNew
	objRecordSet("ＤＶ番号") = F_DV_NO.value
	objRecordSet("取込開始位置") = F_IMP_START_TIME.value
	objRecordSet("取込終了位置") = F_IMP_END_TIME.value
	objRecordSet("録画日付") = F_REC_DATE.value
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

	strSQL = "SELECT * FROM DvRec.csv order by 1"

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
	F_DV_NO.value = ""
	F_IMP_START_TIME.value = ""
	F_IMP_END_TIME.value = ""
	F_REC_DATE.value = ""
End Sub

Private Sub body_Load

End Sub

Private Sub Window_onload()
	Call window.resizeTo(640, 480)
End Sub

