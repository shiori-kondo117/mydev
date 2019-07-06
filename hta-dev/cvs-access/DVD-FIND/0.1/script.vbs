Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001
Const DC_CSVFILE_PATH = "D:\MyDatabase\FileInfo"

On Error Resume Next

Private Sub cmdSearch_Click()

	'// �L�[���ڑI���`�F�b�N
	If (KEY_ITEM.value = "-") Then
		Msgbox("�L�[���ڂ�I�����ĉ������B")
		Exit Sub
	End If

	'// ���ލ��ڑI���`�F�b�N
	If (ORDER_BY.value = "-") Then
		'// �f�t�H���g���ڂ�ݒ肷��B
		ORDER_BY.value = "�t�@�C����"
	End If

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strPathtoTextFile = DC_CSVFILE_PATH

'//	Call OpenDbAccessFile(objConnection, DC_CSVFILE_PATH & "\" & "MyDatabase.mdb")
	Call OpenDbTextFile(objConnection, DC_CSVFILE_PATH)

	'//-- �L�[�����`�F�b�N
	strConditionType = KEY_CONDITION.value

	strWhere = _
		KEY_ITEM.value & " " & _
		setKeyCondition( _
			strConditionType, _
			KEYWORD.value)

	strSQL = "SELECT * FROM DVD_info_db.txt "
	If (strWhere <> "") Then
		strSQL = strSQL & _
			"where " & _
			strWhere
	End If

	strSQL = strSQL & _
			"order by " & ORDER_BY.value
	
	Call DEBUGLOG(strSQL)

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	strHTML="<table border='1' cellspacing='0' bgcolor='white'>" & vbCrLf & _
		"<tr style='background-color:skyblue;'>" & _
		"<th nowrap='true'>�{�����[����</th>" & _
		"<th>�t�@�C����</th>" & _
		"<th>�p�X</th>" & _
		"<th>���t</th>" & _
		"</tr>"
	If (objRecordset.RecordCount > 999) Then
		'//Msgbox "���R�[�h������1000���𒴂��Ă��܂��B" & _
		'//"rec-cnt=" & objRecordset.RecordCount
	End If

	RecCnt=0

	Do Until (objRecordset.EOF)
		strField_1 = objRecordset.Fields.Item("�{�����[����")
		strField_2 = objRecordset.Fields.Item("�t�@�C����")
		strField_3 = objRecordset.Fields.Item("�p�X")
		strField_4 = objRecordset.Fields.Item("���t")
		strHTML = strHTML & _
			"<tr>" & vbCrLf

	    strHTML = strHTML & _
			"<td nowrap>" & strField_1 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_2 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_3 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_4 & "<br/>" & "</td>"

		strHTML = strHTML & _
			"</tr>" & vbCrLf

	    objRecordset.MoveNext
		RecCnt = RecCnt + 1
	Loop
	RESULT.innerHTML = strHTML

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub

Private Function setKeyCondition( _
		byVal p_strCondType, _
		byVal p_strKeyword _
		)
'//	strKeyword = UCase(p_strKeyword)
	strKeyword = p_strKeyword
	Select Case p_strCondType
		case "1"	'// �L�[�Ɉ�v����
			strKeyCondition = " = '" & _
				strKeyword & "' "
		case "2"	'// �L�[�Ɉ�v���Ȃ�
			strKeyCondition = " not = '" & _
				strKeyword & "' "
		case "3"	'// �L�[�Ŏn�܂�
			strKeyCondition = " like '" & _
				strKeyword & "%' "
		case "4"	'// �L�[�ŏI���
			strKeyCondition = " like '%" & _
				strKeyword & "' "
		case "5"	'// �L�[���܂�
			strKeyCondition = " like '%" & _
				strKeyword & "%' "
		case else	'// �L�[���܂܂Ȃ�
			strKeyCondition = " not like '%" & _
				strKeyword & "%' "
	End Select

	setKeyCondition = strKeyCondition

End Function

Private Sub cmdClear_Click
	KEYWORD.value = ""
End Sub
Private Sub body_Load
	Call cmdSetup()
End Sub
Private Function MakeItems(byref p_objControl, byval p_strTitle)
	Dim objField
	Dim objOption

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	Call OpenDbTextFile(objConnection, DC_CSVFILE_PATH)

	strSQL = "SELECT * FROM DVD_info_db.txt " & _
				"WHERE �{�����[���� = 'DVD-0601'"
	Call DEBUGLOG(strSQL)

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	objRecordset.MoveFirst

	'// �_�E�����X�g�̃^�C�g���f�[�^��ݒ肷��B
	Set objOption = Document.createElement("OPTION")
	objOption.Text = p_strTitle
    objOption.Value = "-"
    p_objControl.Add(objOption)

	For Each objField in objRecordset.Fields
		Set objOption = Document.createElement("OPTION")
		objOption.Text = objField.Name
        objOption.Value = objField.Name
        p_objControl.Add(objOption)
	Next

	Set objOption = Nothing
	Set objField = Nothing
	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Function

Private Sub cmdSetup()
	Call MakeItems(ID_KEY_ITEM, "--�L�[����--")
	Call MakeItems(ID_ORDER_BY, "--���ލ���--")
End Sub

Private Sub window_onload()
	Call window.resizeTo(800, 600)
End Sub

Private Sub OpenDbTextFile(ByRef p_objConnection, ByVal p_TextFilePath)
	p_objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	          "Data Source=" & p_TextFilePath & ";" & _
	          "Extended Properties=""text;HDR=YES;FMT=Delimited"""
End Sub

Private Sub OpenDbAccessFile(ByRef p_objConnection, ByVal p_AccessFile)
	p_objConnection.Open _
	    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
	        "Data Source = " & p_AccessFile
End Sub

Private Sub DEBUGLOG(byval p_strText)
	Dim strMessage

	If (DEBUG_VIEW.innerText = "") Then
		strMessage = p_strText
	Else
		strMessage = _
			DEBUG_VIEW.innerText & vbNewLine & _
			p_strText & vbNewLine
	End If

	DEBUG_VIEW.innerText = strMessage			
End Sub

Private Sub ClearDebugView()
	DEBUG_VIEW.innerText = ""
End Sub
