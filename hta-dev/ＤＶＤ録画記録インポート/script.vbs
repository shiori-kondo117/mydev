<!-- version:0.1.0.1
'//---------------------------------------------------------
'// �O���[�o���ϐ� ��`
'//---------------------------------------------------------
Public g_objDebugView			'// �f�o�b�O�\���I�u�W�F�N�g��
'//---------------------------------------------------------
'// �O���[�o���萔 ��`
'//---------------------------------------------------------
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001

Private Sub cmdSearch_Click(byVal intRecLogNo)
	Dim strHTML
	Dim strWQ

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT * FROM L_�^��L�^ WHERE �^��L�^�ԍ� = " & intRecLogNo

	DEBUGLOG strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	'// �^����t�̃t�B�[���h�𐶐�����B
	strHTML = "<input type=" & strWQ & "text" & strWQ & " " & _
		"name=" & strWQ & "REC_DATE" & strWQ & " " & _
		"size=" & strWQ & "13" & strWQ & "/>"
	ID_REC_DATE.innerHTML = strHTML

	'// �^�掞�Ԃ̃t�B�[���h�𐶐�����B
	strHTML = "<input type=" & strWQ & "text" & strWQ & " " & _
		"name=" & strWQ & "REC_TIME" & strWQ & " " & _
		"size=" & strWQ & "9" & strWQ & "/>"
	ID_REC_TIME.innerHTML = strHTML

	Do Until (objRecordset.EOF)
		ID_REC_LOG_NO.innerText = objRecordset.Fields.Item("�^��L�^�ԍ�")
		DVD_NO.value = CInt(Right(objRecordset.Fields.Item("�f�B�X�N�h�c"),4))
		If (objRecordset.Fields.Item("�{�����[���ԍ�") <> "") Then
			VOLUME_NO.value = objRecordset.Fields.Item("�{�����[���ԍ�")
		Else
			VOLUME_NO.value = ""
		End If
		BANGUMI_NAME.value = objRecordset.Fields.Item("�ԑg��")
		TITLE_NAME.value = objRecordset.Fields.Item("�^�C�g����")
		REC_DATE.value = objRecordset.Fields.Item("�^����t")
		REC_TIME.value = objRecordset.Fields.Item("�^�掞��")
		REC_MODE.value = objRecordset.Fields.Item("�^�惂�[�h")
		COPYONCE.checked = objRecordset.Fields.Item("�R�s�[�����X")
		DELETE_FLAG.checked = objRecordset.Fields.Item("�폜�t���O")
		F_EDIT_FLAG.checked = objRecordset.Fields.Item("�ҏW�t���O")
		F_DELETE_DATE.value = objRecordset.Fields.Item("�폜���t")
		F_MEMO.value = objRecordset.Fields.Item("����")
	    objRecordset.MoveNext
		RecCnt = RecCnt + 1
	Loop
	'//DEBUGLOG strHTML

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Sub cmdUpdate_Click(byVal intRecLogNo)
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3
	Dim strHTML
	Dim strWQ

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	'// �g�����U�N�V�������J�n����B
	objConnection.BeginTrans()

'	strSQL = "UPDATE T_�^��L�^ " & _
'			"set �f�B�X�N�h�c = " & _
'				strWQ & "DVD-" & string(len(DVD_NO.value), "0") & DVD_NO.value & strWQ & "," & vbCrLf & _
'			"�^�C�g���� = " & strWQ & TITLE_NAME.value & strWQ & "," & vbCrLf & _
'			"�^����t = " & strWQ & REC_DATE.value & strWQ & "," & vbCrLf & _
'			"�^�掞�� = " & strWQ & REC_TIME.value & strWQ & "," & vbCrLf & _
'			"�{�����[���ԍ� = " & VOLUME_NO.value & " " & vbCrLf & _
'			"where �^��L�^�ԍ� = " & intRecLogNo

	strSQL = "SELECT * FROM T_�^��L�^ WHERE �^��L�^�ԍ� = " & intRecLogNo
	DEBUGLOG strSQL

	objRecordset.CursorLocation = adUseClient
	objRecordset.Open strSQL , objConnection, _
	    adOpenStatic, adLockOptimistic
'	Set objRecordset = objConnection.Execute(strSQL)

	objRecordset.MoveFirst

	objRecordset("�f�B�X�N�h�c") = makeDiscID(DVD_NO.value)
	objRecordset("�ԑg��") = BANGUMI_NAME.value
	objRecordset("�^�C�g����") = TITLE_NAME.value
	objRecordset("�{�����[���ԍ�") = VOLUME_NO.value
	objRecordset("�^����t") = REC_DATE.value
	objRecordset("�^�掞��") = FormatDateTime(REC_TIME.value, vbLongTime) 

	objRecordset("�^�惂�[�h") = REC_MODE.value
	'//objRecordset("�^���") = F_REC_COUNT.value
	objRecordset("�R�s�[�����X") = COPYONCE.checked
	objRecordset("�폜�t���O") = DELETE_FLAG.checked
	objRecordset("�폜���t") = F_DELETE_DATE.value
	objRecordset("�ҏW�t���O") = F_EDIT_FLAG.checked
	objRecordset("����") = F_MEMO.value
	objRecordset.Update

	'// �g�����U�N�V�������R�~�b�g����B
	objConnection.CommitTrans()

	Msgbox "�X�V���������܂����B"

	objRecordset.Close
	objConnection.Close

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Sub cmdDelete_Click(byVal intRecLogNo)
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3
	Dim strHTML
	Dim strWQ

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	'// �g�����U�N�V�������J�n����B
	objConnection.BeginTrans()

'	strSQL = "DELETE FROM T_�^��L�^ " & _
'			"WHERE �^��L�^�ԍ� = " & intRecLogNo

	strSQL = "SELECT * FROM T_�^��L�^ WHERE �^��L�^�ԍ� = " & intRecLogNo
	DEBUGLOG strSQL

	objRecordset.CursorLocation = adUseClient
	objRecordset.Open strSQL , objConnection, _
	    adOpenStatic, adLockOptimistic

	objRecordset.Delete

	'// �g�����U�N�V�������R�~�b�g����B
	objConnection.CommitTrans()

	Msgbox "�폜���������܂����B"

	objRecordset.Close
	objConnection.Close

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Sub insertData()
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")
	Set objRecordSet2 = CreateObject("ADODB.Recordset")

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	'// �g�����U�N�V�������J�n����B
	objConnection.BeginTrans()

	strSQL = F_SQLCOMMAND2.value

	set objRecordSet = objConnection.Execute(strSQL)

	'// �g�����U�N�V�������R�~�b�g����B
	objConnection.CommitTrans()

	msgbox "����ɓo�^����܂����B"

	objConnection.Close

	set objRecordSet = nothing
	set objConnection = nothing

End Sub
Private Sub cmdRegister2_Click
Const adOpenStatic = 3
Const adLockOptimistic = 3
	Dim strDvdNo
	Dim strRecDate
	Dim strRecTime
	Dim intRecCount

	Err.Clear

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	'// �g�����U�N�V�������J�n����B
	objConnection.BeginTrans()

	'// �^��L�^�e�[�u������S�f�[�^����������B
	strSQL = "select * from T_�^��L�^"

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText
	'// �c�u�c�ԍ���ҏW����B
	strDvdNo = makeDiscID(DVD_NO.value)

	'// �^����t��ҏW����B
	strRecDate = REC_DATE_YY.value & "/" & _
				REC_DATE_MM.value & "/" & _
				REC_DATE_DD.value

	'// �^�掞�Ԃ�ҏW����B
	strRecTime = REC_TIME_HH.value & ":" & _
				REC_TIME_MM_H.value & _
				REC_TIME_MM_L.value & ":" & _
				REC_TIME_SS_H.value & _
				REC_TIME_SS_L.value

	objRecordSet.AddNew
	objRecordSet("�f�B�X�N�h�c") = strDvdNo
	objRecordSet("�{�����[���ԍ�") = VOLUME_NO.value
	objRecordSet("�ԑg��") = BANGUMI_NAME.value
	objRecordSet("�^�C�g����") = TITLE_NAME.value
	objRecordSet("�^����t") = strRecDate
	objRecordSet("�^�掞��") = strRecTime
	objRecordSet("�^�惂�[�h") = REC_MODE.value
	objRecordSet("�R�s�[�����X") = COPYONCE.checked
	objRecordSet("�폜�t���O") = DELETE_FLAG.checked
	objRecordSet("�폜���t") = 0
	objRecordSet("�ҏW�t���O") = F_EDIT_FLAG.checked
	objRecordSet("����") = F_DEST_DISCNO.value & vbCrLf & F_MEMO.value
	objRecordSet.Update

	If (Err) Then
		Msgbox Err.Description
		objConnection.RollbackTrans()
		objRecordSet.Close
		objConnection.Close
		set objRecordSet = nothing
		set objConnection = nothing
		Exit Sub
	End If

	'// �g�����U�N�V�������R�~�b�g����B
	objConnection.CommitTrans()

	msgbox "����ɓo�^����܂����B"

	objRecordSet.Close

	objConnection.Close

	set objRecordSet = nothing
	set objConnection = nothing
End Sub
Private Sub cmdSetup()
	'// �f�o�b�O�\���I�u�W�F�N�g��ݒ肷��B
	If (isObject(ID_DEBUG_VIEW)) Then
		Set g_objDebugView = ID_DEBUG_VIEW
	End If

	ID_REC_DATE.innerHTML = MakeSelectDate()
	ID_REC_TIME.innerHTML = MakeSelectTime()
	Call MakeItems(ID_KEY_ITEM, "--�L�[����--")
	Call MakeItems(ID_ORDERBY, "--���ލ���--")
	Call MakeListBox("ConditionList.txt", KEY_CONDITION)

End Sub
Private Function MakeSelectDate()
	Dim strRecDate
	Dim objDate
	Dim strYear, strMonth, strDay
	Dim strDefault
	Dim i

	objDate = Now()
	strYear = Year(objDate)
	strMonth = Month(objDate)
	strDay = Day(objDate)

	'// �N
	strRecDate = "<select name='REC_DATE_YY'>"
	For i=2000 To 2010
		strDefault = " "
		If (i = strYear) Then
			strDefault = "selected='true' "
		End If

		strRecDate = strRecDate & _
			"<option " & strDefault & "value='" & i & "'>" & i & "</option>" & vbCrLf
	Next
	strRecDate = strRecDate & "</select>�N"

	'// ��
	strRecDate = strRecDate & "<select name='REC_DATE_MM'>"
	For i=1 To 12
		strDefault = " "
		If (i = strMonth) Then
			strDefault = "selected='true' "
		End If

		strRecDate = strRecDate & _
			"<option " & strDefault & "value='" & i & "'>" & i & "</option>" & vbNewLine
	Next
	strRecDate = strRecDate & "</select>"

	'// ��
	strRecDate = strRecDate & "<select name='REC_DATE_DD'>"
	For i=1 To 31
		strDefault = " "
		If (i = strDay) Then
			strDefault = "selected='true' "
		End If

		strRecDate = strRecDate & _
			"<option " & strDefault & "value='" & i & "'>" & i & "</option>" & vbNewLine
	Next
	strRecDate = strRecDate & "</select>"

	MakeSelectDate = strRecDate
End Function
Private Function MakeSelectTime()
	Dim strRecTime
	Dim i

	'// ��
	strRecTime = "<select name='REC_TIME_HH'>"
	For i=0 To 24
		strRecTime = strRecTime & _
			"<option value='" & i & "'>" & i & "</option>" & vbCrLf
	Next
	strRecTime = strRecTime & "</select>��"

	'// ��(��)
	strRecTime = strRecTime & "<select name='REC_TIME_MM_H'>"
	For i=0 To 5
		strRecTime = strRecTime & _
			"<option value='" & i & "'>" & i & "</option>" & vbNewLine
	Next
	strRecTime = strRecTime & "</select>"

	'// ��(��)
	strRecTime = strRecTime & "<select name='REC_TIME_MM_L'>"
	For i=0 To 9
		strRecTime = strRecTime & _
			"<option value='" & i & "'>" & i & "</option>" & vbNewLine
	Next
	strRecTime = strRecTime & "</select>��"

	'// �b(��)
	strRecTime = strRecTime & "<select name='REC_TIME_SS_H'>"
	For i=0 To 5
		strRecTime = strRecTime & _
			"<option value='" & i & "'>" & i & "</option>" & vbNewLine
	Next
	strRecTime = strRecTime & "</select>"

	'// �b(��)
	strRecTime = strRecTime & "<select name='REC_TIME_SS_L'>"
	For i=0 To 9
		strRecTime = strRecTime & _
			"<option value='" & i & "'>" & i & "</option>" & vbNewLine
	Next
	strRecTime = strRecTime & "</select>�b"

	MakeSelectTime = strRecTime
End Function
Private Function MakeItems(byref p_objControl, byval p_strTitle)
	Dim objField
	Dim objOption

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT * FROM T_�^��L�^ where �^��L�^�ԍ� = 1"
	DEBUGLOG strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	objRecordset.MoveFirst

	'// �_�E�����X�g�̃^�C�g���f�[�^��ݒ肷��B
	Set objOption = Document.createElement("OPTION")
	objOption.Text = p_strTitle
    objOption.Value = "-"
    p_objControl.Add(objOption)

	Do Until (objRecordset.EOF)
		For Each objField in objRecordset.Fields
			Set objOption = Document.createElement("OPTION")
			objOption.Text = objField.Name
	        objOption.Value = objField.Name
	        p_objControl.Add(objOption)
		Next
	    objRecordset.MoveNext
	Loop

	Set objOption = Nothing
	Set objField = Nothing
	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Function
Private Sub window_onload()
	Call window.resizeTo(1024, 480)
End Sub
Private Sub cmdClear_Click()
	DVD_NO.value = ""
	VOLUME_NO.value = ""
	BANGUMI_NAME.value = ""
	TITLE_NAME.value = ""
	REC_MODE.value = "SP"
	COPYONCE.checked = "0"
	DELETE_FLAG.checked = "0"
	F_EDIT_FLAG.checked = "0"
	F_DELETE_DATE.value = ""
	F_MEMO.value = ""
End Sub
Private Sub cmdClear2_Click()
	KEY_ITEM.value = "-"
	KEYWORD.value = ""
	KEY_CONDITION.value = "-"
	F_ORDERBY.value = "-"
	RESULT_LIST.innerHTML = ""
End Sub

Private Sub ClearDebugView()
	ID_DEBUG_VIEW.innerText = ""
End Sub

Private Sub cmdSearch2_Click
	Dim strFileType
	Dim strWQ

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	'//-- �L�[�l��ݒ�
	Dim strKeyValue
	strKeyValue = KEYWORD.value
	If (KEY_ITEM.value = "�f�B�X�N�h�c") Then
		strKeyValue = makeDiscID(strKeyValue)
	End If

	'//-- �L�[�����`�F�b�N
	strConditionType = KEY_CONDITION.value

	'//-- �L�[�����쐬
	strWhere = ""
	If (KEY_ITEM.value <> "") Then
		strWhere = _
			KEY_ITEM.value & " " & _
			setKeyCondition( _
				strConditionType, _
				strKeyValue)
	End If

	strSQL = "SELECT * FROM T_�^��L�^ "
	If (strWhere <> "") Then
		strSQL = strSQL & _
				"where " & _
				strWhere
	End If

	strSQL = strSQL & _
			"order by " & F_ORDERBY.value

	DEBUGLOG strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	strHTML="<table border='1' cellspacing='0'>" & vbCrLf & _
		"<tr style='background-color:skyblue;'>" & _
		"<th>�f�B�X�N�h�c</th>" & _
		"<th>�{�����[���ԍ�</th>" & _
		"<th>�ԑg��</th>" & _
		"<th>�^�C�g����</th>" & _
		"<th>�^����t</th>" & _
		"<th>�^�掞��</th>" & _
		"<th>�폜�t���O</th>" & _
		"<th>�ҏW�t���O</th>" & _
		"<th>�A�N�V����</th>" & _
		"</tr>"
	If (objRecordset.RecordCount > 999) Then
		'//Msgbox "���R�[�h������1000���𒴂��Ă��܂��B" & _
		'//"rec-cnt=" & objRecordset.RecordCount
	End If

	RecCnt=0

	Do Until (objRecordset.EOF)
		strField_1 = objRecordset.Fields.Item("�f�B�X�N�h�c")
		strField_2 = objRecordset.Fields.Item("�^�C�g����")
		strField_3 = objRecordset.Fields.Item("�^����t")
		strField_4 = objRecordset.Fields.Item("�^�掞��")
		strField_5 = objRecordset.Fields.Item("�{�����[���ԍ�")
		strField_6 = objRecordset.Fields.Item("�^��L�^�ԍ�")
		strField_7 = objRecordset.Fields.Item("�폜�t���O")
		strField_8 = objRecordset.Fields.Item("�ҏW�t���O")
		strField_9 = objRecordset.Fields.Item("�ԑg��")
		strHTML = strHTML & _
			"<tr>" & vbCrLf

	    strHTML = strHTML & _
			"<td nowrap>" & strField_1 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_5 & "<br/>" & "</td>" & _
			"<td>" & strField_9 & "<br/>" & "</td>" & _
			"<td>" & strField_2 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_3 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_4 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_7 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_8 & "<br/>" & "</td>" & _
			"<td nowrap><button onclick=" & _
					strWQ & "cmdSearch_Click " & strField_6 & strWQ & _
					">�ҏW</button><br/>" & "</td>"

		strHTML = strHTML & _
			"</tr>" & vbCrLf

	    objRecordset.MoveNext
		RecCnt = RecCnt + 1
	Loop
	RESULT_LIST.innerHTML = strHTML
'//	ID_FRM_CONTENT.document.body.innerHTML = strHTML
	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Function setKeyCondition( _
		byVal p_strCondType, _
		byVal p_strKeyword _
		)
	strKeyword = EncodeValue(UCase(p_strKeyword))
	Select Case p_strCondType
		case "1"	'// �L�[�Ɉ�v����
			strKeyCondition = " = " & _
				strKeyword & " "
		case "2"	'// �L�[�Ɉ�v���Ȃ�
			strKeyCondition = " not = " & _
				strKeyword & " "
		case "3"	'// �L�[�Ŏn�܂�
			strKeyCondition = " like '" & _
				p_strKeyword & "%' "
		case "4"	'// �L�[�ŏI���
			strKeyCondition = " like '%" & _
				p_strKeyword & "' "
		case "5"	'// �L�[���܂�
			strKeyCondition = " like '%" & _
				p_strKeyword & "%' "
		case "6"	'// �L�[���܂܂Ȃ�
			strKeyCondition = " not like '%" & _
				p_strKeyword & "%' "
		case "7"	'// �L�[���傫��
			strKeyCondition = " > " & _
				strKeyword & " "
		case "8"	'// �L�[��菬����
			strKeyCondition = " < " & _
				strKeyword & " "
		case "9"	'// �L�[���ȏ�
			strKeyCondition = " >= " & _
				strKeyword & " "
		case else	'// �L�[���ȉ�
			strKeyCondition = " <= " & _
				strKeyword & " "
	End Select

	setKeyCondition = strKeyCondition

End Function
Private Function makeDiscID(ByVal p_strDiscNo)
	makeDiscID = ""

	makeDiscID = "DVD-" & _
		String(4 - len(p_strDiscNo), "0") & _
		p_strDiscNo
End Function
Private Function EncodeValue(byval p_strItemValue)
Const adVarWChar = 202
	Dim strHTML
	Dim strWQ

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT * FROM T_�^��L�^ where �^��L�^�ԍ� = 1"

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	objRecordset.MoveFirst

	strWQ=""
	If (objRecordset.Fields.Item(KEY_ITEM.value).Type = adVarWChar) Then
		strWQ="'"
	End If

	strHTML = strWQ & p_strItemValue & strWQ

	'//DEBUGLOG strHTML & objRecordset.Fields.Item(KEY_ITEM.value).Type & strWQ

	EncodeValue = strHTML

	objRecordSet.Close()
	objConnection.Close()

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Function
Private Sub cmdTest_Click
	Dim strHTML

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT * FROM T_�^��L�^ where �^��L�^�ԍ� = 1"

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	Dim objField
	Do Until (objRecordset.EOF)
		For Each objField in objRecordset.Fields
			strHTML = strHTML & objField.Name & "=" & objField.Type & vbNewLine
		Next
	    objRecordset.MoveNext
	Loop

	DEBUGLOG strHTML

	objRecordSet.Close()
	objConnection.Close()

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Function ValidateData()
	ValidateData = false

	'// �L�[����
	If (KEY_ITEM.value = "-") Then
		Msgbox("�L�[���ڂ��I������Ă��܂���B")
		Exit Function
	End If

	'// �l
	If (KEYWORD.value = "") Then
		Msgbox("�l�����͂���Ă��܂���B")
		Exit Function
	End If

	'// ����
	If (KEY_CONDITION.value = "-") Then
		'// �f�t�H���g�����i1:��v����j��ݒ肷��B
		KEY_CONDITION.value = 1
		'//--Msgbox("�������I������Ă��܂���B")
		'//--Exit Function
	End If

	'// ���ލ���
	If (F_ORDERBY.value = "-") Then
		'// �f�t�H���g���ލ��ځi�^��L�^�ԍ��j��ݒ肷��B
		F_ORDERBY.value = "�^��L�^�ԍ�"
		'//--Msgbox("���ލ��ڂ��I������Ă��܂���B")
		'//--Exit Function
	End If

	ValidateData = true
End Function
Private Sub DEBUGLOG(byval p_strMessage)
	Dim strMessage

	If (isObject(g_objDebugView) = false) Then Exit Sub

	If (g_objDebugView.innerText = "") Then
		strMessage = p_strMessage
	Else
		strMessage = _
			g_objDebugView.innerText & vbNewLine & _
			p_strMessage & vbNewLine
	End If

	g_objDebugView.innerText = strMessage
End Sub
Private Sub MakeListBox( _
	byval p_strTextFile, _
	byref p_objConditionList)
	Dim arrConditionList

    ForReading = 1
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile _
        (p_strTextFile, ForReading)
    Do Until objFile.AtEndOfStream
        strLine = objFile.ReadLine
		arrConditionList = Split(strLine , ",")
        Set objOption = Document.createElement("OPTION")
        objOption.Text = arrConditionList(1)
        objOption.Value = arrConditionList(0)
        p_objConditionList.Add(objOption)
    Loop
    objFile.Close

	Set arrConditionList = Nothing
End Sub
Private Sub TestClick()
	DEBUGLOG TEST_LIST_BOX.value
End Sub

Private Sub viewData()
	Dim strFileType
	Dim strWQ

	'// ���͏����̑Ó����`�F�b�N
	If (Not ValidateData()) Then
		Exit Sub
	End If

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	'//-- �L�[�l��ݒ�
	Dim strKeyValue
	strKeyValue = KEYWORD.value
	If (KEY_ITEM.value = "�f�B�X�N�h�c") Then
		strKeyValue = makeDiscID(strKeyValue)
	End If

	'//-- �L�[�����`�F�b�N
	strConditionType = KEY_CONDITION.value

	'//-- �L�[�����쐬
	strWhere = ""
	If (KEY_ITEM.value <> "") Then
		strWhere = _
			KEY_ITEM.value & " " & _
			setKeyCondition( _
				strConditionType, _
				strKeyValue)
	End If

	strSQL = "SELECT * FROM T_�^��L�^ "
	If (strWhere <> "") Then
		strSQL = strSQL & _
				"where " & _
				strWhere
	End If

	strSQL = strSQL & _
			"order by " & F_ORDERBY.value

	DEBUGLOG strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	strHTML="<table border='1' cellspacing='0'>" & vbCrLf & _
		"<tr style='background-color:skyblue;'>" & _
		"<th>�f�B�X�N�h�c</th>" & _
		"<th>�{�����[���ԍ�</th>" & _
		"<th>�ԑg��</th>" & _
		"<th>�^�C�g����</th>" & _
		"<th>�^����t</th>" & _
		"<th>�^�掞��</th>" & _
		"<th>�폜�t���O</th>" & _
		"<th>�ҏW�t���O</th>" & _
		"<th>�A�N�V����</th>" & _
		"</tr>"
	If (objRecordset.RecordCount > 999) Then
		'//Msgbox "���R�[�h������1000���𒴂��Ă��܂��B" & _
		'//"rec-cnt=" & objRecordset.RecordCount
	End If

	RecCnt=0

	Do Until (objRecordset.EOF)
		strField_1 = objRecordset.Fields.Item("�f�B�X�N�h�c")
		strField_2 = objRecordset.Fields.Item("�^�C�g����")
		strField_3 = objRecordset.Fields.Item("�^����t")
		strField_4 = objRecordset.Fields.Item("�^�掞��")
		strField_5 = objRecordset.Fields.Item("�{�����[���ԍ�")
		strField_6 = objRecordset.Fields.Item("�^��L�^�ԍ�")
		strField_7 = objRecordset.Fields.Item("�폜�t���O")
		strField_8 = objRecordset.Fields.Item("�ҏW�t���O")
		strField_9 = objRecordset.Fields.Item("�ԑg��")
		strHTML = strHTML & _
			"<tr>" & vbCrLf

	    strHTML = strHTML & _
			"<td nowrap>" & strField_1 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_5 & "<br/>" & "</td>" & _
			"<td>" & strField_9 & "<br/>" & "</td>" & _
			"<td>" & strField_2 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_3 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_4 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_7 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_8 & "<br/>" & "</td>" & _
			"<td nowrap><button onclick=" & _
					strWQ & "cmdSearch_Click " & strField_6 & strWQ & _
					">�ҏW</button><br/>" & "</td>"

		strHTML = strHTML & _
			"</tr>" & vbCrLf

	    objRecordset.MoveNext
		RecCnt = RecCnt + 1
	Loop
'	set new_win = window.open("viewData.htm")
	new_win = window.showModalDialog("viewData.htm")
'	if (isObject(new_win)) Then
'		new_win.document.body.InnerHTML = strHTML
'	End If
'    Set objIE = CreateObject("InternetExplorer.Application")
'    objIE.Navigate("about:blank")
'    objIE.ToolBar = 0
'    objIE.StatusBar = 0
'    Set objDoc = objIE.Document.Body
'    objDoc.InnerHTML = strHTML
'    objIE.Visible = True

'//	RESULT_LIST.innerHTML = strHTML
'//	ID_FRM_CONTENT.document.body.innerHTML = strHTML
	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Sub viewList()
	Dim strFileType
	Dim strWQ

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = F_SQLCOMMAND.value

	DEBUGLOG strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	strHTML="<table border=""1"" cellspacing=""0"">" & vbCrLf & _
		"<tr style=""background-color:skyblue;"">" & vbCrLf

	For Each f in objRecordset.Fields
		strHTML = strHTML & _
			"<th nowrap=""yes"">" & f.name & "</th>" & vbCrLf
	Next

	strHTML = strHTML & _
		"</tr>" & vbCrLf

	Do Until (objRecordset.EOF)
		strHTML = strHTML & _
			"<tr>" & vbCrLf

		For Each f in objRecordset.Fields
			strHTML = strHTML & _
				"<td nowrap=""yes"">" & f.value & "<br/></td>" & vbCrLf
		Next

		strHTML = strHTML & _
			"</tr>" & vbCrLf

	    objRecordset.MoveNext
		RecCnt = RecCnt + 1
	Loop

	RESULT_LIST.innerHTML = strHTML

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
-->
