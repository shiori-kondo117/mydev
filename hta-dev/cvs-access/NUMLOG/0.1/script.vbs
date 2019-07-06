Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001
Const DC_CSVFILE_PATH = "D:\home\kazu\MyData\Database\TempFolderNumber\"
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const DC_LOGFILE = "D:\temp\DeleteFile.log"
'Const vbLongDate = 1
'Const vbLongTime = 3

On Error Resume Next

Private Sub cmdSearch_Click
	Dim strFileType
	Dim strWQ

	'// �L�[���ڂ��I������Ă���ꍇ�A�����I���̃f�t�H���g������ݒ肷��B
	If (SELECT_KEY.value <> "0" _
	And	KEY_CONDITION.value = "-") Then
		'// �f�t�H���g�i1:��v����j��ݒ肷��B
		KEY_CONDITION.value = 1
	End If

	'// ���ލ��ڂ��I������Ă��邩�`�F�b�N����B
	If (ORDER_BY.value = "-") Then
		'// ���ލ��ڂ̃f�t�H���g���ڂ�ݒ肷��B
		ORDER_BY.value = "Number"
	End If

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strPathtoTextFile = DC_CSVFILE_PATH

	objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	          "Data Source=" & strPathtoTextFile & ";" & _
	          "Extended Properties=""text;HDR=YES;FMT=Delimited"""

	strSelectType = SELECT_KEY.value

	'//-- �L�[�����`�F�b�N
	strConditionType = KEY_CONDITION.value

	'//-- �L�[�I���ɂ��[where]�������쐬
	strWhere = ""

	If (SELECT_KEY.value <> "-") Then
			strWhere = _
				SELECT_KEY.value & " " & _
				setKeyCondition( _
					strConditionType, _
					KEYWORD.value)
	End If

	'// �t�@�C���I��
	for each x in FILE_TYPE
		if (x.checked) then
			strFileType = x.value
			exit for
		end if
	next

	strSQL = "SELECT * FROM " & strFileType & " "
	If (strWhere <> "") Then
		strSQL = strSQL & _
				"where " & _
				strWhere
	End If

	strSQL = strSQL & _
			"order by Number"

	TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	strHTML="<table border='1' cellspacing='0'>" & vbCrLf & _
		"<tr style='background-color:skyblue;'>" & _
		"<th>�ԍ�</th>" & _
		"<th>�J�����g�E�f�B���N�g��</th>" & _
		"<th>���t</th>" & _
		"<th>����</th>" & _
		"<th>�ҏW�t���O</th>" & _
		"<th>�폜�t���O</th>" & _
		"<th nowrap='yes'>�A�N�V����</th>" & _
		"</tr>"
	If (objRecordset.RecordCount > 999) Then
		'//Msgbox "���R�[�h������1000���𒴂��Ă��܂��B" & _
		'//"rec-cnt=" & objRecordset.RecordCount
	End If

	RecCnt=0

	Dim blnDeleteFlag, blnEditFlag, blnOutFlag
	Do Until (objRecordset.EOF)
		strField_1 = objRecordset.Fields.Item("Number")
		strField_2 = objRecordset.Fields.Item("Old")
		strField_3 = objRecordset.Fields.Item("New")
		strField_4 = objRecordset.Fields.Item("CurrentDirectory")
		strField_5 = objRecordset.Fields.Item("Date")
		strField_6 = objRecordset.Fields.Item("Time")

		'// �^��ԍ��t�H���_���Ƃ̍폜�t���O�A�ҏW�t���O���擾����B
		blnDeleteFlag = GetDeleteFlag(strField_1)
		blnEditFlag = GetEditFlag(strField_1)

		'// �o�̓t���O������������B
		blnOutFlag = False

		'// ���ҏW�`�F�b�N��TRUE�̏ꍇ�A
		If (NONE_EDIT.checked) Then
'//			If (Not blnDeleteFlag) Then
			If (Not blnEditFlag) Then
				blnOutFlag = True
			End If
		Else
			blnOutFlag = True
		End If

		If (blnOutFlag) Then
		
			strHTML = strHTML & _
				"<tr>" & vbCrLf

			'//strField_1 = String(3 - Len(strField_1), "0") & strField_1

		    strHTML = strHTML & _
				"<td nowrap>" & strField_1 & "<br/>" & "</td>" & _
				"<td><a href='" & strField_4 & strField_3 & "'>" & _
					strField_4 & strField_3 & "</a><br/>" & "</td>" & _
				"<td nowrap>" & strField_5 & "<br/>" & "</td>" & _
				"<td nowrap>" & strField_6 & "<br/>" & "</td>"

			strHTML = strHTML & "<td>" & GetEditFlag(strField_1) & "</td>"
			strHTML = strHTML & "<td>" & blnDeleteFlag & "</td>"
			strHTML = strHTML & "<td><button onclick=" & strWQ & "cmdDeleteFile '" & _
				strField_4 & strField_3 & "'" & strWQ & ">" & _
				"�폜</button>" & vbNewLine

			strHTML = strHTML & "<button onclick=" & strWQ & "cmdEditFile '" & _
				strField_4 & strField_3 & "'" & strWQ & ">" & _
				"�ҏW</button></td>" & vbNewLine

			strHTML = strHTML & _
				"</tr>" & vbCrLf
		End If
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
	strKeyword = UCase(p_strKeyword)
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
		case "6"	'// �L�[���܂܂Ȃ�
			strKeyCondition = " not like '%" & _
				strKeyword & "%' "
		case "7"	'// �L�[���傫��
			strKeyCondition = " > '" & _
				strKeyword & "' "
		case "8"	'// �L�[��菬����
			strKeyCondition = " < '" & _
				strKeyword & "' "
		case "9"	'// �L�[���ȏ�
			strKeyCondition = " >= '" & _
				strKeyword & "' "
		case else	'// �L�[���ȉ�
			strKeyCondition = " <= '" & _
				strKeyword & "' "
	End Select

	setKeyCondition = strKeyCondition

End Function

Private Sub cmdClear_Click
	SELECT_KEY.value = "-"
	KEYWORD.value = ""
	KEY_CONDITION.value = "-"
End Sub
Private Sub Setup()
	'// �f�o�b�O�\���I�u�W�F�N�g��ݒ肷��B
	Set g_objDebugView = TEXT_DATA

	'// �L�[���ڃ��X�g���c�a��񂩂�쐬����B
	Call MakeItems(ID_SELECT_KEY, "--�L�[����--")

	'// �������X�g�{�b�N�X���e�L�X�g�t�@�C������쐬����B
	Call MakeListBox("ConditionList.txt", KEY_CONDITION)

	'// ���ލ��ڃ��X�g�{�b�N�X���c�a��񂩂�쐬����B
	Call MakeItems(ID_ORDER_BY, "--���ލ���--")

End Sub
Private Sub cmdDeleteFile(strFolderName)
'//	On Error Resume Next

	Dim objFSO
	Dim blnYesNo
	Dim objLogFile
	Dim objFolder
	Dim blnFileExist
	Dim strFolderNo

    Set objFSO = CreateObject("Scripting.FileSystemObject")

	'// �t�H���_�����擾����B
	Set objFolder=objFSO.GetFolder(strFolderName)

	'// �폜����O�Ƀt�H���_�����t�H���_�ԍ��ɕۑ�����B
	strFolderNo = objFolder.Name
	
	blnYesNo = Msgbox("�t�H���_ " & strFolderName & " ���폜���܂��B" , _
			vbYesNo, "�t�H���_�폜")

	If (blnYesNo = vbYes) Then
		Call objFSO.DeleteFolder(strFolderName)
		If Err Then
			Msgbox Err.Description
			Exit Sub
		End If
		Msgbox "�t�H���_ " & strFolderName & " ���폜���܂����B"
		If (WriteLog(strFolderName)) Then
			Msgbox "���O�t�@�C���Ɍ��ʂ��o�͂��܂����B" & vbNewLine & _
				"(" & DC_LOGFILE & ")"
		End If
		Call UpdateDeleteFlag(strFolderNo)
	End If

	Set objFSO = Nothing
End Sub

Private Sub cmdEditFile(strFolderName)
'//	On Error Resume Next

	Dim objFSO
	Dim blnYesNo
	Dim objLogFile
	Dim objFolder
	Dim blnFileExist
	Dim strFolderNo

    Set objFSO = CreateObject("Scripting.FileSystemObject")

	If (Not objFSO.FolderExists(strFolderName)) Then
		Msgbox "�t�H���_��������܂���B"
		Set objFSO = Nothing
		Exit Sub
	End If

	'// �t�H���_�����擾����B
	Set objFolder=objFSO.GetFolder(strFolderName)

	'// �폜����O�Ƀt�H���_�����t�H���_�ԍ��ɕۑ�����B
	strFolderNo = objFolder.Name
	
	Call UpdateEditFlag(strFolderNo)

	Set objFSO = Nothing
End Sub
Private Function WriteLog(byVal strFolderName)

	Dim objFSO
	Dim objLogFile
	Dim blnFileExist

	On Error Resume Next

    Set objFSO = CreateObject("Scripting.FileSystemObject")

	'// ���O�t�@�C�������݃`�F�b�N���邩�H
	blnFileExist = False
	if (objFSO.FileExists(DC_LOGFILE)) Then
		'// �t�@�C�����݃t���b�O���n�m�ɂ���B
		blnFileExist = True

		'// ���O�t�@�C����ǉ��������݃��[�h�ŃI�[�v������B
		Set objLogFile = objFSO.OpenTextFile(DC_LOGFILE, ForAppending, True)
	Else
		'// ���O�t�@�C�����������݃��[�h�ŃI�[�v������B
		Set objLogFile = objFSO.OpenTextFile(DC_LOGFILE, ForWriting, True)
	End If

	If (Err) Then
		Msgbox "�t�@�C���I�[�v���G���[" & vbNewLine & Err.Description
		WriteLog = False
		Exit Function
	End If
	
	'// �t�@�C�������݂��Ȃ��ꍇ�A�w�b�_���R�[�h���o�͂���B
	If (Not blnFileExist) Then
		objLogFile.WriteLine "�t�H���_��,�폜���t,�폜����"
	End If

	'// ���O�t�@�C���ɍ폜�t�H���_�A�폜���A�폜���Ԃ��o�͂���B
	objLogFile.WriteLine strFolderName & "," & _
				Date() & "," & _
				Time()
				
	If (Err) Then
		Msgbox "�t�@�C�����C�g�G���[" & vbNewLine & Err.Description
		WriteLog = False
		Exit Function
	End If
	
	'// ���O�t�@�C�����N���[�Y����B
	objLogFile.Close

	If (Err) Then
		Msgbox "�t�@�C���N���[�Y�G���[" & vbNewLine & Err.Description
		WriteLog = False
		Exit Function
	End If
	
	WriteLog = True

	Set objLogFile = Nothing
	Set objFSO = Nothing
End Function
Private Sub Window_onload()
	Call window.resizeTo(640, 480)
End Sub
Private Function GetDeleteFlag(byVal intFolderNo)
	Dim strHTML
	Dim strDiscID
	Dim objArr

	GetDeleteFlag = False
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'// �f�B�X�N�h�c���쐬����B
	If (InStr(intFolderNo, "_")) Then
		objArr = Split(intFolderNo, "_", -1)
		intFolderNo = objArr(0)
	End If

	'//msgbox "�t�H���_�ԍ�" & intFolderNo
	strDiscID = makeDiscID(intFolderNo)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT �폜�t���O FROM T_�^��L�^ where �f�B�X�N�h�c = '" & strDiscID & "'"

	'//TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	Do Until (objRecordset.EOF)
		If (objRecordset.Fields.Item("�폜�t���O") = True) Then
			GetDeleteFlag = True
		End If
	    objRecordset.MoveNext
	Loop
	strHTML = "<div>" & GetDeleteFlag & "</div>"
	RESULT.innerHTML= strHTML

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Function
Private Function GetEditFlag(byVal intFolderNo)
	Dim strHTML
	Dim strDiscID
	Dim objArr

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'// �f�B�X�N�h�c���쐬����B
	If (InStr(intFolderNo, "_")) Then
		objArr = Split(intFolderNo, "_", -1)
		intFolderNo = objArr(0)
	End If

	'//msgbox "�t�H���_�ԍ�" & intFolderNo
	strDiscID = makeDiscID(intFolderNo)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT �ҏW�t���O FROM T_�^��L�^ where �f�B�X�N�h�c = '" & strDiscID & "'"

	'//TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	Do Until (objRecordset.EOF)
		GetEditFlag = objRecordset.Fields.Item("�ҏW�t���O")
	    objRecordset.MoveNext
	Loop
	strHTML = "<div>" & GetEditFlag & "</div>"
	RESULT.innerHTML= strHTML

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Function
Private Sub UpdateDeleteFlag(byVal intFolderNo)
	Dim strHTML
	Dim strDiscID

	Err.Clear

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'// �f�B�X�N�h�c���쐬����B
	strDiscID = makeDiscID(intFolderNo)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT �폜�t���O, �폜���t,�ҏW�t���O FROM T_�^��L�^ WHERE �f�B�X�N�h�c = '" & strDiscID & "'"

	TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	Do Until (objRecordset.EOF)
		With objRecordset
			.Fields.Item("�폜�t���O") = True
			.Fields.Item("�ҏW�t���O") = True
			.Fields.Item("�폜���t") = Date()
			.Update
		End With

		'// �X�V�G���[�`�F�b�N
		If (Err) Then
			Msgbox Err.Description
			objRecordSet.Close()
			objConnection.Close()
			Set objRecordSet = Nothing
			Set objConnection = Nothing
			Exit Sub
		End If

	    objRecordset.MoveNext
	Loop
	strHTML = "<div>�Y���f�[�^ " & objRecordSet.RecordCount & " ��������ɍX�V����܂����B</div>"
	RESULT.innerHTML= strHTML

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Sub UpdateEditFlag(byVal intFolderNo)
	Dim strHTML
	Dim strDiscID

	Err.Clear

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'// �f�B�X�N�h�c���쐬����B
	strDiscID = makeDiscID(intFolderNo)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT �ҏW�t���O FROM T_�^��L�^ WHERE �f�B�X�N�h�c = '" & strDiscID & "'"

	TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	Do Until (objRecordset.EOF)
		With objRecordset
			.Fields.Item("�ҏW�t���O") = True
			.Update
		End With

		'// �X�V�G���[�`�F�b�N
		If (Err) Then
			Msgbox Err.Description
			objRecordSet.Close()
			objConnection.Close()
			Set objRecordSet = Nothing
			Set objConnection = Nothing
			Exit Sub
		End If

	    objRecordset.MoveNext
	Loop
	strHTML = "<div>�Y���f�[�^ " & objRecordSet.RecordCount & " ��������ɍX�V����܂����B</div>"
	RESULT.innerHTML= strHTML

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Function makeDiscID(byVal intFolderNo)
	If (InStr(1, intFolderNo, "_")) Then
		intFolderNo = Left(1, InStr(1, intFolderNo, "_") - 1)
	End If
'//	TEXT_DATA.innerText = TEXT_DATA.innerText & _
'//		"�t�H���_�ԍ�=" & intFolderNo & vbNewLine
	makeDiscID = "DVD-" & String(4 - len(intFolderNo), "0") & intFolderNo
End Function
Private Sub modifyDatabase()
	Dim strFileType
	Dim strWQ
	Dim strFolderExist

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strPathtoTextFile = DC_CSVFILE_PATH

	objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	          "Data Source=" & strPathtoTextFile & ";" & _
	          "Extended Properties=""text;HDR=YES;FMT=Delimited"""

	strSelectType = SELECT_KEY.value

	'//-- �L�[�����`�F�b�N
	strConditionType = KEY_CONDITION.value

	'//-- �L�[�I���ɂ��[where]�������쐬
	Select Case strSelectType
		case "1"
			strWhere = _
				"Number " & _
				setKeyCondition( _
					strConditionType, _
					KEYWORD.value)
		case "2"
			strWhere = _
				"CurrentDirectory " & _
				setKeyCondition( _
					strConditionType, _
					KEYWORD.value)
		case "3"
			strWhere = _
				"Date " & _
				setKeyCondition( _
					strConditionType, _
					KEYWORD.value)
		case else
			strWhere = ""
	End Select

	'// �t�@�C���I��
	for each x in FILE_TYPE
		if (x.checked) then
			strFileType = x.value
			exit for
		end if
	next

	strSQL = "SELECT * FROM " & strFileType & " "
	If (strWhere <> "") Then
		strSQL = strSQL & _
				"where " & _
				strWhere
	End If

	strSQL = strSQL & _
			"order by Number"

	TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	Do Until (objRecordset.EOF)
		strField_1 = objRecordset.Fields.Item("Number")
		strField_2 = objRecordset.Fields.Item("Old")
		strField_3 = objRecordset.Fields.Item("New")
		strField_4 = objRecordset.Fields.Item("CurrentDirectory")
		strField_5 = objRecordset.Fields.Item("Date")
		strField_6 = objRecordset.Fields.Item("Time")
		strHTML = strHTML & _
			"<tr>" & vbCrLf

		If (isFolderExist(strField_4 & strField_3) = false) Then
			Call UpdateDeleteFlag(strField_3)
		End If

	    objRecordset.MoveNext
	Loop

	RESULT.innerHTML = "<div>����Ƀf�[�^�x�[�X���X�V����܂����B</div>"

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Function isFolderExist(byval p_strFolderName)
	Dim objFSO

    Set objFSO = CreateObject("Scripting.FileSystemObject")

	isFolderExist = false
	
	If (objFSO.FolderExists(p_strFolderName)) Then
		isFolderExist = true
	End If
'//	TEXT_DATA.innerText = TEXT_DATA.innerText & vbNewLine & _
'//		p_strFolderName & "=" & objFSO.FolderExists(p_strFolderName) & _
'//		vbNewLine
	Set objFSO = Nothing
End Function
Private Sub checkDeleteFolder()
	Dim strFileType
	Dim strWQ
	Dim strFolderExist

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strPathtoTextFile = DC_CSVFILE_PATH

	objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	          "Data Source=" & strPathtoTextFile & ";" & _
	          "Extended Properties=""text;HDR=YES;FMT=Delimited"""

	strSelectType = SELECT_KEY.value

	'//-- �L�[�����`�F�b�N
	strConditionType = KEY_CONDITION.value

	'//-- �L�[�I���ɂ��[where]�������쐬
	Select Case strSelectType
		case "1"
			strWhere = _
				"Number " & _
				setKeyCondition( _
					strConditionType, _
					KEYWORD.value)
		case "2"
			strWhere = _
				"CurrentDirectory " & _
				setKeyCondition( _
					strConditionType, _
					KEYWORD.value)
		case "3"
			strWhere = _
				"Date " & _
				setKeyCondition( _
					strConditionType, _
					KEYWORD.value)
		case else
			strWhere = ""
	End Select

	'// �t�@�C���I��
	for each x in FILE_TYPE
		if (x.checked) then
			strFileType = x.value
			exit for
		end if
	next

	strSQL = "SELECT * FROM " & strFileType & " "
	If (strWhere <> "") Then
		strSQL = strSQL & _
				"where " & _
				strWhere
	End If

	strSQL = strSQL & _
			"order by Number"

	TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	strHTML="<table border=""1"" cellspacing=""0"">" & vbCrLf & _
		"<tr style='background-color:skyblue;'>" & _
		"<th>�ԍ�</th>" & _
		"<th>�J�����g�E�f�B���N�g��</th>" & _
		"<th>���t</th>" & _
		"<th>����</th>" & _
		"<th>�t�H���_�L��</th>" & _
		"<th>�폜�t���O</th>" & _
		"</tr>"

	Do Until (objRecordset.EOF)
		strField_1 = objRecordset.Fields.Item("Number")
		strField_2 = objRecordset.Fields.Item("Old")
		strField_3 = objRecordset.Fields.Item("New")
		strField_4 = objRecordset.Fields.Item("CurrentDirectory")
		strField_5 = objRecordset.Fields.Item("Date")
		strField_6 = objRecordset.Fields.Item("Time")
		strHTML = strHTML & _
			"<tr>" & vbCrLf

		'//strField_1 = String(3 - Len(strField_1), "0") & strField_1

	    strHTML = strHTML & _
			"<td nowrap>" & strField_1 & "<br/>" & "</td>" & _
			"<td><a href='" & strField_4 & strField_3 & "'>" & _
				strField_4 & strField_3 & "</a><br/>" & "</td>" & _
			"<td nowrap>" & strField_5 & "<br/>" & "</td>" & _
			"<td nowrap>" & strField_6 & "<br/>" & "</td>"
		If (isFolderExist(strField_4 & strField_3)) Then
			strFolderExist = "��"
		Else
			strFolderExist = "�~"
		End If
		strHTML = strHTML & "<td>" & strFolderExist & "</td>"
		strHTML = strHTML & "<td>" & GetDeleteFlag(strField_3) & "</td>"

		strHTML = strHTML & _
			"</tr>" & vbCrLf

	    objRecordset.MoveNext
	Loop
	RESULT.innerHTML = strHTML

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Sub sampGetFolderName(strFolderName)
'//	On Error Resume Next

	Dim objFSO
	Dim objFolder

    Set objFSO = CreateObject("Scripting.FileSystemObject")

	Set objFolder=objFSO.GetFolder(strFolderName)

	msgbox objFolder.Name
	
	Set objFolder = Nothing
	Set objFSO = Nothing
End Sub

Private Function MakeItems(byref p_objControl, byval p_strTitle)
	Dim objField
	Dim objOption

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	Call OpenDbTextFile(objConnection, DC_CSVFILE_PATH)

	strSQL = "SELECT * FROM numlog.txt"
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

Private Sub OpenDbTextFile(ByRef p_objConnection, ByVal p_TextFilePath)
	p_objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	          "Data Source=" & p_TextFilePath & ";" & _
	          "Extended Properties=""text;HDR=YES;FMT=Delimited"""
End Sub

