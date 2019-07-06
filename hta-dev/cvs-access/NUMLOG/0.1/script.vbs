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

	'// キー項目が選択されている場合、条件選択のデフォルト条件を設定する。
	If (SELECT_KEY.value <> "0" _
	And	KEY_CONDITION.value = "-") Then
		'// デフォルト（1:一致する）を設定する。
		KEY_CONDITION.value = 1
	End If

	'// 分類項目が選択されているかチェックする。
	If (ORDER_BY.value = "-") Then
		'// 分類項目のデフォルト項目を設定する。
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

	'//-- キー条件チェック
	strConditionType = KEY_CONDITION.value

	'//-- キー選択による[where]条件分作成
	strWhere = ""

	If (SELECT_KEY.value <> "-") Then
			strWhere = _
				SELECT_KEY.value & " " & _
				setKeyCondition( _
					strConditionType, _
					KEYWORD.value)
	End If

	'// ファイル選択
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
		"<th>番号</th>" & _
		"<th>カレント・ディレクトリ</th>" & _
		"<th>日付</th>" & _
		"<th>時間</th>" & _
		"<th>編集フラグ</th>" & _
		"<th>削除フラグ</th>" & _
		"<th nowrap='yes'>アクション</th>" & _
		"</tr>"
	If (objRecordset.RecordCount > 999) Then
		'//Msgbox "レコード件数が1000件を超えています。" & _
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

		'// 録画番号フォルダごとの削除フラグ、編集フラグを取得する。
		blnDeleteFlag = GetDeleteFlag(strField_1)
		blnEditFlag = GetEditFlag(strField_1)

		'// 出力フラグを初期化する。
		blnOutFlag = False

		'// 未編集チェックがTRUEの場合、
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
				"削除</button>" & vbNewLine

			strHTML = strHTML & "<button onclick=" & strWQ & "cmdEditFile '" & _
				strField_4 & strField_3 & "'" & strWQ & ">" & _
				"編集</button></td>" & vbNewLine

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
		case "1"	'// キーに一致する
			strKeyCondition = " = '" & _
				strKeyword & "' "
		case "2"	'// キーに一致しない
			strKeyCondition = " not = '" & _
				strKeyword & "' "
		case "3"	'// キーで始まる
			strKeyCondition = " like '" & _
				strKeyword & "%' "
		case "4"	'// キーで終わる
			strKeyCondition = " like '%" & _
				strKeyword & "' "
		case "5"	'// キーを含む
			strKeyCondition = " like '%" & _
				strKeyword & "%' "
		case "6"	'// キーを含まない
			strKeyCondition = " not like '%" & _
				strKeyword & "%' "
		case "7"	'// キーより大きい
			strKeyCondition = " > '" & _
				strKeyword & "' "
		case "8"	'// キーより小さい
			strKeyCondition = " < '" & _
				strKeyword & "' "
		case "9"	'// キーより以上
			strKeyCondition = " >= '" & _
				strKeyword & "' "
		case else	'// キーより以下
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
	'// デバッグ表示オブジェクトを設定する。
	Set g_objDebugView = TEXT_DATA

	'// キー項目リストをＤＢ情報から作成する。
	Call MakeItems(ID_SELECT_KEY, "--キー項目--")

	'// 条件リストボックスをテキストファイルから作成する。
	Call MakeListBox("ConditionList.txt", KEY_CONDITION)

	'// 分類項目リストボックスをＤＢ情報から作成する。
	Call MakeItems(ID_ORDER_BY, "--分類項目--")

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

	'// フォルダ名を取得する。
	Set objFolder=objFSO.GetFolder(strFolderName)

	'// 削除する前にフォルダ名をフォルダ番号に保存する。
	strFolderNo = objFolder.Name
	
	blnYesNo = Msgbox("フォルダ " & strFolderName & " を削除します。" , _
			vbYesNo, "フォルダ削除")

	If (blnYesNo = vbYes) Then
		Call objFSO.DeleteFolder(strFolderName)
		If Err Then
			Msgbox Err.Description
			Exit Sub
		End If
		Msgbox "フォルダ " & strFolderName & " を削除しました。"
		If (WriteLog(strFolderName)) Then
			Msgbox "ログファイルに結果を出力しました。" & vbNewLine & _
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
		Msgbox "フォルダが見つかりません。"
		Set objFSO = Nothing
		Exit Sub
	End If

	'// フォルダ名を取得する。
	Set objFolder=objFSO.GetFolder(strFolderName)

	'// 削除する前にフォルダ名をフォルダ番号に保存する。
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

	'// ログファイルが存在チェックするか？
	blnFileExist = False
	if (objFSO.FileExists(DC_LOGFILE)) Then
		'// ファイル存在フラッグをＯＮにする。
		blnFileExist = True

		'// ログファイルを追加書き込みモードでオープンする。
		Set objLogFile = objFSO.OpenTextFile(DC_LOGFILE, ForAppending, True)
	Else
		'// ログファイルを書き込みモードでオープンする。
		Set objLogFile = objFSO.OpenTextFile(DC_LOGFILE, ForWriting, True)
	End If

	If (Err) Then
		Msgbox "ファイルオープンエラー" & vbNewLine & Err.Description
		WriteLog = False
		Exit Function
	End If
	
	'// ファイルが存在しない場合、ヘッダレコードを出力する。
	If (Not blnFileExist) Then
		objLogFile.WriteLine "フォルダ名,削除日付,削除時間"
	End If

	'// ログファイルに削除フォルダ、削除日、削除時間を出力する。
	objLogFile.WriteLine strFolderName & "," & _
				Date() & "," & _
				Time()
				
	If (Err) Then
		Msgbox "ファイルライトエラー" & vbNewLine & Err.Description
		WriteLog = False
		Exit Function
	End If
	
	'// ログファイルをクローズする。
	objLogFile.Close

	If (Err) Then
		Msgbox "ファイルクローズエラー" & vbNewLine & Err.Description
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

	'// ディスクＩＤを作成する。
	If (InStr(intFolderNo, "_")) Then
		objArr = Split(intFolderNo, "_", -1)
		intFolderNo = objArr(0)
	End If

	'//msgbox "フォルダ番号" & intFolderNo
	strDiscID = makeDiscID(intFolderNo)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT 削除フラグ FROM T_録画記録 where ディスクＩＤ = '" & strDiscID & "'"

	'//TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	Do Until (objRecordset.EOF)
		If (objRecordset.Fields.Item("削除フラグ") = True) Then
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

	'// ディスクＩＤを作成する。
	If (InStr(intFolderNo, "_")) Then
		objArr = Split(intFolderNo, "_", -1)
		intFolderNo = objArr(0)
	End If

	'//msgbox "フォルダ番号" & intFolderNo
	strDiscID = makeDiscID(intFolderNo)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT 編集フラグ FROM T_録画記録 where ディスクＩＤ = '" & strDiscID & "'"

	'//TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	Do Until (objRecordset.EOF)
		GetEditFlag = objRecordset.Fields.Item("編集フラグ")
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

	'// ディスクＩＤを作成する。
	strDiscID = makeDiscID(intFolderNo)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT 削除フラグ, 削除日付,編集フラグ FROM T_録画記録 WHERE ディスクＩＤ = '" & strDiscID & "'"

	TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	Do Until (objRecordset.EOF)
		With objRecordset
			.Fields.Item("削除フラグ") = True
			.Fields.Item("編集フラグ") = True
			.Fields.Item("削除日付") = Date()
			.Update
		End With

		'// 更新エラーチェック
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
	strHTML = "<div>該当データ " & objRecordSet.RecordCount & " 件が正常に更新されました。</div>"
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

	'// ディスクＩＤを作成する。
	strDiscID = makeDiscID(intFolderNo)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT 編集フラグ FROM T_録画記録 WHERE ディスクＩＤ = '" & strDiscID & "'"

	TEXT_DATA.innerText = strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	Do Until (objRecordset.EOF)
		With objRecordset
			.Fields.Item("編集フラグ") = True
			.Update
		End With

		'// 更新エラーチェック
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
	strHTML = "<div>該当データ " & objRecordSet.RecordCount & " 件が正常に更新されました。</div>"
	RESULT.innerHTML= strHTML

	Set objRecordSet = Nothing
	Set objConnection = Nothing
End Sub
Private Function makeDiscID(byVal intFolderNo)
	If (InStr(1, intFolderNo, "_")) Then
		intFolderNo = Left(1, InStr(1, intFolderNo, "_") - 1)
	End If
'//	TEXT_DATA.innerText = TEXT_DATA.innerText & _
'//		"フォルダ番号=" & intFolderNo & vbNewLine
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

	'//-- キー条件チェック
	strConditionType = KEY_CONDITION.value

	'//-- キー選択による[where]条件分作成
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

	'// ファイル選択
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

	RESULT.innerHTML = "<div>正常にデータベースが更新されました。</div>"

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

	'//-- キー条件チェック
	strConditionType = KEY_CONDITION.value

	'//-- キー選択による[where]条件分作成
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

	'// ファイル選択
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
		"<th>番号</th>" & _
		"<th>カレント・ディレクトリ</th>" & _
		"<th>日付</th>" & _
		"<th>時間</th>" & _
		"<th>フォルダ有無</th>" & _
		"<th>削除フラグ</th>" & _
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
			strFolderExist = "○"
		Else
			strFolderExist = "×"
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

	'// ダウンリストのタイトルデータを設定する。
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

