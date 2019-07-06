<!-- version:0.1.0.1
'//---------------------------------------------------------
'// グローバル変数 定義
'//---------------------------------------------------------
Public g_objDebugView			'// デバッグ表示オブジェクト名
'//---------------------------------------------------------
'// グローバル定数 定義
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

	strSQL = "SELECT * FROM L_録画記録 WHERE 録画記録番号 = " & intRecLogNo

	DEBUGLOG strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	'// 録画日付のフィールドを生成する。
	strHTML = "<input type=" & strWQ & "text" & strWQ & " " & _
		"name=" & strWQ & "REC_DATE" & strWQ & " " & _
		"size=" & strWQ & "13" & strWQ & "/>"
	ID_REC_DATE.innerHTML = strHTML

	'// 録画時間のフィールドを生成する。
	strHTML = "<input type=" & strWQ & "text" & strWQ & " " & _
		"name=" & strWQ & "REC_TIME" & strWQ & " " & _
		"size=" & strWQ & "9" & strWQ & "/>"
	ID_REC_TIME.innerHTML = strHTML

	Do Until (objRecordset.EOF)
		ID_REC_LOG_NO.innerText = objRecordset.Fields.Item("録画記録番号")
		DVD_NO.value = CInt(Right(objRecordset.Fields.Item("ディスクＩＤ"),4))
		If (objRecordset.Fields.Item("ボリューム番号") <> "") Then
			VOLUME_NO.value = objRecordset.Fields.Item("ボリューム番号")
		Else
			VOLUME_NO.value = ""
		End If
		BANGUMI_NAME.value = objRecordset.Fields.Item("番組名")
		TITLE_NAME.value = objRecordset.Fields.Item("タイトル名")
		REC_DATE.value = objRecordset.Fields.Item("録画日付")
		REC_TIME.value = objRecordset.Fields.Item("録画時間")
		REC_MODE.value = objRecordset.Fields.Item("録画モード")
		COPYONCE.checked = objRecordset.Fields.Item("コピーワンス")
		DELETE_FLAG.checked = objRecordset.Fields.Item("削除フラグ")
		F_EDIT_FLAG.checked = objRecordset.Fields.Item("編集フラグ")
		F_DELETE_DATE.value = objRecordset.Fields.Item("削除日付")
		F_MEMO.value = objRecordset.Fields.Item("メモ")
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

	'// トランザクションを開始する。
	objConnection.BeginTrans()

'	strSQL = "UPDATE T_録画記録 " & _
'			"set ディスクＩＤ = " & _
'				strWQ & "DVD-" & string(len(DVD_NO.value), "0") & DVD_NO.value & strWQ & "," & vbCrLf & _
'			"タイトル名 = " & strWQ & TITLE_NAME.value & strWQ & "," & vbCrLf & _
'			"録画日付 = " & strWQ & REC_DATE.value & strWQ & "," & vbCrLf & _
'			"録画時間 = " & strWQ & REC_TIME.value & strWQ & "," & vbCrLf & _
'			"ボリューム番号 = " & VOLUME_NO.value & " " & vbCrLf & _
'			"where 録画記録番号 = " & intRecLogNo

	strSQL = "SELECT * FROM T_録画記録 WHERE 録画記録番号 = " & intRecLogNo
	DEBUGLOG strSQL

	objRecordset.CursorLocation = adUseClient
	objRecordset.Open strSQL , objConnection, _
	    adOpenStatic, adLockOptimistic
'	Set objRecordset = objConnection.Execute(strSQL)

	objRecordset.MoveFirst

	objRecordset("ディスクＩＤ") = makeDiscID(DVD_NO.value)
	objRecordset("番組名") = BANGUMI_NAME.value
	objRecordset("タイトル名") = TITLE_NAME.value
	objRecordset("ボリューム番号") = VOLUME_NO.value
	objRecordset("録画日付") = REC_DATE.value
	objRecordset("録画時間") = FormatDateTime(REC_TIME.value, vbLongTime) 

	objRecordset("録画モード") = REC_MODE.value
	'//objRecordset("録画回数") = F_REC_COUNT.value
	objRecordset("コピーワンス") = COPYONCE.checked
	objRecordset("削除フラグ") = DELETE_FLAG.checked
	objRecordset("削除日付") = F_DELETE_DATE.value
	objRecordset("編集フラグ") = F_EDIT_FLAG.checked
	objRecordset("メモ") = F_MEMO.value
	objRecordset.Update

	'// トランザクションをコミットする。
	objConnection.CommitTrans()

	Msgbox "更新が完了しました。"

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

	'// トランザクションを開始する。
	objConnection.BeginTrans()

'	strSQL = "DELETE FROM T_録画記録 " & _
'			"WHERE 録画記録番号 = " & intRecLogNo

	strSQL = "SELECT * FROM T_録画記録 WHERE 録画記録番号 = " & intRecLogNo
	DEBUGLOG strSQL

	objRecordset.CursorLocation = adUseClient
	objRecordset.Open strSQL , objConnection, _
	    adOpenStatic, adLockOptimistic

	objRecordset.Delete

	'// トランザクションをコミットする。
	objConnection.CommitTrans()

	Msgbox "削除が完了しました。"

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

	'// トランザクションを開始する。
	objConnection.BeginTrans()

	strSQL = F_SQLCOMMAND2.value

	set objRecordSet = objConnection.Execute(strSQL)

	'// トランザクションをコミットする。
	objConnection.CommitTrans()

	msgbox "正常に登録されました。"

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

	'// トランザクションを開始する。
	objConnection.BeginTrans()

	'// 録画記録テーブルから全データを検索する。
	strSQL = "select * from T_録画記録"

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText
	'// ＤＶＤ番号を編集する。
	strDvdNo = makeDiscID(DVD_NO.value)

	'// 録画日付を編集する。
	strRecDate = REC_DATE_YY.value & "/" & _
				REC_DATE_MM.value & "/" & _
				REC_DATE_DD.value

	'// 録画時間を編集する。
	strRecTime = REC_TIME_HH.value & ":" & _
				REC_TIME_MM_H.value & _
				REC_TIME_MM_L.value & ":" & _
				REC_TIME_SS_H.value & _
				REC_TIME_SS_L.value

	objRecordSet.AddNew
	objRecordSet("ディスクＩＤ") = strDvdNo
	objRecordSet("ボリューム番号") = VOLUME_NO.value
	objRecordSet("番組名") = BANGUMI_NAME.value
	objRecordSet("タイトル名") = TITLE_NAME.value
	objRecordSet("録画日付") = strRecDate
	objRecordSet("録画時間") = strRecTime
	objRecordSet("録画モード") = REC_MODE.value
	objRecordSet("コピーワンス") = COPYONCE.checked
	objRecordSet("削除フラグ") = DELETE_FLAG.checked
	objRecordSet("削除日付") = 0
	objRecordSet("編集フラグ") = F_EDIT_FLAG.checked
	objRecordSet("メモ") = F_DEST_DISCNO.value & vbCrLf & F_MEMO.value
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

	'// トランザクションをコミットする。
	objConnection.CommitTrans()

	msgbox "正常に登録されました。"

	objRecordSet.Close

	objConnection.Close

	set objRecordSet = nothing
	set objConnection = nothing
End Sub
Private Sub cmdSetup()
	'// デバッグ表示オブジェクトを設定する。
	If (isObject(ID_DEBUG_VIEW)) Then
		Set g_objDebugView = ID_DEBUG_VIEW
	End If

	ID_REC_DATE.innerHTML = MakeSelectDate()
	ID_REC_TIME.innerHTML = MakeSelectTime()
	Call MakeItems(ID_KEY_ITEM, "--キー項目--")
	Call MakeItems(ID_ORDERBY, "--分類項目--")
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

	'// 年
	strRecDate = "<select name='REC_DATE_YY'>"
	For i=2000 To 2010
		strDefault = " "
		If (i = strYear) Then
			strDefault = "selected='true' "
		End If

		strRecDate = strRecDate & _
			"<option " & strDefault & "value='" & i & "'>" & i & "</option>" & vbCrLf
	Next
	strRecDate = strRecDate & "</select>年"

	'// 月
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

	'// 日
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

	'// 時
	strRecTime = "<select name='REC_TIME_HH'>"
	For i=0 To 24
		strRecTime = strRecTime & _
			"<option value='" & i & "'>" & i & "</option>" & vbCrLf
	Next
	strRecTime = strRecTime & "</select>時"

	'// 分(上)
	strRecTime = strRecTime & "<select name='REC_TIME_MM_H'>"
	For i=0 To 5
		strRecTime = strRecTime & _
			"<option value='" & i & "'>" & i & "</option>" & vbNewLine
	Next
	strRecTime = strRecTime & "</select>"

	'// 分(下)
	strRecTime = strRecTime & "<select name='REC_TIME_MM_L'>"
	For i=0 To 9
		strRecTime = strRecTime & _
			"<option value='" & i & "'>" & i & "</option>" & vbNewLine
	Next
	strRecTime = strRecTime & "</select>分"

	'// 秒(上)
	strRecTime = strRecTime & "<select name='REC_TIME_SS_H'>"
	For i=0 To 5
		strRecTime = strRecTime & _
			"<option value='" & i & "'>" & i & "</option>" & vbNewLine
	Next
	strRecTime = strRecTime & "</select>"

	'// 秒(下)
	strRecTime = strRecTime & "<select name='REC_TIME_SS_L'>"
	For i=0 To 9
		strRecTime = strRecTime & _
			"<option value='" & i & "'>" & i & "</option>" & vbNewLine
	Next
	strRecTime = strRecTime & "</select>秒"

	MakeSelectTime = strRecTime
End Function
Private Function MakeItems(byref p_objControl, byval p_strTitle)
	Dim objField
	Dim objOption

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	strSQL = "SELECT * FROM T_録画記録 where 録画記録番号 = 1"
	DEBUGLOG strSQL

	objRecordset.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	objRecordset.MoveFirst

	'// ダウンリストのタイトルデータを設定する。
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

	'//-- キー値を設定
	Dim strKeyValue
	strKeyValue = KEYWORD.value
	If (KEY_ITEM.value = "ディスクＩＤ") Then
		strKeyValue = makeDiscID(strKeyValue)
	End If

	'//-- キー条件チェック
	strConditionType = KEY_CONDITION.value

	'//-- キー条件作成
	strWhere = ""
	If (KEY_ITEM.value <> "") Then
		strWhere = _
			KEY_ITEM.value & " " & _
			setKeyCondition( _
				strConditionType, _
				strKeyValue)
	End If

	strSQL = "SELECT * FROM T_録画記録 "
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
		"<th>ディスクＩＤ</th>" & _
		"<th>ボリューム番号</th>" & _
		"<th>番組名</th>" & _
		"<th>タイトル名</th>" & _
		"<th>録画日付</th>" & _
		"<th>録画時間</th>" & _
		"<th>削除フラグ</th>" & _
		"<th>編集フラグ</th>" & _
		"<th>アクション</th>" & _
		"</tr>"
	If (objRecordset.RecordCount > 999) Then
		'//Msgbox "レコード件数が1000件を超えています。" & _
		'//"rec-cnt=" & objRecordset.RecordCount
	End If

	RecCnt=0

	Do Until (objRecordset.EOF)
		strField_1 = objRecordset.Fields.Item("ディスクＩＤ")
		strField_2 = objRecordset.Fields.Item("タイトル名")
		strField_3 = objRecordset.Fields.Item("録画日付")
		strField_4 = objRecordset.Fields.Item("録画時間")
		strField_5 = objRecordset.Fields.Item("ボリューム番号")
		strField_6 = objRecordset.Fields.Item("録画記録番号")
		strField_7 = objRecordset.Fields.Item("削除フラグ")
		strField_8 = objRecordset.Fields.Item("編集フラグ")
		strField_9 = objRecordset.Fields.Item("番組名")
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
					">編集</button><br/>" & "</td>"

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
		case "1"	'// キーに一致する
			strKeyCondition = " = " & _
				strKeyword & " "
		case "2"	'// キーに一致しない
			strKeyCondition = " not = " & _
				strKeyword & " "
		case "3"	'// キーで始まる
			strKeyCondition = " like '" & _
				p_strKeyword & "%' "
		case "4"	'// キーで終わる
			strKeyCondition = " like '%" & _
				p_strKeyword & "' "
		case "5"	'// キーを含む
			strKeyCondition = " like '%" & _
				p_strKeyword & "%' "
		case "6"	'// キーを含まない
			strKeyCondition = " not like '%" & _
				p_strKeyword & "%' "
		case "7"	'// キーより大きい
			strKeyCondition = " > " & _
				strKeyword & " "
		case "8"	'// キーより小さい
			strKeyCondition = " < " & _
				strKeyword & " "
		case "9"	'// キーより以上
			strKeyCondition = " >= " & _
				strKeyword & " "
		case else	'// キーより以下
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

	strSQL = "SELECT * FROM T_録画記録 where 録画記録番号 = 1"

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

	strSQL = "SELECT * FROM T_録画記録 where 録画記録番号 = 1"

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

	'// キー項目
	If (KEY_ITEM.value = "-") Then
		Msgbox("キー項目が選択されていません。")
		Exit Function
	End If

	'// 値
	If (KEYWORD.value = "") Then
		Msgbox("値が入力されていません。")
		Exit Function
	End If

	'// 条件
	If (KEY_CONDITION.value = "-") Then
		'// デフォルト条件（1:一致する）を設定する。
		KEY_CONDITION.value = 1
		'//--Msgbox("条件が選択されていません。")
		'//--Exit Function
	End If

	'// 分類項目
	If (F_ORDERBY.value = "-") Then
		'// デフォルト分類項目（録画記録番号）を設定する。
		F_ORDERBY.value = "録画記録番号"
		'//--Msgbox("分類項目が選択されていません。")
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

	'// 入力条件の妥当性チェック
	If (Not ValidateData()) Then
		Exit Sub
	End If

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	strWQ = Chr(34)

	strConStr = "File Name=default.udl"

	objConnection.Open strConStr, "", ""

	'//-- キー値を設定
	Dim strKeyValue
	strKeyValue = KEYWORD.value
	If (KEY_ITEM.value = "ディスクＩＤ") Then
		strKeyValue = makeDiscID(strKeyValue)
	End If

	'//-- キー条件チェック
	strConditionType = KEY_CONDITION.value

	'//-- キー条件作成
	strWhere = ""
	If (KEY_ITEM.value <> "") Then
		strWhere = _
			KEY_ITEM.value & " " & _
			setKeyCondition( _
				strConditionType, _
				strKeyValue)
	End If

	strSQL = "SELECT * FROM T_録画記録 "
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
		"<th>ディスクＩＤ</th>" & _
		"<th>ボリューム番号</th>" & _
		"<th>番組名</th>" & _
		"<th>タイトル名</th>" & _
		"<th>録画日付</th>" & _
		"<th>録画時間</th>" & _
		"<th>削除フラグ</th>" & _
		"<th>編集フラグ</th>" & _
		"<th>アクション</th>" & _
		"</tr>"
	If (objRecordset.RecordCount > 999) Then
		'//Msgbox "レコード件数が1000件を超えています。" & _
		'//"rec-cnt=" & objRecordset.RecordCount
	End If

	RecCnt=0

	Do Until (objRecordset.EOF)
		strField_1 = objRecordset.Fields.Item("ディスクＩＤ")
		strField_2 = objRecordset.Fields.Item("タイトル名")
		strField_3 = objRecordset.Fields.Item("録画日付")
		strField_4 = objRecordset.Fields.Item("録画時間")
		strField_5 = objRecordset.Fields.Item("ボリューム番号")
		strField_6 = objRecordset.Fields.Item("録画記録番号")
		strField_7 = objRecordset.Fields.Item("削除フラグ")
		strField_8 = objRecordset.Fields.Item("編集フラグ")
		strField_9 = objRecordset.Fields.Item("番組名")
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
					">編集</button><br/>" & "</td>"

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
