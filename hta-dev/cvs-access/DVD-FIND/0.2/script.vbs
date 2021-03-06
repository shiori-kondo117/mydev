'//*********************************************************
'//* File Name	: script.vbs
'//* Description: ＤＶＤ／ＣＤのファイルを検索する。
'//* Author		: Kazuhiro Kondo
'//* Version	: 0.2
'//*********************************************************
On Error Resume Next

Private Sub cmdSearch_Click()
	Dim objHTML

	'// キー項目選択チェック
	If (KEY_ITEM.value = "-") Then
		Msgbox("キー項目を選択して下さい。")
		Exit Sub
	End If

	'// 分類項目選択チェック
	If (ORDER_BY.value = "-") Then
		'// デフォルト項目を設定する。
		ORDER_BY.value = "ファイル名"
	End If

	'// レコードセットクラスのインスタンスを生成する。
	Set objHTML = new clsRecordset

	'//-- キー条件チェック
	strConditionType = KEY_CONDITION.value

	strWhere = _
		KEY_ITEM.value & " " & _
		setKeyCondition( _
			strConditionType, _
			KEYWORD.value)

	strSQL = "SELECT " & _
				"ボリューム名, " & _
				"ファイル名, " & _
				"パス, " & _
				"日付 " & _
			"FROM DVD_info_db.txt "
	If (strWhere <> "") Then
		strSQL = strSQL & _
			"where " & _
			strWhere
	End If

	strSQL = strSQL & _
			"order by " & ORDER_BY.value
	
	Call DEBUGLOG(strSQL)

	Set g_objRecordSet = g_objDBI.u_execSQLSelect(strSQL)

	RESULT.innerHTML = objHTML.u_makeHTML(g_objRecordSet)

	Call g_objDBI.u_disconnect()

	Set g_objRecordSet = Nothing
End Sub

Private Function setKeyCondition( _
		byVal p_strCondType, _
		byVal p_strKeyword _
		)
'//	strKeyword = UCase(p_strKeyword)
	strKeyword = p_strKeyword
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
		case else	'// キーを含まない
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
	Set g_objRecordSet = CreateObject("ADODB.Recordset")

	Call OpenDbTextFile(objConnection, DC_CSVFILE_PATH)

	strSQL = "SELECT * FROM DVD_info_db.txt " & _
				"WHERE ボリューム名 = 'DVD-0601'"
	Call DEBUGLOG(strSQL)

	g_objRecordSet.Open strSQL, _
	          objConnection, adOpenStatic, adLockOptimistic, adCmdText

	g_objRecordSet.MoveFirst

	'// ダウンリストのタイトルデータを設定する。
	Set objOption = Document.createElement("OPTION")
	objOption.Text = p_strTitle
    objOption.Value = "-"
    p_objControl.Add(objOption)

	For Each objField in g_objRecordSet.Fields
		Set objOption = Document.createElement("OPTION")
		objOption.Text = objField.Name
        objOption.Value = objField.Name
        p_objControl.Add(objOption)
	Next

	Set objOption = Nothing
	Set objField = Nothing
	Set g_objRecordSet = Nothing
	Set objConnection = Nothing
End Function

Private Sub cmdSetup()
	Call MakeItems(ID_KEY_ITEM, "--キー項目--")
	Call MakeItems(ID_ORDER_BY, "--分類項目--")
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
