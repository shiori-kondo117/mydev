Option Explicit

'//*********************************************************
'//* @procedure cmdExecSqlCommand_Click
'//*********************************************************
Private Sub cmdExecSqlCommand_Click( _
	byval p_SqlCmdType, _
	byval p_strSqlCommand _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath

	'-- var Integer
	Dim intUpdateNum
	
	If p_SqlCmdType <> "query" _
	And p_SqlCmdType <> "execute" Then
		Msgbox "sql command type illigal[{query|execute}]",vbCritical,"SQLコマンド種類エラー"
		Exit Sub
	End If
	
	Select Case f_ddl_connstr_types.value
		Case "udl"
			strConnStr = createConnectionString_Udl("../conf/csv.udl")
		Case "csv"
			If f_ddl_file_paths.value <> "-" Then
				If InStr(f_ddl_file_paths.value,"\") Then
					strPath = f_ddl_file_paths.value
				Else
					strPath = g_strHomeDir &  "data\" & f_ddl_file_paths.value
				End If
			Else
				strPath = g_strHomeDir & "data"
			End If
			strConnStr = createConnectionString_Csv(strPath)
		Case Else
			Msgbox "Illigal db connection types!!",vbCritical,"選択エラー"
			Exit Sub
	End Select
	
	If p_strSqlCommand = "" Then
		Msgbox "sql command nothing!!!",vbCritical,"SQLコマンド未入力エラー"
		Exit Sub
	End If

	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	objCn.Open strConnStr
	
	strSql = p_strSqlCommand
	f_txa_sql_command.value = strSql

	If p_SqlCmdType = "query" Then
		objRs.Open strSql,objCn,  adOpenStatic, adLockReadOnly
		'-- create HTML TABLE
		createHtmlTable_Basic objRs, "id_view_result", "id"
	Else
		objCn.Execute strSql,intUpdateNum
		If intUpdateNum = 0 Then Msgbox "command failer!!":Exit Sub
		Exit Sub
	End If
End Sub

'//*********************************************************
'//* @procedure cmdActionCommand_Click
'//*********************************************************
Private Sub cmdActionCommand_Click( _
	byref p_objForm _
	)
	'-- var String
	Dim strCmd
	
	'-- enable action-events
	If p_objForm.value = "-" Then
		Exit Sub
'--		Msgbox "not select action-events!!",vbCritical,"選択エラー":Exit Sub
	End If
	
	strCmd = FUNC_EVENT_PREFIX & p_objForm.value
	
	eval(strCmd)
	
	cmdFunctionCommand_Click f_ddl_functions
End Sub

'//*********************************************************
'//* @function evtActionCommand_SuccessCopy
'//*********************************************************
Private Function evtActionCommand_SuccessCopy
	'-- var Object
	Dim objId

	'-- var String
	Dim strSql
	
	For Each objId In f_cbx_ids
		If objId.Checked Then
			strSql = "insert into vr_copy_log_statuses.csv values(" & objId.value & ",now);"
			cmdExecSqlCommand_Click "execute",strSql
		End If
	Next

	Msgbox "command successfully!!",vbInformation
	
	Set objId = Nothing
End Function

'//*********************************************************
'//* @function evtActionCommand_SuccessEdit
'//*********************************************************
Private Function evtActionCommand_SuccessEdit
	'-- var Object
	Dim objId

	'-- var String
	Dim strSql
	
	For Each objId In f_cbx_ids
		If objId.Checked Then
			strSql = "insert into vr_edit_log_statuses.csv values(" & objId.value & ",now);"
			cmdExecSqlCommand_Click "execute",strSql
		End If
	Next
	
	Msgbox "command successfully!!",vbInformation
	
	Set objId = Nothing
End Function

'//*********************************************************
'//* @procedure cmdFunctionCommand_Click [arg1]
'//* @arg1 form-functions
'//*********************************************************
Private Sub cmdFunctionCommand_Click( _
	byref p_objForm _
	)
	'-- var String
	Dim strSql
	
	'-- enable action-events
	If p_objForm.value = "-" Then
		Exit Sub
	End If
	
	strSql = getSqlCommand(p_objForm.value)
	
	cmdExecSqlCommand_Click "query", strSql
End Sub

