Option Explicit

'-- CursorTypeEnum
Const adOpenDynamic="2"
Const adOpenForwardOnly="0"
Const adOpenKeyset="1"
Const adOpenStatic="3"
Const adOpenUnspecified="-1"

'-- LockTypeEnum
Const adLockBatchOptimistic="4"
Const adLockOptimistic="3"
Const adLockPessimistic="2"
Const adLockReadOnly="1"
Const adLockUnspecified="-1"

'//*********************************************************
'//* @function getSqlCommand
'//*********************************************************
Private Function getSqlCommand( _
	byval p_intSqlId _
	)
	'-- var Object
	Dim objCn
	Dim objRs

	'-- var String
	Dim strConnStr
	Dim strSql
	
	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	strConnStr = createConnectionString_Csv(g_strHomeDir & "data\conf")

	objCn.Open strConnStr
	
	strSql = "select sql_command from sql_commands.csv where id = " & p_intSqlId

	objRs.Open strSql,objCn,  adOpenStatic, adLockReadOnly

	getSqlCommand = objRs("sql_command")
	
	objRs.Close
	objCn.Close
End Function

'//*********************************************************
'//* @function createConnectionString_Csv
'//*********************************************************
Private Function createConnectionString_Csv( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"Provider=Microsoft.Jet.OLEDB.4.0;" & _
		"Data Source=" & p_strPath & "\;" & _
		"Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
	
	createConnectionString_Csv = strConnStr
End Function

'//*********************************************************
'//* @function createConnectionString_Udl
'//*********************************************************
Private Function createConnectionString_Udl( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"File Name=" & p_strPath & ";"
	createConnectionString_Udl = strConnStr
End Function
