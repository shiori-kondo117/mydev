Option Explicit

'//*********************************************************
'//* @procedure cmdExecSqlCommand_Click
'//*********************************************************
Private Sub cmdExecSqlCommand_Click
	'-- var Object
	Dim objCn
	Dim objRs
	Dim objField
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath
	Dim strHtml
	
	If f_ddl_connstr_types.value <> "udl" Then
		Msgbox "not select conn types",vbCritical,"ëIëÉGÉâÅ["
		Exit Sub
	End If
	
	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	strConnStr = createConnectionString_Udl("../conf/csv.udl")
	
	objCn.Open strConnStr
	
	strSql = f_txa_sql_command.value

	objRs.Open strSql,objCn,  adOpenStatic, adLockReadOnly
	
	strHtml = _
		"<table border=""3"" cellpadding=""3"" cellspacing=""3"">" & vbCrLf

	strHtml = strHtml & _
		"<tr>" & vbCrLf
		
	For Each objField In objRs.Fields
		strHtml = strHtml & _
			"<th>" & objField.Name & "</th>" & vbCrLf
	Next

	strHtml = strHtml & _
		"</tr>" & vbCrLf
		
	Do Until objRs.EOF
		strHtml = strHtml & _
			"<tr>" & vbCrLf
			
		For Each objField In objRs.Fields
			If InStr(objField.Name,"path") Then
				strHtml = strHtml & _
					"<td><a href=""" & objField.Value & """>" & objField.Value & "</a></td>" & vbCrLf
			Else
				strHtml = strHtml & _
					"<td>" & objField.Value & "</td>" & vbCrLf
			End If
		Next

		strHtml = strHtml & _
			"</tr>" & vbCrLf
			
		objRs.MoveNext
	Loop

	strHtml = strHtml & _
		"</table>" & vbCrLf
	
	document.getElementById("id_view_result").innerHTML = strHtml
End Sub
