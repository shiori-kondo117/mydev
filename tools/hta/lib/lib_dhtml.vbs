Option Explicit

'//*********************************************************
'//* @procedure createHtmlTable_Basic
'//*********************************************************
Private Sub createHtmlTable_Basic( _
	byref p_objRs, _
	byval p_strFormId, _
	byval p_strId _
	)
	'-- var Object
	Dim objField

	'-- var String
	Dim strSql
	Dim strHtml

	strHtml = _
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"">" & vbCrLf
		
	strHtml = strHtml & _
		"<tr>" & vbCrLf

	For Each objField In p_objRs.Fields
		strHtml = strHtml & _
			"<th>" & objField.Name & "</th>" & vbCrLf
	Next

	strHtml = strHtml & _
		"</tr>" & vbCrLf

	Do Until p_objRs.EOF
		
		strHtml = strHtml & _
			"<tr>" & vbCrLf

		For Each objField In p_objRs.Fields
			If objField.Name = p_strId Then
				strHtml = strHtml & _
					"<td><input type=""checkbox"" name =""f_cbx_ids"" value=""" & _
					p_objRs(p_strId) & """>" & p_objRs(p_strId) & "</input></td>" & vbCrLf
			ElseIf InStr(objField.Name, "dir") _
			Or InStr(objField.Name, "path") Then
				strHtml = strHtml & _
					"<td><a href=""" & objField.Value & """>" & _
					objField.Value & "</a></td>" & vbCrLf
			Else
				strHtml = strHtml & _
					"<td>" & objField.Value & "</td>" & vbCrLf
			End If
		Next

		strHtml = strHtml & _
			"</tr>" & vbCrLf

		p_objRs.MoveNext
	Loop
	
	strHtml = strHtml & _
		"</table>" & vbCrLf

	document.getElementById(p_strFormId).InnerHtml = strHtml

	Set objField = Nothing
End Sub

