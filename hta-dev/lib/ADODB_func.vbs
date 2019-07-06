Option Explicit
'//--------------------------------------------------------//
'//- グローバル関数 定義
'//--------------------------------------------------------//
'//********************************************************//
'//*	FunctionName	: creADODBConn
'//********************************************************//
Private Function creADODBConn
	Set creADODBConn = CreateObject("ADODB.Connection")
End Function

'//********************************************************//
'//*	FunctionName	: creADODBRecset
'//********************************************************//
Private Function creADODBRecset
	Set creADODBRecset = CreateObject("ADODB.Recordset")
End Function

'//********************************************************//
'//*	FunctionName	: execSQL
'//********************************************************//
Private Function execSQL( _
		byRef p_objConn, _
		byVal p_strSQL _
		)
	Set execSQL = p_objConn.Execute(p_strSQL)
End Function

'//--------------------------------------------------------//
'//- グローバルプロシージャ 定義
'//--------------------------------------------------------//

'//********************************************************//
'//*	FunctionName	: connOracle
'//********************************************************//
Private Sub connOracleDB( _
		byRef p_objConn, _
		byVal p_strSID, _
		byVal p_strUserID, _
		buVal p_strPasswd
		)
	Dim strConnStr

	strConnStr = "Provider=MSDAORA;Data Source=" & _
		p_strSID & ";"

	p_objConn.Open strConnStr, _
		p_strUserID, _
		p_strPasswd

End Sub

'//********************************************************//
'//*	FunctionName	: creHTMLRecset
'//********************************************************//
Private Sub creHTMLRecset( _
		byRef p_objRecset _
		)
	Dim strHTML

	p_objRecset.MoveFirst

	strHTML = "テーブル情報:<br/>" & _
		"<table border='1' cellspacing='0'>"

	strHTML = _
			strHTML & _
			"<tr>" & vbCrLf
	for each field in p_objRecset.Fields
	    	strHTML = _
					strHTML & "<th>" & _
					field.name & _
					"</th>" & vbCrLf
	next

	strHTML = _
			strHTML & _
			"</tr>" & vbCrLf
	
	strData = ""
	Do Until p_objRecset.EOF
		strHTML = strHTML & "<tr>" & vbCrLf
		For Each field in p_objRecset.Fields
			'// カラム値
	    	strHTML = _
					strHTML & "<td>" & _
					field.value & _
					"<br/>" & _
					"</td>" & vbCrLf
	    Next 
		strHTML = strHTML & "</tr>" & vbCrLf
		p_objRecset.MoveNext
	Loop

	strHTML = _
			strHTML & _
			"</table>" & vbCrLf

End Sub

