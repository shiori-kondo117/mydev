'// File Name	: clsDBI.vbs
'// Description	: ＤＢインタフェースクラスライブラリ
'// Version		: 0.1
'// Author		: Kazuhiro Kondo
Option Explicit

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001

Class clsDBI
	Dim m_objConn
	Dim m_objRs
	Dim m_strConnStr

	Sub u_connect(byval p_strConnStr)
		'// ADODB.Connectionオブジェクトを生成する。
		Set m_objConn = CreateObject("ADODB.Connection")
		'// ADODB.Recordsetオブジェクトを生成する。
		Set m_objRs = CreateObject("ADODB.Recordset")

		'// ＤＢに接続する。
		Call m_objConn.Open(p_strConnStr)
	End Sub

	Function u_execSQLSelect(byval p_strSQL)
		'// ＤＢに接続する。
		Call m_objRs.Open(p_strSQL, _
			m_objConn, _
			adOpenStatic, _
			adLockOptimistic, _
			adCmdText)

		Set u_execSQLSelect = m_objRs
	End Function

	Sub u_disconnect()
		'// レコードセットをクローズする。
		If (isObject(m_objRs)) Then
			Call m_objRs.Close()
		End If

		'// ＤＢを切断する。
		If (isObject(m_objConn)) Then
			Call m_objConn.Close()
		End If
	End Sub
End Class
'// Class Name	: UC_ConntionString
Class clsConnectionString
	Dim m_strConnStr

	Property Get u_ConnectionString()
		u_ConnectionString = m_strConnStr
	End Property

	Sub u_makeConnStrCsvFile(byval p_strFilePath)
		m_strConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	          "Data Source=" & p_strFilePath & ";" & _
	          "Extended Properties=""text;HDR=YES;FMT=Delimited"""
	End Sub
End Class

Class clsRecordset
	Function u_makeHTML(byref p_objRecordset)
		Dim objField
		Dim strHTML

		strHTML = "<table border=""1"" cellspacing=""0"" bgcolor=""white"">" & vbCrLf
		strHTML = strHTML & _
			"<tr style=""background-color:skyblue;"">" & vbCrLf
		For Each objField in p_objRecordset.Fields
			strHTML = strHTML & _
				"<th>" & objField.Name & "</th>" & vbCrLf
		Next
		strHTML = strHTML & "</tr>" & vbCrLf

		Do Until (p_objRecordset.EOF)
			strHTML = strHTML & _
				"<tr>" & vbCrLf

			For Each objField in p_objRecordset.Fields
				strHTML = strHTML & _
					"<td nowrap=""yes"">" & objField.Value & "<br/></td>" & vbCrLf
			Next

			strHTML = strHTML & _
				"</tr>" & vbCrLf

		    p_objRecordset.MoveNext
		Loop

		strHTML = strHTML & "</table>" & vbCrLf

		u_makeHTML = strHTML

		Set objField = Nothing
	End Function
End Class
