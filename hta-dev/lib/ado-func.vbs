'//---------------------------------------------------------//
'//- ADO(ActiveXDataObject) Procudure/Function for VBScript
'//---------------------------------------------------------//
Option Explicit

'//#########################################################//
'//# グローバルコンスタント定義
'//#########################################################//
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001
Const DC_DATABASEPATH = "\\pc-lavie01\home\kondo\ＡＬＳＯＫ\ＧＣシステム\91.ドキュメント管理\"
Const DC_MDBPATH = "docman.mdb"

'//#########################################################//
'//# プロセージャー定義
'//#########################################################//
'//*********************************************************//
'//* Procudure  	: OpenDBTextFile
'//* Description 	: 
'//* Arguments 		: 
'//*********************************************************//
Private Sub OpenDBTextFile( _
	byRef p_objConn _
	)
	Call p_objConn.Open( _
	  "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=" & _
	  DC_DATABASEPATH & ";" & _
      "Extended Properties=""text;HDR=YES;FMT=Delimited""" _
	)

End Sub

'//*********************************************************//
'//* Procudure  	: OpenDBAccess
'//* Description 	: 
'//* Arguments 		: 
'//*********************************************************//
Private Sub OpenDBAccess( _
	byRef p_objConn _
	)
	Call p_objConn.Open( _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
		DC_DATABASEPATH & _
		DC_MDBPATH & ";" _
	)

End Sub

'//*********************************************************//
'//* Procudure  	: OpenRecordset
'//* Description 	: 
'//* Arguments 		: 
'//*********************************************************//
Private Sub OpenRecordset( _
	byRef p_objConn, _
	byRef p_objRecset, _
	byVal p_strSQL _
	)

	'// オープンレコードセット
	Call p_objRecset.Open( _
		p_strSQL, _
		p_objConn, _
		adOpenStatic, _
		adLockOptimistic, _
		adCmdText _
	)

End Sub

'//#########################################################//
'//# ファンクション定義
'//#########################################################//

'//*********************************************************//
'//* Procudure  	: CreateADODBConnection
'//* Description 	: ADODB.Connectionオブジェクトを生成する。
'//* Arguments 		: 
'//*********************************************************//
Private Function CreateADODBConnection( _
	)
	Set CreateADODBConnection = CreateObject("ADODB.Connection")
End Function

'//*********************************************************//
'//* Procudure  	: CreateADODBRecordset
'//* Description 	: ADODB.Recordsetオブジェクトを生成する。
'//* Arguments 		: 
'//*********************************************************//
Private Function CreateADODBRecordset( _
	)
	Set CreateADODBRecordset = CreateObject("ADODB.Recordset")
End Function

'//*********************************************************//
'//* Procudure  	: CreateHtmlTableHeader
'//* Description 	: レコードセットオブジェクトからHTML
'//*                  テーブルヘッダ部を生成する。
'//* Arguments 		: 
'//*********************************************************//
Private Function CreateHtmlTableHeader( _
	p_objRecset _
	)
	Dim field
	Dim strHTML

	p_objRecset.MoveFirst

	strHTML = "<table border='1' cellspacing='0'>"

	strHTML = _
			strHTML & _
			"<tr class='type_a'>" & vbCrLf
	For Each field in p_objRecset.Fields
	    	strHTML = _
					strHTML & "<th>" & _
					field.name & _
					"</th>" & vbCrLf
	Next

	strHTML = _
			strHTML & _
			"</tr>" & vbCrLf
	
	Set field = Nothing

	CreateHtmlTableHeader = strHTML
End Function

'//*********************************************************//
'//* Procudure  	: CreateHtmlTableBody
'//* Description 	: レコードセットオブジェクトからHTML
'//*                  テーブルボディ部を生成する。
'//* Arguments 		: 
'//*********************************************************//
Private Function CreateHtmlTableBody( _
	p_objRecset _
	)
	Dim field
	Dim strHTML

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

	Set field = Nothing

	CreateHtmlTableBody = strHTML
End Function
