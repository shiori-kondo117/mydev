'//---------------------------------------------------------//
'//- ADO(ActiveXDataObject) Procudure/Function for VBScript
'//---------------------------------------------------------//
Option Explicit

'//#########################################################//
'//# �O���[�o���R���X�^���g��`
'//#########################################################//
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001
Const DC_DATABASEPATH = "\\pc-lavie01\home\kondo\�`�k�r�n�j\�f�b�V�X�e��\91.�h�L�������g�Ǘ�\"
Const DC_MDBPATH = "docman.mdb"

'//#########################################################//
'//# �v���Z�[�W���[��`
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

	'// �I�[�v�����R�[�h�Z�b�g
	Call p_objRecset.Open( _
		p_strSQL, _
		p_objConn, _
		adOpenStatic, _
		adLockOptimistic, _
		adCmdText _
	)

End Sub

'//#########################################################//
'//# �t�@���N�V������`
'//#########################################################//

'//*********************************************************//
'//* Procudure  	: CreateADODBConnection
'//* Description 	: ADODB.Connection�I�u�W�F�N�g�𐶐�����B
'//* Arguments 		: 
'//*********************************************************//
Private Function CreateADODBConnection( _
	)
	Set CreateADODBConnection = CreateObject("ADODB.Connection")
End Function

'//*********************************************************//
'//* Procudure  	: CreateADODBRecordset
'//* Description 	: ADODB.Recordset�I�u�W�F�N�g�𐶐�����B
'//* Arguments 		: 
'//*********************************************************//
Private Function CreateADODBRecordset( _
	)
	Set CreateADODBRecordset = CreateObject("ADODB.Recordset")
End Function

'//*********************************************************//
'//* Procudure  	: CreateHtmlTableHeader
'//* Description 	: ���R�[�h�Z�b�g�I�u�W�F�N�g����HTML
'//*                  �e�[�u���w�b�_���𐶐�����B
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
'//* Description 	: ���R�[�h�Z�b�g�I�u�W�F�N�g����HTML
'//*                  �e�[�u���{�f�B���𐶐�����B
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
			'// �J�����l
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
