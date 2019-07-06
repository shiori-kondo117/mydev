Private Sub OpenDbTextFile(ByRef p_objConnection, ByVal p_TextFilePath)
	p_objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	          "Data Source=" & p_TextFilePath & ";" & _
	          "Extended Properties=""text;HDR=YES;FMT=Delimited"""
End Sub

