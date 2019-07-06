'// Version	: 0.1

'// グローバル定義 定義
Const DC_CSVFILE_PATH = "D:\MyDatabase\FileInfo"

'// グローバル変数 定義
Public g_objDBI
Public g_objConnStr
Public g_objRecordset

Private Sub window_onload()
	Call initProc()
End Sub

Private Sub initProc()
	Call createObject()
	
	Call connectDB()
End Sub

Private Sub createObject()
	'// ＤＢインタフェースクラスのインスタンスを生成する。
	Set g_objDBI = new clsDBI

	'// ＤＢ接続文字列クラスのインスタンスを生成する。
	Set g_objConnStr = new clsConnectionString
End Sub

Private Sub connectDB()
	Call g_objConnStr.u_makeConnStrCsvFile(DC_CSVFILE_PATH)

	Call g_objDBI.u_connect(g_objConnStr.u_ConnectionString)
End Sub

