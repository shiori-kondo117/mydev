'// Version	: 0.1

'// �O���[�o����` ��`
Const DC_CSVFILE_PATH = "D:\MyDatabase\FileInfo"

'// �O���[�o���ϐ� ��`
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
	'// �c�a�C���^�t�F�[�X�N���X�̃C���X�^���X�𐶐�����B
	Set g_objDBI = new clsDBI

	'// �c�a�ڑ�������N���X�̃C���X�^���X�𐶐�����B
	Set g_objConnStr = new clsConnectionString
End Sub

Private Sub connectDB()
	Call g_objConnStr.u_makeConnStrCsvFile(DC_CSVFILE_PATH)

	Call g_objDBI.u_connect(g_objConnStr.u_ConnectionString)
End Sub

