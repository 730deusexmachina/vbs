'�G�N�X�v���[�����J��
Sub execCommand(targetFile)
	Set shell = CreateObject("WScript.Shell")
	shell.Run "explorer.exe /e,/select,""" & targetFile & """"
	Set shell = Nothing
End Sub


'InDesign�I�u�W�F�N�g�쐬
Set indd = CreateObject("InDesign.Application")

'Saved�̃h�L�������g���J���Ă���ꍇ
If 0 < indd.Documents.Count Then
	If indd.ActiveDocument.Saved = True Then

		'app.activeDocument�ɑ΂��ď��������s
		execCommand(indd.ActiveDocument.FullName)
	End If

'Saved�̃h�L�������g���J���Ă��Ȃ����A�u�b�N���J���Ă���ꍇ
ElseIf 0 < indd.Books.Count Then

	'app.activeBook�ɑ΂��ď��������s
	execCommand(indd.ActiveBook.FullName)
End If
Set inddObj = Nothing

'InDesign�I�u�W�F�N�g�j��
Set indd = Nothing


