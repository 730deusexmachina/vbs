Option Explicit
Dim inddObj
Set inddObj = CreateObject("InDesign.Application")
If 0 < inddObj.Documents.Count Then
	If inddObj.ActiveDocument.Saved = True Then
		Dim shellObj
		Set shellObj = CreateObject("WScript.Shell")
		shellObj.Run "cmd /c echo " & inddObj.ActiveDocument.FullName & "|clip", 0, True
		Set shellObj = Nothing
		MsgBox("�N���b�v�{�[�h�ɓ]�����܂����F" & vbCrLf & inddObj.ActiveDocument.FullName)
	End If
End If
Set inddObj = Nothing

