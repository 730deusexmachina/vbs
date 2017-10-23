Set indd = CreateObject("InDesign.Application")
If 0 < indd.Documents.Count Then
	If indd.ActiveDocument.Saved = True Then
		Set shell = CreateObject("WScript.Shell")
		shell.Run "explorer.exe /e,/select,""" & indd.ActiveDocument.FullName & """"
		Set shell = Nothing
	End If
End If
Set indd = Nothing
