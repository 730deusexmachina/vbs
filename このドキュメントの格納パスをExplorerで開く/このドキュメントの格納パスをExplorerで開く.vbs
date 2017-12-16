'エクスプローラを開く
Sub execCommand(targetFile)
	Set shell = CreateObject("WScript.Shell")
	shell.Run "explorer.exe /e,/select,""" & targetFile & """"
	Set shell = Nothing
End Sub


'InDesignオブジェクト作成
Set indd = CreateObject("InDesign.Application")

'Savedのドキュメントが開いている場合
If 0 < indd.Documents.Count Then
	If indd.ActiveDocument.Saved = True Then

		'app.activeDocumentに対して処理を実行
		execCommand(indd.ActiveDocument.FullName)
	End If

'Savedのドキュメントが開いていないが、ブックが開いている場合
ElseIf 0 < indd.Books.Count Then

	'app.activeBookに対して処理を実行
	execCommand(indd.ActiveBook.FullName)
End If
Set inddObj = Nothing

'InDesignオブジェクト破棄
Set indd = Nothing


