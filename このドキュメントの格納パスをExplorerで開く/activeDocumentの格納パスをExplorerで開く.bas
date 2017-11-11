Attribute VB_Name = "id_openExplorer"
Option Explicit
Public Sub activeDocument‚ÌŠi”[ƒpƒX‚ðExplorer‚ÅŠJ‚­()
    Dim inddObj
    Set inddObj = CreateObject("InDesign.Application")
    If 0 < inddObj.Documents.Count Then
        If inddObj.ActiveDocument.Saved = True Then
            Dim shellObj
            Set shellObj = CreateObject("WScript.Shell")
            shellObj.Run "explorer.exe /e,/select,""" & inddObj.ActiveDocument.FullName & """"
            Set shellObj = Nothing
        End If
    End If
    Set inddObj = Nothing
End Sub

