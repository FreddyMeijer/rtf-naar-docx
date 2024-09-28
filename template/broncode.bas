Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Sub UpdateDocumentFormats()
  Application.ScreenUpdating = False
  
  Dim strFolder As String, 
  strFile As String, 
  wdDoc As Document
  
  strFolder = GetFolder
  
  If strFolder = "" Then Exit Sub
  
  strFile = Dir(strFolder & "\*.rtf", vbNormal)
  While strFile <> ""
    Set wdDoc = Documents.Open(FileName:=strFolder & "\" & strFile, AddToRecentFiles:=False, Visible:=False)
    With wdDoc
      .SaveAs2 FileName:=Left(.FullName, InStrRev(.FullName, ".")) & "docx", Fileformat:=wdFormatXMLDocument, AddToRecentFiles:=False
      .Close wdDoNotSaveChanges
    End With
    strFile = Dir()
  Wend
  
  Set wdDoc = Nothing
  
  Application.ScreenUpdating = True

End Sub

Function GetFolder() As String
  Dim oFolder As Object
  GetFolder = ""
  Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0, "Choose a folder", 0)
  If (Not oFolder Is Nothing) Then GetFolder = oFolder.Items.Item.Path
  Set oFolder = Nothing
End Function


