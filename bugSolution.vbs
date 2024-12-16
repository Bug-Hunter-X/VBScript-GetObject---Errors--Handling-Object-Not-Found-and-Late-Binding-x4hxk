Function GetObjectFixed(path)
  On Error Resume Next
  Dim obj
  Set obj = GetObject(path)
  If Err.Number <> 0 Then
    Err.Clear
    ' Handle the error - the object was not found
    MsgBox "Object not found: " & path, vbCritical
    Set obj = Nothing
  End If
  Set GetObjectFixed = obj
End Function

'Example Usage:
Dim myExcelApp
Set myExcelApp = GetObjectFixed("Excel.Application")
If Not myExcelApp Is Nothing Then
    'Work with Excel Application
    myExcelApp.Quit
    Set myExcelApp = Nothing
Else
    MsgBox "Excel is not running", vbExclamation
End If

'For Late Binding, use error handling:
On Error GoTo ErrHandler
Dim objUnknownType
Set objUnknownType = CreateObject("Some.Unknown.Object")
'Work with objUnknownType
Exit Sub
ErrHandler:
    MsgBox "Error creating object: " & Err.Number & ": " & Err.Description, vbCritical
    Err.Clear
End Sub