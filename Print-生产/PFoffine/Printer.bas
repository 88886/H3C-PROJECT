Attribute VB_Name = "Printer"
Public Function PrintLabel(ByVal File As String, ByRef Dic As dictionary) As Boolean
    'Dim fs As New Scripting.FileSystemObject
    'If (fs.FileExists(File) = False) Then
    '    PrintLabel = False
    '    Exit Function
    'End If
    
    'If (InStr(1, LCase(File), ".btw", vbTextCompare) > -1) Then
    '    PrintLabel = BartenderPrint(File, Dic)
    '    Exit Function
    'End If
    
End Function

Private Function BartenderPrint(ByVal File As String, ByRef Dic As dictionary) As Boolean
    'Dim btApp As BarTender.Application
    'Dim btFormat As BarTender.Format
    'Dim btNamesSubStrings As BarTender.NamedSubStrings
    
    'Dim value As String, name As String
    
    'Set btApp = CreateObject("BarTender.Application")
    'Set btFormat = btApp.Formats.Open(File, True)
    'Set btNamesSubStrings = btFormat.NamedSubStrings
    'For i = 1 To btNamesSubStrings.count
    '    name = btNamesSubStrings.Item(i).name
    '    value = Dic.GetValue(name)
    '    If (value = "Null") Then
            
    '    Else
    '        btNamesSubStrings.Item(i).value = value
     '   End If
    'Next
    'Call btFormat.PrintOut(False, False)
    'btFormat.Close
    'btApp.Quit
    'Set btApp = Nothing
End Function
