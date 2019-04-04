Attribute VB_Name = "CommonFunc"

Public Function ConfirmConfig(ByVal oldArg As String, ByVal newArg As String, ByVal oldArg1 As String, ByVal newArg1 As String) As Boolean
    Dim a
    a = MsgBox("修改前<" & oldArg & "， " & oldArg1 & "> 修改后<" & newArg & ", " & newArg1 & ">，请确认是否修改?", vbOKCancel, "确认")
    If (a = VbMsgBoxResult.vbOK) Then
        ConfirmConfig = True
        Exit Function
    End If
    ConfirmConfig = False
End Function
