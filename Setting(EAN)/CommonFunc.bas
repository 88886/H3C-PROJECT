Attribute VB_Name = "CommonFunc"

Public Function ConfirmConfig(ByVal oldArg As String, ByVal newArg As String, ByVal oldArg1 As String, ByVal newArg1 As String) As Boolean
    Dim a
    a = MsgBox("�޸�ǰ<" & oldArg & "�� " & oldArg1 & "> �޸ĺ�<" & newArg & ", " & newArg1 & ">����ȷ���Ƿ��޸�?", vbOKCancel, "ȷ��")
    If (a = VbMsgBoxResult.vbOK) Then
        ConfirmConfig = True
        Exit Function
    End If
    ConfirmConfig = False
End Function
