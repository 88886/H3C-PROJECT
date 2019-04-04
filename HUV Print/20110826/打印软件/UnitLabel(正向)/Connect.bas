Attribute VB_Name = "Connect"
Public conn1 As New ADODB.Connection
Public conn As New ADODB.Connection
Public golUSERID As String
Public golUSERNAME As String
Public golPath As String
Public info As String
Public nver As String
Public result As String
Public status As String

Public Function getConnectionstring() As String
    Dim strLine As String
    Open App.Path + "\Connectionstring.ini" For Input As #1
    Do While EOF(1) = False
    Line Input #1, strLine
    Loop
    Close #1
    getConnectionstring = strLine
End Function

Function chknull(Data1 As Variant, defa As Variant) As Variant
    If IsNull(Data1) Then
        chknull = defa
    Else
        chknull = Trim(Data1)
    End If
End Function

Function getmaxUserID() As String
   Dim sql As String
   Dim rec As New ADODB.Recordset
   sql = "select Max(Userid) from users"
   rec.Open sql, conn, adOpenKeyset, adLockBatchOptimistic
   If rec.EOF = False Then
      getmaxUserID = Val(chknull(rec.Fields(0), "0")) + 1
   Else
     getmaxUserID = "1"
   End If
End Function

Function getmaxID(Form As String) As String
   Dim sql As String
   Dim rec As New ADODB.Recordset
   sql = "select Max(id) from " & Form
   rec.Open sql, conn, adOpenKeyset, adLockBatchOptimistic
   If rec.EOF = False Then
      getmaxID = Val(chknull(rec.Fields(0), "0")) + 1
   Else
     getmaxID = "1"
   End If
End Function

Public Function excuteUpdate(sSQLStatement As String) As String
  On Error GoTo errorHandler
  conn.Execute (sSQLStatement)
  excuteUpdate = ""
  Exit Function
errorHandler:
  excuteUpdate = Err.Description
End Function

Public Sub AdjustColWidth(frmCur As Form, gridCur As Object, Optional bNullRow As Boolean = True, Optional dblIncWidth As Double = 0)
'--------------------------------------------------------------------
'功能:
'       自动调整Grid各列列宽为最合适的宽度
'参数:
'       [frmCur].........................................当前工作窗体
'       [gridCur]........................................当前要调整的Grid
'--------------------------------------------------------------------
Dim i, j As Integer
Dim dblWidth As Double
    With gridCur
        For i = 0 To .Cols - 1
            dblWidth = 0
            If .ColWidth(i) <> 0 Then
                For j = 0 To .Rows - 1
                    If frmCur.TextWidth(.TextMatrix(j, i)) > dblWidth Then
                        dblWidth = frmCur.TextWidth(.TextMatrix(j, i))
                    End If
                Next
                .ColWidth(i) = dblWidth + dblIncWidth + 100
            End If
        Next
    End With
End Sub

