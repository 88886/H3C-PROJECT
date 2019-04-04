Attribute VB_Name = "Connect"
Public conn1 As New ADODB.Connection
Public conn As New ADODB.Connection
Public connFTPC As New ADODB.Connection
Public golUSERID As String
Public golUSERNAME As String
Public golPath As String
Public info As String
Public nver As String
Public result As String
Public status As String

Public reprint As Boolean
Public Function isPrintedLabel(ByVal barcode As String, ByVal formName As String) As Boolean
    Dim rec As New ADODB.Recordset
    sql = "select * from printedBarcode where barcode='" & barcode & "' and form_name='" & formName & "'"
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    If rec.EOF = False Then
        isPrintedLabel = True
    Else
        isPrintedLabel = False
    End If
End Function
'
Public Sub addPrintedLabel(ByVal barcode As String, ByVal formName As String)
    Dim rec As New ADODB.Recordset
    Dim sql As String
    sql = "insert into printedBarcode(barcode, form_name, creation_time, user_name) " & _
    "values('" & barcode & "', '" & formName & "', getdate(), '" & golUSERNAME & "') "
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    status = Connect.excuteUpdate(sql)
End Sub

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

Public Function excuteUpdate(sSQLStatement As String) As String
  On Error GoTo errorHandler
  conn.Execute (sSQLStatement)
  excuteUpdate = ""
  Exit Function
errorHandler:
  excuteUpdate = Err.Description
End Function
Public Function checkVersion(argMaintainVer As String, argBeforeVer As String, argCurrentVer As String, argEndDate As String) As Boolean
    Dim beforeVer As String
    Dim currentVer As String
    Dim maintainVer As String
    Dim endDate As Date
    
    maintainVer = Replace(argMaintainVer, vbCrLf, "")
    currentVer = Replace(argCurrentVer, vbCrLf, "")
    beforeVer = Replace(argBeforeVer, vbCrLf, "")
    endDate = CDate(argEndDate)
    
    If UCase(maintainVer) = UCase(currentVer) Then
        checkVersion = True
        Exit Function
    Else
        If Trim(beforeVer) = "" Then
            MsgBox "查询软件版本资料时错误(版本匹配错误)!"
            checkVersion = False
            Exit Function
        End If
        If UCase(maintainVer) = UCase(beforeVer) Then
            If DateDiff("d", Now, endDate) < 0 Then
                MsgBox "查询软件版本资料时错误(超过有效期)!"
                checkVersion = False
                Exit Function
            End If
            checkVersion = True
            Exit Function
        Else
            MsgBox "查询软件版本资料时错误(版本匹配错误)!"
            checkVersion = False
            Exit Function
        End If
    End If
    
End Function










