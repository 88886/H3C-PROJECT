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


Public Function getPbByPartList(tempWO As String) As String
    Dim sql As String, partlist As String
    Dim rec As New Recordset
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & tempWO & "' and assembly like 'HWF0302%'"
    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
    If rec.EOF = True Then
        MsgBox "SAP中此工单不包含0302阶单板不能打印,请确认!"
        rec.Close
        Exit Function
    Else
        Do While Not rec.EOF
            partlist = partlist + Mid(rec!assembly, 4, 8) + ";"
            rec.MoveNext
        Loop
    End If
    rec.Close
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[PbHandler]"
    cmd.Parameters.Append cmd.CreateParameter("partlist", adVarChar, adParamInput, 8000, partlist)
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 8)
    cmd.Parameters.Append cmd.CreateParameter("first", adVarChar, adParamOutput, 1)
    cmd.Execute
'    tbFirst.Text = cmd("first")
    Select Case cmd("res")
        Case "No"
           getPbByPartList = "Y2"
        Case "Non"
            getPbByPartList = "Non"
            Exit Function
        Case "Half"
            getPbByPartList = "Y*"
        Case "Yes"
           getPbByPartList = "N*"
    End Select
End Function




