Attribute VB_Name = "Connect"
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

Public Function excuteUpdate(sSQLStatement As String) As String
  On Error GoTo errorHandler
  conn.Execute (sSQLStatement)
  excuteUpdate = ""
  Exit Function
errorHandler:
  excuteUpdate = Err.Description
End Function


Public Function AccessCheck(user As String, partition As String) As Boolean
    Dim temp As New Recordset
    Dim sql As String
    On Error GoTo errorHandler
    Dim rs14 As ADODB.Recordset
    Dim com As ADODB.Command
   
    Set rs14 = New ADODB.Recordset
    Set cmd = New ADODB.Command
    If conn.State = 0 And conn.ConnectionString = "" Then
        conn.ConnectionString = Connect.getConnectionstring()
        conn.Open
    End If
    
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[AccessHandler]"
    cmd.Parameters.Append cmd.CreateParameter("user", adVarChar, adParamInput, 16, user)
    cmd.Parameters.Append cmd.CreateParameter("password", adVarChar, adParamInput, 16, "")
    cmd.Parameters.Append cmd.CreateParameter("newpass", adVarChar, adParamInput, 16, "")
    cmd.Parameters.Append cmd.CreateParameter("partition", adVarChar, adParamInput, 16, partition)
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 16, "Validate")
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 32)
    cmd.Execute
        
    If InStr(1, cmd("res"), "OK", vbTextCompare) > 0 Then
        AccessCheck = True
    End If

    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
    AccessCheck = False
End Function

Public Function addPrintedLabel(ByVal barcode As String, ByVal formName As String)
    Dim rec As New ADODB.Recordset
    Dim sql As String
    sql = "insert into printedBarcode(barcode, form_name, creation_time, user_name) " & _
    "values('" & barcode & "', '" & formName & "', getdate(), '" & golUSERNAME & "') "
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    status = Connect.excuteUpdate(sql)
End Function

Public Function ChangePassword(user As String, oldPass As String, newPass As String) As Boolean
    Dim temp As New Recordset
    Dim sql As String
    On Error GoTo errorHandler
    Dim rs14 As ADODB.Recordset
    Dim com As ADODB.Command
   
    Set rs14 = New ADODB.Recordset
    Set cmd = New ADODB.Command
    If conn.State = 0 And conn.ConnectionString = "" Then
        conn.ConnectionString = Connect.getConnectionstring()
        conn.Open
    End If
    
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[AccessHandler]"
    cmd.Parameters.Append cmd.CreateParameter("user", adVarChar, adParamInput, 16, user)
    cmd.Parameters.Append cmd.CreateParameter("password", adVarChar, adParamInput, 16, oldPass)
    cmd.Parameters.Append cmd.CreateParameter("newpass", adVarChar, adParamInput, 16, newPass)
    cmd.Parameters.Append cmd.CreateParameter("partition", adVarChar, adParamInput, 16, "offline")
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 16, "ChangePassword")
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 32)
    cmd.Execute
        
    If InStr(1, cmd("res"), "OK", vbTextCompare) > 0 Then
        ChangePassword = True
    End If

    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
    ChangePassword = False
End Function
