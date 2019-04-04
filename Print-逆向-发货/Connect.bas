Attribute VB_Name = "Connect"
Public conn As New ADODB.Connection
Public golUSERID As String
Public golUSERNAME As String
Public golPath As String
Public info As String
Public nver As String
Public result As String
Public Status As String
Public reprint As Boolean

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

'    @sn varchar(32),
'    @hv varchar(8),
'    @5000_status varchar(4),
'    @power_code varchar(32),
'    @power_origin varchar(16),
'    @update_user varchar(16)

Public Function UploadH3CInfo(Pb As String, serial_number As String, hv As String, Status As String, power_code As String, power_origin As String, update_user As String) As Boolean
    On Error GoTo errorHandler
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com As ADODB.Command
    
    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
    'con13.ConnectionTimeout = 50
    con14.Open ConnectionString:=strConn
    Set com = New ADODB.Command
    com.ActiveConnection = con14

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con14
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "upH3CUpload"
    cmd.Parameters.Append cmd.CreateParameter("sn", adVarChar, adParamInput, 32, serial_number)
    cmd.Parameters.Append cmd.CreateParameter("hv", adVarChar, adParamInput, 16, hv)
    cmd.Parameters.Append cmd.CreateParameter("5000_status", adVarChar, adParamInput, 4, Status)
    cmd.Parameters.Append cmd.CreateParameter("power_code", adVarChar, adParamInput, 16, power_code)
    cmd.Parameters.Append cmd.CreateParameter("power_origin", adVarChar, adParamInput, 16, power_origin)
    cmd.Parameters.Append cmd.CreateParameter("pb", adVarChar, adParamInput, 8, Pb)
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 32, update_user)
    cmd.Execute

    Set cmd.ActiveConnection = Nothing
    UploadH3CInfo = True
    Exit Function
errorHandler:
    UploadH3CInfo = False
End Function
Public Function UploadH3CInfo2(pc As Boolean, serial_number As String, hv As String, Status As String, power_code As String, power_origin As String, update_user As String, Pb As String) As Boolean
    On Error GoTo errorHandler
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com As ADODB.Command
    Dim str As String, partList As String

    If Pb = "" Then
        UploadH3CInfo2 = False
        MsgBox "该条码所对应的铅状态无法判断", vbOKOnly + vbExclamation, "该条码所对应的铅状态无法判断"
        Exit Function
    End If
    
    If pc = True And power_code = "" Then
        power_code = Connect.GetResByAction(serial_number, "GetPowerCode")
        If power_code = "" Then
            UploadH3CInfo2 = False
            MsgBox "该条码没有对应电源代码", vbOKOnly + vbExclamation, "该条码所对应的电源代码不存在"
            Exit Function
        End If
    ElseIf pc = False Then
        power_code = "N/A"
    End If
    

    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
    'con13.ConnectionTimeout = 50
    con14.Open ConnectionString:=strConn
    Set com = New ADODB.Command
    com.ActiveConnection = con14

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con14
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "upH3CUpload"
    cmd.Parameters.Append cmd.CreateParameter("sn", adVarChar, adParamInput, 32, serial_number)
    cmd.Parameters.Append cmd.CreateParameter("hv", adVarChar, adParamInput, 100, hv)
    cmd.Parameters.Append cmd.CreateParameter("5000_status", adVarChar, adParamInput, 4, Status)
    cmd.Parameters.Append cmd.CreateParameter("power_code", adVarChar, adParamInput, 16, power_code)
    cmd.Parameters.Append cmd.CreateParameter("power_origin", adVarChar, adParamInput, 16, power_origin)
    cmd.Parameters.Append cmd.CreateParameter("pb", adVarChar, adParamInput, 8, Pb)
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 16, update_user)
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    UploadH3CInfo2 = True
    Exit Function
errorHandler:
    UploadH3CInfo2 = False
End Function

'    @sn varchar(32),
'    @hv varchar(8),
'    @5000_status varchar(4),
'    @power_code varchar(32),
'    @power_origin varchar(16),
'    @update_user varchar(16)
'   data upload for Consume, BB check PB

Public Function UploadH3C_PB(Pb As String, serial_number As String, hv As String, Status As String, power_code As String, power_origin As String, update_user As String) As Boolean
    On Error GoTo errorHandler
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com As ADODB.Command
    
    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
    'con13.ConnectionTimeout = 50
    con14.Open ConnectionString:=strConn
    Set com = New ADODB.Command
    com.ActiveConnection = con14

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con14
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "upH3C_PB"
    cmd.Parameters.Append cmd.CreateParameter("sn", adVarChar, adParamInput, 32, serial_number)
    cmd.Parameters.Append cmd.CreateParameter("hv", adVarChar, adParamInput, 16, hv)
    cmd.Parameters.Append cmd.CreateParameter("5000_status", adVarChar, adParamInput, 4, Status)
    cmd.Parameters.Append cmd.CreateParameter("power_code", adVarChar, adParamInput, 16, power_code)
    cmd.Parameters.Append cmd.CreateParameter("power_origin", adVarChar, adParamInput, 16, power_origin)
    cmd.Parameters.Append cmd.CreateParameter("pb", adVarChar, adParamInput, 8, Pb)
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 32, update_user)
    cmd.Execute

    Set cmd.ActiveConnection = Nothing
    UploadH3C_PB = True
    Exit Function
errorHandler:
    UploadH3C_PB = False
End Function

Public Sub addPrintedLabel(ByVal barcode As String, ByVal formName As String)
    Dim rec As New ADODB.Recordset
    Dim sql As String
    sql = "insert into printedBarcode(barcode, form_name, creation_time, user_name) " & _
    "values('" & barcode & "', '" & formName & "', getdate(), '" & golUSERNAME & "') "
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Status = Connect.excuteUpdate(sql)
End Sub

Public Function isPrintedLabel(ByVal barcode As String, ByVal formName As String) As Boolean
    Dim rec As New ADODB.Recordset
    sql = "select * from printedBarcode where barcode='" & barcode & "' and form_name='" & formName & "'"
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
    If rec.EOF = False Then
        isPrintedLabel = True
    Else
        isPrintedLabel = False
    End If
End Function

Public Function GetPbProperty(sn As String) As String
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com14 As ADODB.Command
    Dim str As String

    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
    
    con14.Open ConnectionString:=strConn
    Set com14 = New ADODB.Command
    com14.ActiveConnection = con14
    str = "select * from afg_b_h3c_upload where serial_number = '" & sn & "'  "
    com14.CommandText = str
    rs14.Open Source:=com14
    If rs14.EOF = True Then
        MsgBox "该SN" & sn & "Pb属性错误,请重新打印该SN的Pb属性！", vbOKOnly + vbExclamation, "Pb属性错误"
        GetPbProperty = ""
        rs14.Close
        Exit Function
        
    Else
        GetPbProperty = rs14.Fields("Pb")
        rs14.Close
        con14.Close
        Exit Function
    End If
End Function
Public Function GetResByAction(serial_number As String, action As String) As String
    On Error GoTo errorHandler
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com As ADODB.Command
    Dim str As String

    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
    'con13.ConnectionTimeout = 50
    con14.Open ConnectionString:=strConn
    Set com = New ADODB.Command
    com.ActiveConnection = con14

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con14
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[dsPrintHandler]"
    cmd.Parameters.Append cmd.CreateParameter("sn", adVarChar, adParamInput, 32, serial_number)
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 16, action)
    cmd.Parameters.Append cmd.CreateParameter("powerCode", adVarChar, adParamOutput, 16)
    cmd.Parameters.Append cmd.CreateParameter("partList", adVarChar, adParamOutput, 64)
    cmd.Parameters.Append cmd.CreateParameter("powerOrigin", adVarChar, adParamOutput, 32)
    cmd.Execute
    Select Case action
        Case "GetPartList"
            GetResByAction = cmd("partList")
        Case "GetPowerCode"
            GetResByAction = cmd("powerCode")
        Case "GetPowerOrigin"
            GetResByAction = cmd("powerOrigin")
    End Select
    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
    GetResByAction = ""
End Function

