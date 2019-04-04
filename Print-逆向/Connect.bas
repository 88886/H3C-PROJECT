Attribute VB_Name = "Connect"
Public conn As New ADODB.Connection
Public connFTPC As New ADODB.Connection
Public golUSERID As String
Public golUSERNAME As String
Public golPath As String
Public info As String
Public nver As String
Public result As String
Public Status As String

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


Public Function AccessCheck(User As String, partition As String) As Boolean
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
    cmd.Parameters.Append cmd.CreateParameter("user", adVarChar, adParamInput, 16, User)
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
    Status = Connect.excuteUpdate(sql)
End Function

Public Function addPrintedLabelMAC(ByVal barcode As String, ByVal formName As String, ByVal mac As String)
    Dim rec As New ADODB.Recordset
    Dim sql As String
    sql = "insert into printedBarcode(barcode, form_name, creation_time, user_name, mac) " & _
    "values('" & barcode & "', '" & formName & "', getdate(), '" & golUSERNAME & "', '" & mac & "') "
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Status = Connect.excuteUpdate(sql)
End Function

Public Function ChangePassword(User As String, oldPass As String, newPass As String) As Boolean
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
    cmd.Parameters.Append cmd.CreateParameter("user", adVarChar, adParamInput, 16, User)
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


Public Function GetPBStatusfromDB(part_number As String) As String
    Dim rec As New ADODB.Recordset
    Dim sql As String
    sql = "select [pb_value] from [H3C_PB_Setting] where part_number = '" & part_number & "' or part_number = 'HWF" & part_number & "'"
    
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
    If rec.EOF = True Then
        MsgBox "当前机种未设定Y1,N1或者Y3,N3,请确认ME相关人员维护!"
        GetPBStatusfromDB = ""
        If rec.State = 1 Then
            rec.Close
        End If
    Else
        GetPBStatusfromDB = rec!pb_value
    End If
End Function

Public Function getPartListByOrder(order As String) As String
    Dim partlist As String, sql As String
    Dim rec As New Recordset
    
    sql = "select b.part_number from [10.11.1.130].afg_active_90.dbo.BOM A,[10.11.1.130].afg_active_90.dbo.BOM_PART_LIST B " & _
        "WHERE A.bom_key = B.bom_key AND a.uda_1 = '" & order & "' and (b.part_number like 'HUV0302%' or b.part_number like 'HWF0302%' )"
        
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    
    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
    If rec.EOF = True Then
        rec.Close
        sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & order & "' and (assembly like 'HWF0302%' or assembly like 'HUV0302%')"
        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
        If rec.EOF = True Then
            MsgBox "SAP中此工单不包含0302阶单板不能打印,请确认!"
'            txtWO.Text = ""
'            txtVer.Text = ""
'            txtWO.SetFocus
            rec.Close
        Else
            Do While Not rec.EOF
                partlist = partlist + Mid(rec!assembly, 4, 8) + ";"
                rec.MoveNext
            Loop
        End If
        getPartListByOrder = partlist
        If rec.State = 1 Then
            rec.Close
        End If
    Else
        Do While Not rec.EOF
            partlist = partlist + Mid(rec!part_number, 4, 8) + ";"
            rec.MoveNext
        Loop
        getPartListByOrder = partlist
    End If
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


