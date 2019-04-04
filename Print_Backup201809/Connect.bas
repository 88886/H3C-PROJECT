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


Public Function duplicateCheck(sql As String) As Boolean
'   Dim sql As String
   Dim rec As New ADODB.Recordset
   rec.Open sql, conn, adOpenKeyset, adLockReadOnly
   If rec.EOF = False Then
      duplicateCheck = False
   Else
      If rec.Fields(0) = 0 Then
        duplicateCheck = False
      Else
        duplicateCheck = True
      End If
  End If
  On Error GoTo errorHandler
  conn.Execute (sSQLStatement)
  duplicateCheck = True
  Exit Function
errorHandler:
  duplicateCheck = True
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

'Public Sub addPrintedLabel(ByVal barcode As String, ByVal formName As String)
'    Dim rec As New ADODB.Recordset
'    Dim sql As String
'    sql = "insert into printedBarcode(barcode, form_name, creation_time, user_name) " & _
'    "values('" & barcode & "', '" & formName & "', getdate(), '" & golUSERNAME & "') "
'    If conn.State = 0 Then
'        conn.ConnectionString = Connect.getConnectionstring
'        conn.Open
'    End If
'    Status = Connect.excuteUpdate(sql)
'End Sub

'write a public method to get part list by work order
Public Function getPartListByOrder(order As String) As String
    Dim partlist As String, sql As String
    Dim rec As New Recordset
    tempWO = order
    sql = "select b.part_number from [10.11.1.130].afg_active_90.dbo.BOM A,[10.11.1.130].afg_active_90.dbo.BOM_PART_LIST B " & _
        "WHERE A.bom_key = B.bom_key AND a.uda_1 = '" & order & "' and (b.part_number like 'HUV0302%' or b.part_number like 'HWF0302%' )"
        
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    
    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
    If rec.EOF = True Then
        rec.Close
        'sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & order & "' and (assembly like 'HWF0302%' or assembly like 'HUV0302%')"
        sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & order & "' and (assembly like 'HWF0302%' or assembly like 'HUV0302%') and assembly not like '%-SMT%'"
'        sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport   where leading_order ='" & order & "' and (assembly like 'HWF0302%' or assembly like 'HUV0302%') and assembly not like '%-SMT%'"
        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
        If rec.EOF = True Then
            rec.Close
            sql = "select distinct B.Order_Type_S from [10.11.1.130].afg_active_90.dbo.Work_Order A,[10.11.1.130].afg_active_90.dbo.UDA_Order B where A.order_key = B.object_key and A.order_number = '" & tempWO & "'"
            rec.Open sql, conn, adOpenKeyset, adLockReadOnly
            If rec.Fields(0) = "PP05" Then
                getPbByPartList = "Y*"
                taskOrderFlag = True
                rec.Close
                Exit Function
            Else
                rec.Close
                'sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & tempWO & "'"
                sql = "select distinct B.part_number from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A JOIN [10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B ON A.order_key=B.order_key where order_number='" & tempWO & "' and ( B.part_number like 'HWF0231%' or B.part_number like 'HWF0235%'or B.part_number like 'HWF0303%')"
                rec.Open sql, conn, adOpenKeyset, adLockReadOnly
                If rec.EOF = True Then
                    getPbByPartList = "Non"
                    MsgBox "SAP中此工单不包含0302阶或当前阶,不能打印,请确认!"
                    rec.Close
                    Exit Function
                Else
                    Do While Not rec.EOF
                        partlist = partlist + Mid(rec!part_number, 4, 8) + ";"
                        rec.MoveNext
                    Loop
                End If
            End If
        
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
'the main purpose of this function it to get pb status from the database
'modified by allen.yan 2014/07/21 by amor gu requirement
'the return value should be Y3 or Y2

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


'get the  recordset and convert it into array
Public Function GetUnitsArray(order_number As String) As String()
    Dim partlist As String, sql As String
    Dim rec As New Recordset
    
    sql = "select A.part_number from [10.11.1.130].afg_active_90.dbo.UNIT A,[10.11.1.130].afg_active_90.dbo.WORK_ORDER B " & _
        "WHERE A.order_key = B.order_key and b.order_number ='" & order & "'"

        
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    
    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
    If rec.EOF = True Then
         MsgBox "当前工单不存在或者没有释放!"
         'getUnitsArray = ""
         Dim temp() As String
         GetUnitsArray = temp
         
    Else
        arrRows = rec.GetRows()
        intRows = UBound(arrRows)
        
'        Dim info(intRows) As String
'        For intRow = 0 To intRows
'            info(intRow) = arrRows(intRow)
'        Next intRow
        Dim info() As String
        ReDim info(intRows) As String
        For intRow = 0 To intRows
            info(intRow) = arrRows(intRow)
        Next intRow


    End If
    GetUnitsArray = info
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
'    UploadH3CInfo(Pb, Trim(str), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME)
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

'    @sn varchar(32),
'    @hv varchar(8),
'   add by Robin 2018.8.9

Public Function UploadH3C_PB_Version(serial_number As String, hv As String) As Boolean
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
    cmd.CommandText = "upH3C_PB_Version"
    cmd.Parameters.Append cmd.CreateParameter("sn", adVarChar, adParamInput, 32, serial_number)
    cmd.Parameters.Append cmd.CreateParameter("hv", adVarChar, adParamInput, 16, hv)
    cmd.Execute

    Set cmd.ActiveConnection = Nothing
    UploadH3C_PB_Version = True
    Exit Function
errorHandler:
    UploadH3C_PB_Version = False
End Function
'add by carson 20151215
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
'add by carson 20171207

Public Sub addPrintedMACAndSN(ByVal Operator As String, ByVal serial_number As String, ByVal mac As String, ByVal MAC_RECORD_ID As String, ByVal MAC_START As String, ByVal MAC_END As String, ByVal XH As String, ByVal Test As String)
    Dim rec As New ADODB.Recordset
    Dim sql As String
    sql = "update C_MACAndSN_PrintRecord set EFFE_FLAG='0',Last_Modify_Operator='" & golUSERNAME & "',Last_Modify_Time=getdate() where  serial_number='" & serial_number & "'" & _
    "insert into C_MACAndSN_PrintRecord(Operator, serial_number, MAC, MAC_RECORD_ID,MAC_START,MAC_END,SingleUnit_Type,AutoTest) " & _
    "values('" & golUSERNAME & "', N'" & serial_number & "', N'" & mac & "', '" & MAC_RECORD_ID & "', N'" & MAC_START & "', N'" & MAC_END & "', '" & XH & "', '" & Test & "') "
    If connFTPC.State = 0 Then
        connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
        connFTPC.Open
    End If
    connFTPC.Execute (sql)
    connFTPC.Close
End Sub
'add by carson 20171207 end

Public Sub addPrintedLabelMAC(ByVal barcode As String, ByVal mac As String, ByVal formName As String)
    Dim rec As New ADODB.Recordset
    Dim sql As String
    sql = "insert into printedBarcode(barcode, form_name, creation_time, user_name,mac) " & _
    "values('" & barcode & "', '" & formName & "', getdate(), '" & golUSERNAME & "','" & mac & "') "
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Status = Connect.excuteUpdate(sql)
End Sub

'update High_end 2D print program 2014/4/18
'all the sub orders are pb, then Y*
'one or more of sub orders is pb then Y*
'none of sub orders is pub and none of sub orders is not with pb or non pub, then Y2
'if one or more sub orders is without pb setting, then Non.


Public Function getPbByPartList_1(ByVal tempWO As String, ByRef first As String) As String
    Dim sql As String, partlist As String
    Dim rec As New Recordset
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & tempWO & "' and assembly like 'HWF0302%' and assembly not like '%-SMT%'"
    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
    If rec.EOF = True Then
        rec.Close
        sql = "select distinct b.Order_Type_S from [10.11.1.130].afg_active_90.dbo.Work_Order A,[10.11.1.130].afg_active_90.dbo.UDA_Order B where A.order_key = b.object_key and a.order_number = '" & tempWO & "'"
        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
        If rec.Fields(0) = "PP05" Then
            getPbByPartList_1 = "Y*"
            taskOrderFlag = True
            rec.Close
            Exit Function
        Else
            rec.Close
            'sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & tempWO & "' and ( assembly like 'HWF0231%' or assembly like 'HWF0235%')"
            sql = "select distinct B.part_number from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A JOIN [10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B ON A.order_key=B.order_key where order_number='" & tempWO & "' and ( B.part_number like 'HWF0231%' or B.part_number like 'HWF0235%' or B.part_number like 'HWF0212%')"
            rec.Open sql, conn, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                getPbByPartList_1 = "Non"
                MsgBox "SAP中此工单不包含0302阶或整机阶,不能打印,请确认!"
                rec.Close
                Exit Function
            Else
                Do While Not rec.EOF
                    'partlist = partlist + Mid(rec!assembly, 4, 8) + ";"
                    partlist = partlist + Mid(rec!part_number, 4, 8) + ";"
                    rec.MoveNext
                Loop
            End If
        End If
        
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
    first = cmd("first")
    Select Case cmd("res")
        Case "No"
           getPbByPartList_1 = "Y2"
        Case "Non"
            getPbByPartList_1 = "Non"
            Exit Function
        Case "Half"
            getPbByPartList_1 = "N4"
        Case "Yes"
           getPbByPartList_1 = "N4"
    End Select
'    rec.Close
End Function
