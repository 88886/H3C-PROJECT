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
Public Status As String
Public taskOrderFlag As Boolean

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
    Status = Connect.excuteUpdate(sql)
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

'update High_end 2D print program 2014/4/18
'all the sub orders are pb, then Y*
'one or more of sub orders is pb then Y*
'none of sub orders is pub and none of sub orders is not with pb or non pub, then Y2
'if one or more sub orders is without pb setting, then Non.


Public Function getPbByPartList(ByVal tempWO As String, ByRef first As String) As String
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
            getPbByPartList = "Y*"
            taskOrderFlag = True
            rec.Close
            Exit Function
        Else
            rec.Close
            'sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & tempWO & "' and ( assembly like 'HWF0231%' or assembly like 'HWF0235%')"
            sql = "select distinct B.part_number from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A JOIN [10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B ON A.order_key=B.order_key where order_number='" & tempWO & "' and ( B.part_number like 'HWF0231%' or B.part_number like 'HWF0235%' or B.part_number like 'HWF0212%' or B.part_number like 'HWF0223%')"
            rec.Open sql, conn, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                getPbByPartList = "Non"
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
           getPbByPartList = "Y2"
        Case "Non"
            getPbByPartList = "Non"
            Exit Function
        Case "Half"
            getPbByPartList = "N4"
        Case "Yes"
           getPbByPartList = "N4"
    End Select
'    rec.Close
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
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 16, update_user)
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





