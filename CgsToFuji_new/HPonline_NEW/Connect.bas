Attribute VB_Name = "Connect"
Public conn1 As New ADODB.Connection
Public conn11 As New ADODB.Connection
Public conn As New ADODB.Connection
Public golUSERID As String
Public golUSERNAME As String
Public golPath As String
Public info As String
Public nver As String
Public result As String
Public Status As String
Public golWorkOrder As String

Public reprint As Boolean

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
Public Function isPrintedHKLabel(ByVal barcode As String) As Boolean
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com14 As ADODB.Command
    Dim str As String
    

    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
    'con13.ConnectionTimeout = 50
    con14.Open ConnectionString:=strConn
    Set com14 = New ADODB.Command
    com14.ActiveConnection = con14
    str = "select 1 from C_HK_Ship_Print_Rec WHERE EFFE_FLAG='1' AND H3C_Serial_Number ='" & barcode & "' "
    com14.CommandText = str
       'rs14.Open str
    rs14.Open Source:=com14
    If rs14.EOF = False Then
        isPrintedHKLabel = True
    Else
        isPrintedHKLabel = False
    End If
    rs14.Close
    con14.Close
    
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

Public Function checkWeighInformation(part_number As String, part_reversion As String) As Boolean
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com14 As ADODB.Command
    Dim str As String
    

    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
    'con13.ConnectionTimeout = 50
    con14.Open ConnectionString:=strConn
    Set com14 = New ADODB.Command
    com14.ActiveConnection = con14
    'str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtHPSN.Text) & "'"
      str = "select 1 from [H3C_HPWeight] " & _
        " where ((Part_Number = '" & part_number & "' and part_revision = '" & part_reversion & "') or " & _
        " (Part_Number = '" & part_number & "' and part_revision = 'ALL')) " & _
        " and GrossWeight is not null and GrossWeight <> '' and is_Valid = 1 "
    com14.CommandText = str
       'rs14.Open str
    rs14.Open Source:=com14
    If rs14.EOF = True Then
        MsgBox "没有维护重量", vbOKOnly + vbExclamation, "警告"
        checkWeighInformation = False
        Exit Function
        rs14.Close
    Else
        rs14.Close
        con14.Close
        checkWeighInformation = True
    End If
End Function

Public Function getPowerCode(serial_number As String) As String
    Dim power_code As String
    power_code = Connect.GetResByAction(serial_number, "GetPowerCode")
    If power_code = "" Then
        getPowerCode = ""
        MsgBox "该条码没有对应电源代码", vbOKOnly + vbExclamation, "该条码所对应的电源代码不存在"
        Exit Function
    Else
        getPowerCode = power_code
        Exit Function
    End If
End Function

'    @sn varchar(32),
'    @hv varchar(8),
'    @5000_status varchar(4),
'    @power_code varchar(32),
'    @power_origin varchar(16),
'    @update_user varchar(16)

Public Function UploadH3CInfo(pc As Boolean, serial_number As String, hv As String, Status As String, power_code As String, power_origin As String, update_user As String) As Boolean
    On Error GoTo errorHandler
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com As ADODB.Command
    Dim str As String, PB As String, partList As String
    
    partList = Connect.GetResByAction(serial_number, "GetPartList")
    If partList = "" And pc = True Then
        PB = "Pb"
    ElseIf partList = "" And pc = False Then
        UploadH3CInfo = False
        MsgBox "该条码所对应的BOM中没有0302的下阶", vbOKOnly + vbExclamation, "BOM中没有0302的下阶"
        Exit Function
    End If
   
    
    If PB = "" Then
        PB = Connect.GetPBState(partList)
        If PB = "" Then
            UploadH3CInfo = False
            MsgBox "该条码所对应的铅状态无法判断", vbOKOnly + vbExclamation, "该条码所对应的铅状态无法判断"
            Exit Function
        End If
    End If
    
    
    If pc = True And power_code = "" Then
        power_code = Connect.GetResByAction(serial_number, "GetPowerCode")
        If power_code = "" Then
            UploadH3CInfo = False
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
    cmd.Parameters.Append cmd.CreateParameter("hv", adVarChar, adParamInput, 16, hv)
    cmd.Parameters.Append cmd.CreateParameter("5000_status", adVarChar, adParamInput, 4, Status)
    cmd.Parameters.Append cmd.CreateParameter("power_code", adVarChar, adParamInput, 16, power_code)
    cmd.Parameters.Append cmd.CreateParameter("power_origin", adVarChar, adParamInput, 16, power_origin)
    cmd.Parameters.Append cmd.CreateParameter("pb", adVarChar, adParamInput, 8, PB)
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 16, update_user)
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    UploadH3CInfo = True
    Exit Function
errorHandler:
    UploadH3CInfo = False
End Function


Public Function UploadH3CInfo2(pc As Boolean, serial_number As String, hv As String, Status As String, power_code As String, power_origin As String, update_user As String, PB As String) As Boolean
    On Error GoTo errorHandler
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com As ADODB.Command
    Dim str As String, partList As String

    If PB = "" Then
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
    cmd.Parameters.Append cmd.CreateParameter("pb", adVarChar, adParamInput, 8, PB)
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 16, update_user)
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    UploadH3CInfo2 = True
    Exit Function
errorHandler:
    UploadH3CInfo2 = False
End Function
Public Function UploadHKShip_Rec(H3C_Serial_Number As String, H3C_Part As String, H3C_Rev As String, HK_Part As String, HK_Serial_Number As String, HK_DESC As String, HK_TYPE As String, Material As String, SalesOrder As String, H3C_SN_PB As String, CCC As String, update_user As String) As Boolean
    On Error GoTo errorHandler
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com As ADODB.Command
    Dim str As String, partList As String

'    If PB = "" Then
'        UploadH3CInfo2 = False
'        MsgBox "该条码所对应的铅状态无法判断", vbOKOnly + vbExclamation, "该条码所对应的铅状态无法判断"
'        Exit Function
'    End If
'
'    If pc = True And power_code = "" Then
'        power_code = Connect.GetResByAction(Serial_Number, "GetPowerCode")
'        If power_code = "" Then
'            UploadH3CInfo2 = False
'            MsgBox "该条码没有对应电源代码", vbOKOnly + vbExclamation, "该条码所对应的电源代码不存在"
'            Exit Function
'        End If
'    ElseIf pc = False Then
'        power_code = "N/A"
'    End If
    

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
    cmd.CommandText = "upHK_Ship_Rec"
    cmd.Parameters.Append cmd.CreateParameter("H3C_Serial_Number", adVarChar, adParamInput, 100, H3C_Serial_Number)
    cmd.Parameters.Append cmd.CreateParameter("H3C_Part", adVarChar, adParamInput, 100, H3C_Part)
    cmd.Parameters.Append cmd.CreateParameter("H3C_Rev", adVarChar, adParamInput, 100, H3C_Rev)
    cmd.Parameters.Append cmd.CreateParameter("HK_Part ", adVarChar, adParamInput, 100, HK_Part)
    cmd.Parameters.Append cmd.CreateParameter("HK_Serial_Number", adVarChar, adParamInput, 100, HK_Serial_Number)
    cmd.Parameters.Append cmd.CreateParameter("HK_DESC", adVarChar, adParamInput, 100, HK_DESC)
    cmd.Parameters.Append cmd.CreateParameter("HK_TYPE", adVarChar, adParamInput, 100, HK_TYPE)
    cmd.Parameters.Append cmd.CreateParameter("Material", adVarChar, adParamInput, 100, Material)
    If SalesOrder = "" Then
        cmd.Parameters.Append cmd.CreateParameter("Sales_Order", adVarChar, adParamInput, 100, "")
    Else
        cmd.Parameters.Append cmd.CreateParameter("Sales_Order", adVarChar, adParamInput, 100, SalesOrder)
    End If
    cmd.Parameters.Append cmd.CreateParameter("H3C_SN_PB", adVarChar, adParamInput, 10, H3C_SN_PB)
    cmd.Parameters.Append cmd.CreateParameter("CCC", adVarChar, adParamInput, 10, CCC)
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 16, update_user)
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    UploadHKShip_Rec = True
    Exit Function
errorHandler:
    UploadHKShip_Rec = False
End Function
Public Function UploadConsen_Rec(serial_number As String, order_number As String, part_number As String, Product_Desc As String, Product_Model As String, Product_Material As String, Product_Power As String, SalesOrder As String, CCC As String, ChinaRoHS As String, WEEE As String, Laser As String, update_user As String) As Boolean
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
    cmd.CommandText = "upConsenRec"
    cmd.Parameters.Append cmd.CreateParameter("Serial_Number", adVarChar, adParamInput, 100, serial_number)
    cmd.Parameters.Append cmd.CreateParameter("Order_Number", adVarChar, adParamInput, 100, order_number)
    cmd.Parameters.Append cmd.CreateParameter("Part_Number", adVarChar, adParamInput, 100, part_number)
    cmd.Parameters.Append cmd.CreateParameter("Product_Desc", adVarChar, adParamInput, 100, Product_Desc)
    cmd.Parameters.Append cmd.CreateParameter("Product_Model", adVarChar, adParamInput, 100, Product_Model)
    cmd.Parameters.Append cmd.CreateParameter("Product_Material", adVarChar, adParamInput, 100, Product_Material)
    cmd.Parameters.Append cmd.CreateParameter("Product_Power", adVarChar, adParamInput, 100, Product_Power)
    cmd.Parameters.Append cmd.CreateParameter("SalesOrder", adVarChar, adParamInput, 100, SalesOrder)
    cmd.Parameters.Append cmd.CreateParameter("CCC", adVarChar, adParamInput, 10, CCC)
    cmd.Parameters.Append cmd.CreateParameter("ChinaRoHS", adVarChar, adParamInput, 10, ChinaRoHS)
    cmd.Parameters.Append cmd.CreateParameter("WEEE", adVarChar, adParamInput, 10, WEEE)
    cmd.Parameters.Append cmd.CreateParameter("Laser", adVarChar, adParamInput, 10, Laser)
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 32, update_user)
    cmd.Execute

    Set cmd.ActiveConnection = Nothing
    UploadConsen_Rec = True
    Exit Function
errorHandler:
    UploadConsen_Rec = False
End Function

'    @sn varchar(32),
'    @hv varchar(8),
'    @5000_status varchar(4),
'    @power_code varchar(32),
'    @power_origin varchar(16),
'    @update_user varchar(16)
'   data upload for Consume, BB check PB

Public Function UploadH3C_PB(PB As String, serial_number As String, hv As String, Status As String, power_code As String, power_origin As String, update_user As String) As Boolean
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
    cmd.Parameters.Append cmd.CreateParameter("pb", adVarChar, adParamInput, 8, PB)
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 16, update_user)
    cmd.Execute

    Set cmd.ActiveConnection = Nothing
    UploadH3C_PB = True
    Exit Function
errorHandler:
    UploadH3C_PB = False
End Function


'main purpose is to judge the
Public Function IsTaskOrder(work_order As String) As Boolean
    Set conn = New ADODB.Connection
    Set rec = New ADODB.Recordset
    strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
    'con13.ConnectionTimeout = 50
    conn.Open ConnectionString:=strConn
    sql = "select distinct b.Order_Type_S from [10.11.1.130].afg_active_90.dbo.Work_Order A,[10.11.1.130].afg_active_90.dbo.UDA_Order B where A.order_key = b.object_key and a.order_number = '" & tempWO & "'"
    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
    If rec.Fields(0) = "PP05" Then
        IsTaskOrder = True
        rec.Close
        Exit Function
    Else
        IsTaskOrder = False
        rec.Close
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

Public Function getPartListByOrder(order As String) As String
    Dim partList As String, sql As String
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
                partList = partList + Mid(rec!assembly, 4, 8) + ";"
                rec.MoveNext
            Loop
        End If
        getPartListByOrder = partList
        If rec.State = 1 Then
            rec.Close
        End If
    Else
        Do While Not rec.EOF
            partList = partList + Mid(rec!part_number, 4, 8) + ";"
            rec.MoveNext
        Loop
        getPartListByOrder = partList
    End If
End Function

Public Function GetPBState(partList As String) As String
    On Error GoTo errorHandler
    Dim cmd As New ADODB.Command
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
    End If
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[PbHandler]"
    cmd.Parameters.Append cmd.CreateParameter("partlist", adVarChar, adParamInput, 8000, partList)
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 8)
    cmd.Parameters.Append cmd.CreateParameter("first", adVarChar, adParamOutput, 1)
    cmd.Execute
    Select Case cmd("res")
        Case "No"
            GetPBState = "NPb"
        Case "Non"
            MsgBox "此工单包含0302阶单板未设定有铅无铅,请相关ME去设定!"
            GetPBState = ""
            Exit Function
        Case "Half"
             GetPBState = "Pb"
        Case "Yes"
             GetPBState = "Pb"
    End Select
    Exit Function
errorHandler:
    GetPBState = ""
End Function

Public Function GetUploadInfo(model As String, hv As String, project As String) As String
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com14 As ADODB.Command
    Dim str As String

    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    strConn = Connect.getConnectionstring
    'con13.ConnectionTimeout = 50
    con14.Open ConnectionString:=strConn
    Set com14 = New ADODB.Command
    com14.ActiveConnection = con14
    If project = "NEC" Or project = "3COM" Then
        str = "select case when Print_SV = 1 then 'Y' else 'N' end,case when Print_Power = 1 then 'Y' else 'N' end,[5000_State] from tblOthers where Part_Number = '" & model & "' and Part_Revision = '" & hv & "'"
    End If
    com14.CommandText = str
    rs14.Open Source:=com14
    If rs14.EOF = True Then
        MsgBox "没有维护该机种" & model & "在Setting中信息", vbOKOnly + vbExclamation, "请在Setting维护该机种信息"
        GetUploadInfo = ""
        rs14.Close
        Exit Function
        
    Else
        GetUploadInfo = rs14.Fields(0) + ";" + rs14.Fields(1) + ";" + rs14.Fields(2)
        rs14.Close
        con14.Close
        Exit Function
    End If
End Function

'this function is to check whether current version is valid in the ECO control system

Public Function IsValidECOVersion(model As String, part_revision As String) As Boolean
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com14 As ADODB.Command
    Dim str As String

    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    strConn = Connect.getConnectionstring
    'con13.ConnectionTimeout = 50
    con14.Open ConnectionString:=strConn
    Set com14 = New ADODB.Command
    com14.ActiveConnection = con14
    str = "select * from tblECO_Ver where PartNumber = '" & model & "' and Version = '" & part_revision & "' and Active = 0 "
    com14.CommandText = str
    rs14.Open Source:=com14
    If rs14.EOF = False Then
        MsgBox "该机种" & model & "版本:" & part_revision & "在条码ECO防呆维护为无效,不能打印！", vbOKOnly + vbExclamation, "在Setting中条码ECO防呆维护设定为无效"
        IsValidECOVersion = False
        rs14.Close
        Exit Function
        
    Else
        IsValidECOVersion = True
        rs14.Close
        con14.Close
        Exit Function
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

Public Function GetMacFromTestRecord(sn As String, equipment As String) As String

    Dim con14, conSQL01_1 As ADODB.Connection
    Dim rs14, rsSQL01_1 As ADODB.Recordset
    Dim com14, comSQL01_1 As ADODB.Command
    Dim str As String
    
    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    Set conSQL01_1 = New ADODB.Connection
    Set rsSQL01_1 = New ADODB.Recordset

    strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.31;Initial Catalog=dataT; User ID=sa; PWD=Itadmin1"
    
    con14.Open ConnectionString:=strConn
    Set com14 = New ADODB.Command
    com14.ActiveConnection = con14
    If equipment = "FT" Then 'equipment like '%FT'
        str = "select top 1 MAC from [test_equ] " & _
                    " where barcode = '" & Trim(sn) & "' and pass = N'通过' and len(mac)=12 and  RIGHT(RTRIM(equipment),2)='FT' order by TESTTIME DESC "
        com14.CommandText = str
        rs14.Open Source:=com14
        If rs14.EOF = False Then 'MAC  exist
            GetMacFromTestRecord = UCase(Trim(rs14.Fields("MAC")))
            
        Else
            GetMacFromTestRecord = ""
        End If
        rs14.Close
        con14.Close
        Exit Function
    ElseIf equipment = "" Then 'equipment =""
        str = "select top 1 MAC from [test_equ] " & _
                    " where barcode = '" & Trim(sn) & "' and pass = N'通过' and len(mac)=12  order by TESTTIME DESC "
        com14.CommandText = str
        rs14.Open Source:=com14
        If rs14.EOF = False Then 'MAC  exist
            GetMacFromTestRecord = UCase(Trim(rs14.Fields("MAC")))
            
        Else
            GetMacFromTestRecord = ""
        End If
        rs14.Close
        con14.Close
        Exit Function
    End If
    
    str = "select  1 from [test_equ] " & _
        " where barcode = '" & Trim(sn) & "' and pass = N'通过'  AND equipment LIKE '" & equipment & "%' "
    com14.CommandText = str
    rs14.Open Source:=com14
    If rs14.EOF = False Then 'equipment LIKE 'XXXX%'  exist
        '''''''''''''
        conSQL01_1.Open ConnectionString:=strConn
        Set comSQL01_1 = New ADODB.Command
        comSQL01_1.ActiveConnection = conSQL01_1
        
        str = "select top 1 MAC from [test_equ] " & _
        " where barcode = '" & Trim(sn) & "' and pass = N'通过' and len(mac)=12 AND equipment LIKE '" & equipment & "%' order by TESTTIME DESC "
        comSQL01_1.CommandText = str
        rsSQL01_1.Open Source:=comSQL01_1
        If rsSQL01_1.EOF = False Then ''  exist mac
            GetMacFromTestRecord = UCase(Trim(rsSQL01_1.Fields("MAC")))
        Else
            GetMacFromTestRecord = "1"
        End If
        ''''''''''''
        rsSQL01_1.Close
        conSQL01_1.Close
        rs14.Close
        con14.Close
        Exit Function
        
    Else
        GetMacFromTestRecord = ""
        rs14.Close
        con14.Close
        Exit Function
    End If
    
End Function
