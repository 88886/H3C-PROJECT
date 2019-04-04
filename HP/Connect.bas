Attribute VB_Name = "Connect"
'' TBT 20150415

Public conn1 As New ADODB.Connection
'conn as default connection to server08
Public conn As New ADODB.Connection
'connection to the FTPC production
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
'进制转换函数：


'16 to 10
Function HEXTODEC(ByVal X As String) As String
Dim A() As String, i As Long, UNIT As Integer
For i = 1 To Len(X)
If Not IsNumeric("&h" & Mid(X, i, 1)) Then MsgBox "NOT A HEX FORMAT!", 64, "INFO": Exit Function
Next
X = String((6 - Len(X) Mod 6) Mod 6, "0") & X

UNIT = Len(X) \ 6 - 1
ReDim A(UNIT)
For i = 0 To UNIT
A(i) = CLng("&h" & Mid(X, i * 6 + 1, 6))
Next
For i = 0 To UNIT
A(i) = multi(A(i), POWERS(UNIT - i))
HEXTODEC = sums(HEXTODEC, A(i))
Next
End Function
Function sums(ByVal X As String, ByVal Y As String) As String ' sum of two hugehexnum（两个大数之和）
Dim max As Long, temp As Long, i As Long, result As Variant
max = IIf(Len(X) >= Len(Y), Len(X), Len(Y))
X = Right(String(max, "0") & X, max)
Y = Right(String(max, "0") & Y, max)
ReDim result(0 To max)
For i = max To 1 Step -1
result(i) = Val(Mid(X, i, 1)) + Val(Mid(Y, i, 1))
Next
For i = max To 1 Step -1
temp = result(i) \ 10
result(i) = result(i) Mod 10
result(i - 1) = result(i - 1) + temp
Next
If result(0) = 0 Then result(0) = ""
sums = Join(result, "")
Erase result

End Function

Function multi(ByVal X As String, ByVal Y As String) As String 'multi of two huge hexnum（两个大数之积）
Dim result As Variant
Dim xl As Long, yl As Long, temp As Long, i As Long
xl = Len(Trim(X))
yl = Len(Trim(Y))

ReDim result(1 To xl + yl)
For i = 1 To xl
For temp = 1 To yl
result(i + temp) = result(i + temp) + Val(Mid(X, i, 1)) * Val(Mid(Y, temp, 1))
Next
Next

For i = xl + yl To 2 Step -1
temp = result(i) \ 10
result(i) = result(i) Mod 10
result(i - 1) = result(i - 1) + temp
Next

If result(1) = "0" Then result(1) = ""
multi = Join(result, "")
Erase result

End Function
Function POWERS(ByVal X As Integer) As String ' GET 16777216^X,ie 16^(6*x)（16777216的X 次方）
POWERS = 1
Dim i As Integer
For i = 1 To X
POWERS = multi(POWERS, CLng(&H1000000))
Next
End Function
Function half(ByVal X As String) As String 'get half of x（取半）
X = 0 & X
Dim i As Long
Dim result As Variant
ReDim result(2 To Len(X)) As String
For i = 2 To Len(X)
result(i) = CStr(Val(Mid(X, i, 1)) \ 2 + IIf(Val(Mid(X, i - 1, 1)) Mod 2 = 1, 5, 0))
Next
half = Join(result, "")
If Left(half, 1) = "0" Then half = Right(half, Len(half) - 1) ' no zero ahead
End Function


'另一个有用的函数：
Function POWERXY(ByVal X As Integer, ByVal Y As Integer) As String 'GET X^Y（X 的 Y 次方）
Dim i As Integer
POWERXY = X
For i = 2 To Y
POWERXY = multi(POWERXY, X)
Next
End Function

Public Function excuteUpdate(sSQLStatement As String) As String
  On Error GoTo errorHandler
  conn.Execute (sSQLStatement)
  excuteUpdate = ""
  Exit Function
errorHandler:
  excuteUpdate = Err.Description
End Function


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

Public Function checkPrintPreCondition(model As String, position As Integer) As Boolean
    Dim temp As New Recordset
    Dim sql As String
    If position = 1 Then
        sql = "select * from EAN_HP_Setting_2014 where part_number = '" + model + "' and [double_50*20] = 1"
    ElseIf position = 2 Then
        sql = "select * from EAN_HP_Setting_2014 where part_number = '" + model + "' and [double_48*6] = 1"
    ElseIf position = 3 Then
        sql = "select * from EAN_HP_Setting_2014 where part_number = '" + model + "' and [single_50*20] = 1"
    ElseIf position = 4 Then
        sql = "select * from EAN_HP_Setting_2014 where part_number = '" + model + "' and [single_48*6] = 1"
    End If
    temp.Open sql, conn, adOpenForwardOnly
    If temp.EOF = True Then
        checkPrintPreCondition = False
        temp.Close
        Exit Function
    Else
        checkPrintPreCondition = True
        temp.Close
        Exit Function
    End If
End Function
'return 2 means print two labels
'return 1 means print one label
'return 0 means cannot print label
Public Function HPPrintDouble(model As String) As Integer
    Dim temp As New Recordset
    Dim sql As String
    sql = "select * from EAN_HP_Setting_2014 where part_number = '" + model + "' and [single_14.6*7.7] =1 and [backup] = 1"
    temp.Open sql, conn, adOpenForwardOnly
    If temp.EOF = True Then
        temp.Close
        sql = "select * from EAN_HP_Setting_2014 where part_number = '" + model + "' and [single_14.6*7.7] = 1"
        temp.Open sql, conn, adOpenForwardOnly
        If temp.EOF = False Then
            HPPrintDouble = 1
        Else
            HPPrintDouble = 0
        End If

        temp.Close
        Exit Function
    Else
        HPPrintDouble = 2
        temp.Close
        Exit Function
    End If
End Function

Public Function GetResByAction(serial_number As String, action As String) As String
    On Error GoTo errorHandler
    Dim rs14 As ADODB.Recordset
    Dim cmd As ADODB.Command
   
    Set rs14 = New ADODB.Recordset
    Set com = New ADODB.Command

    cmd.ActiveConnection = connFTPC
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

Public Function GetPBState(partlist As String) As String
    On Error GoTo errorHandler
    Dim cmd As New ADODB.Command
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
    End If
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[PbHandler]"
    cmd.Parameters.Append cmd.CreateParameter("partlist", adVarChar, adParamInput, 8000, partlist)
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 8)
    cmd.Parameters.Append cmd.CreateParameter("first", adVarChar, adParamOutput, 1)
    cmd.Execute
    Select Case cmd("res")
        Case "No"
            GetPBState = "NPb"
        Case "Non"
            MsgBox "此工单包含0302阶单板或整机阶未设定有铅无铅,请相关ME去设定!"
            GetPBState = ""
            Exit Function
        Case "Half"
             GetPBState = "N4"
        Case "Yes"
             GetPBState = "N4"
    End Select
    Exit Function
errorHandler:
    GetPBState = ""
End Function

Public Function getPartList(tempWO As String) As String
    Dim sql As String, partlist As String
    Dim rec As New Recordset
    If connFTPC.State = 0 Then
        connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
        connFTPC.Open
    End If
    sql = "select distinct assembly from afg_b_SAPWIPReport A  where A.leading_order ='" & tempWO & "' and assembly like 'HWF0302%' and assembly not like '%-S%'"
    
    rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
    If rec.EOF = True Then
        rec.Close
        sql = "select distinct b.Order_Type_S from Work_Order A,UDA_Order B where A.order_key = b.object_key and a.order_number = '" & tempWO & "'"
        rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
        If rec.Fields(0) = "PP05" Then
            getPartList = "TTTTTTTT"
            rec.Close
            Exit Function
        Else
'            getPartList = ""
'            MsgBox "SAP中此工单不包含0302阶单板不能打印,请确认!"
'            rec.Close
'            Exit Function
            rec.Close
            'sql = "select distinct assembly from afg_b_SAPWIPReport A  where A.leading_order ='" & tempWO & "' and ( assembly like 'HWF0231%' or assembly like 'HWF0235%')"
            'rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
            'If rec.EOF = True Then
             '   getPartList = ""
             '   MsgBox "SAP中此工单不包含0302阶或整机阶,不能打印,请确认!"
             '   rec.Close
             '   Exit Function
            'Else
            '    Do While Not rec.EOF
            '        partlist = partlist + Mid(rec!assembly, 4, 8) + ";"
            '        rec.MoveNext
             '   Loop
            'End If
            sql = "select distinct B.part_number from afg_active_90.dbo.WORK_ORDER A JOIN afg_active_90.dbo.WORK_ORDER_ITEMS B ON A.order_key=B.order_key where order_number='" & tempWO & "' and ( B.part_number like 'HWF0231%' or B.part_number like 'HWF0235%')"
            rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                getPartList = ""
                MsgBox "SAP中此工单不包含0302阶或整机阶,不能打印,请确认!"
                rec.Close
                Exit Function
            Else
                Do While Not rec.EOF
                    'partlist = partlist + Mid(rec!assembly, 4, 8) + ";"
                    partlist = partlist + Mid(rec!Part_Number, 4, 8) + ";"
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
'    Debug.Print (partlist)
    getPartList = partlist
    rec.Close
End Function

Public Function IsCheckRohs(model As String) As Boolean
    On Error GoTo errorHandler
    Dim sql As String, partlist As String
    Dim rec As New Recordset
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
    End If
    sql = "select upload_power_code from hp where h3c_bom_code = '" & model & "'"
    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
    If rec.Fields(0) = True Then
        IsCheckRohs = False
        rec.Close
        Exit Function
    Else
        IsCheckRohs = True
    End If
    rec.Close
    Exit Function
errorHandler:
    IsCheckRohs = True
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
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 16, update_user)
    cmd.Execute

    Set cmd.ActiveConnection = Nothing
    UploadH3C_PB = True
    Exit Function
errorHandler:
    UploadH3C_PB = False
End Function

'    @Serial_Number nvarchar(100),
'    @Order_Number nvarchar(100),
'    @Part_Number nvarchar(100),
'    @Product_Desc nvarchar(100),
'    @Product_Model nvarchar(100),
'    @Product_Material nvarchar(100),
'    @Product_Power nvarchar(100),
'    @update_user varchar(32)
Public Function UploadConsen_Rec(serial_number As String, Order_Number As String, Part_Number As String, Product_Desc As String, Product_Model As String, Product_Material As String, Product_Power As String, SalesOrder As String, CCC As String, ChinaRoHS As String, WEEE As String, Laser As String, update_user As String) As Boolean
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
    cmd.Parameters.Append cmd.CreateParameter("Order_Number", adVarChar, adParamInput, 100, Order_Number)
    cmd.Parameters.Append cmd.CreateParameter("Part_Number", adVarChar, adParamInput, 100, Part_Number)
    cmd.Parameters.Append cmd.CreateParameter("Product_Desc", adVarChar, adParamInput, 100, Product_Desc)
    cmd.Parameters.Append cmd.CreateParameter("Product_Model", adVarChar, adParamInput, 100, Product_Model)
    cmd.Parameters.Append cmd.CreateParameter("Product_Material", adVarChar, adParamInput, 100, Product_Material)
    cmd.Parameters.Append cmd.CreateParameter("Product_Power", adVarChar, adParamInput, 100, Product_Power)
    cmd.Parameters.Append cmd.CreateParameter("SalesOrder", adVarChar, adParamInput, 50, SalesOrder)
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
            getPbByPartList = "N*"
            taskOrderFlag = True
            rec.Close
            Exit Function
        Else
'            getPbByPartList = "Non"
'            MsgBox "SAP中此工单不包含0302阶单板不能打印,请确认!"
'            rec.Close
'            Exit Function
            rec.Close
            'sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & tempWO & "' and ( assembly like 'HWF0231%' or assembly like 'HWF0235%')"
            sql = "select distinct B.part_number from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A JOIN [10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B ON A.order_key=B.order_key where order_number='" & tempWO & "' and ( B.part_number like 'HWF0231%' or B.part_number like 'HWF0235%')"
            rec.Open sql, conn, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                getPbByPartList = "Non"
                MsgBox "SAP中此工单不包含0302阶或整机阶,不能打印,请确认!"
                rec.Close
                Exit Function
            Else
                Do While Not rec.EOF
                    partlist = partlist + Mid(rec!Part_Number, 4, 8) + ";"
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
