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

Public Function UploadH3CInfo(pc As Boolean, serial_number As String, hv As String, status As String, power_code As String, power_origin As String, update_user As String) As Boolean
    On Error GoTo errorHandler
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com As ADODB.Command
    Dim str As String, pb As String, partlist As String
    
'    partlist = Connect.GetResByAction(serial_number, "GetPartList")
'    If partlist = "" Then
'        pb = "Pb"
'    End If
'
'    If pb = "" Then
'        pb = Connect.GetPBState(partlist)
'        If pb = "" Then
'            UploadH3CInfo = False
'            MsgBox "该条码所对应的铅状态无法判断", vbOKOnly + vbExclamation, "该条码所对应的铅状态无法判断"
'            Exit Function
'        End If
'    End If
'
'    If power_code = "" Or power_code = "/" Then
'        power_code = "N/A"
'    End If
    
    partlist = Connect.GetResByAction(serial_number, "GetPartList")
    If partlist = "" And pc = True Then
        pb = "Pb"
    ElseIf partlist = "" And pc = False Then
        UploadH3CInfo = False
        MsgBox "该条码所对应的BOM中没有0302的下阶", vbOKOnly + vbExclamation, "BOM中没有0302的下阶"
        Exit Function
    End If
   
    
    If pb = "" Then
        pb = Connect.GetPBState(partlist)
        If pb = "" Then
            UploadH3CInfo = False
            MsgBox "该条码所对应的铅状态无法判断", vbOKOnly + vbExclamation, "该条码所对应的铅状态无法判断"
            Exit Function
        End If
    End If
    
    If power_code = "" Or power_code = "/" Then
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
    cmd.Parameters.Append cmd.CreateParameter("5000_status", adVarChar, adParamInput, 4, status)
    cmd.Parameters.Append cmd.CreateParameter("power_code", adVarChar, adParamInput, 16, power_code)
    cmd.Parameters.Append cmd.CreateParameter("power_origin", adVarChar, adParamInput, 16, power_origin)
    cmd.Parameters.Append cmd.CreateParameter("pb", adVarChar, adParamInput, 8, pb)
    cmd.Parameters.Append cmd.CreateParameter("update_user", adVarChar, adParamInput, 16, update_user)
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    UploadH3CInfo = True
    Exit Function
errorHandler:
    UploadH3CInfo = False
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


