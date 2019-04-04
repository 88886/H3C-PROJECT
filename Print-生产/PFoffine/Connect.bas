Attribute VB_Name = "Connect"
Public conn As New ADODB.Connection
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

Public Function getPrintedhistoryData(ByVal sql As String, beginsn As String, qty As Integer) As Boolean
    
    Dim rec As New ADODB.Recordset
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    If rec.EOF = False Then
        beginsn = Trim(rec.Fields(2))
        qty = Trim(rec.Fields(4))
        getPrintedhistoryData = True
    Else
        beginsn = ""
        qty = 0
        getPrintedhistoryData = False
    End If
End Function

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

Function getStartSN(part_number As String) As String
On Error GoTo errorHandler
     Dim maxSerial As Integer
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[PacketFrontPrintHandler]"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 8, "query")
    cmd.Parameters.Append cmd.CreateParameter("startString", adVarChar, adParamInput, 4, part_number)
    cmd.Parameters.Append cmd.CreateParameter("printDate", adVarChar, adParamInput, 8, Format$(Now, "yyyyMMdd"))
    cmd.Parameters.Append cmd.CreateParameter("count", adInteger, adParamInput, 4, 0)
    cmd.Parameters.Append cmd.CreateParameter("maxSerial", adInteger, adParamOutput)
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 100)
    cmd.Execute
    
    maxSerial = cmd("maxSerial")
    getStartSN = part_number + "11" + Format$(Now, "yyMMdd") + Right$("000" + CStr(maxSerial), 3)
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
   getStartSN = Err.Description
End Function

Function getNewStartSN(part_number As String, part_version As String) As String
On Error GoTo errorHandler
     Dim maxSerial As Integer
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[PacketFrontPrintHandler]"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 8, "query")
    cmd.Parameters.Append cmd.CreateParameter("startString", adVarChar, adParamInput, 4, part_number)
    cmd.Parameters.Append cmd.CreateParameter("printDate", adVarChar, adParamInput, 8, Format$(Now, "yyyyMMdd"))
    cmd.Parameters.Append cmd.CreateParameter("count", adInteger, adParamInput, 4, 0)
    cmd.Parameters.Append cmd.CreateParameter("maxSerial", adInteger, adParamOutput)
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 100)
    cmd.Execute
    
    maxSerial = cmd("maxSerial")
    getNewStartSN = part_number + Mid(part_version, 1, 2) + Format$(Now, "yyMMdd") + Right$("000" + CStr(maxSerial), 4)
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
   getNewStartSN = Err.Description
End Function
Function getNewStartSN_6(part_number As String, part_version As String) As String
On Error GoTo errorHandler
     Dim maxSerial As Integer
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[PacketFrontPrintHandler]"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 8, "query")
    cmd.Parameters.Append cmd.CreateParameter("startString", adVarChar, adParamInput, 4, part_number)
    cmd.Parameters.Append cmd.CreateParameter("printDate", adVarChar, adParamInput, 8, Format$(Now, "yyyyMMdd"))
    cmd.Parameters.Append cmd.CreateParameter("count", adInteger, adParamInput, 4, 0)
    cmd.Parameters.Append cmd.CreateParameter("maxSerial", adInteger, adParamOutput)
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 100)
    cmd.Execute
    
    maxSerial = cmd("maxSerial")
    getNewStartSN_6 = part_number + Mid(part_version, 1, 2) + Format$(Now, "yyMMdd") + Right$("00" + CStr(maxSerial), 3)
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
   getNewStartSN_6 = Err.Description
End Function
Function getNewStartSN_7(part_number As String, part_version As String) As String    '复制方法_6,区别是前面的sn取值位数不同
On Error GoTo errorHandler
     Dim maxSerial As Integer
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[PacketFrontPrintHandler]"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 8, "query")
    cmd.Parameters.Append cmd.CreateParameter("startString", adVarChar, adParamInput, 5, part_number)   ' sn前面取5位
    cmd.Parameters.Append cmd.CreateParameter("printDate", adVarChar, adParamInput, 8, Format$(Now, "yyyyMMdd"))
    cmd.Parameters.Append cmd.CreateParameter("count", adInteger, adParamInput, 4, 0)
    cmd.Parameters.Append cmd.CreateParameter("maxSerial", adInteger, adParamOutput)
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 100)
    cmd.Execute
    
    'maxSerial = 1(值等于1 save)
    maxSerial = cmd("maxSerial")
    getNewStartSN_7 = part_number + Mid(part_version, 1, 2) + Format$(Now, "yyMMdd") + Right$("00" + CStr(maxSerial), 3) 'sn后面三位
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
   getNewStartSN_7 = Err.Description
End Function



Function getWeekOfYear(date_1 As Date) As String
    Dim date_0 As Date
    date_0 = Year(date_1) & "/01/01"
    weekOfYear = DateDiff("ww", date_0, date_1) + 1
    getWeekOfYear = IIf(weekOfYear < 10, "0" & weekOfYear, "" & weekOfYear)
End Function

Function getStartSNF846(part_number As String, part_version As String) As String
On Error GoTo errorHandler
     Dim maxSerial As Integer
     Dim yearFlag As String
     
     yearFlag = Chr(Year(Now) - 2015 + 15 + 65)     ''2015 : P, 2016 : Q, 2017 : R...
     
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[PacketFrontPrintHandler_F846]"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 8, "query")
    cmd.Parameters.Append cmd.CreateParameter("startString", adVarChar, adParamInput, 4, part_number)
    cmd.Parameters.Append cmd.CreateParameter("printDate", adVarChar, adParamInput, 8, Format$(Now, "yyyyMMdd"))
    cmd.Parameters.Append cmd.CreateParameter("count", adInteger, adParamInput, 4, 0)
    cmd.Parameters.Append cmd.CreateParameter("maxSerial", adInteger, adParamOutput)
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 100)
    cmd.Parameters.Append cmd.CreateParameter("txtsn", adInteger, adParamInput, 3, 0)
    cmd.Execute
    
    maxSerial = cmd("maxSerial")
    getStartSNF846 = part_number + "" + part_version + yearFlag + getWeekOfYear(Now) + Right$("00" + dectohex(maxSerial), 2)
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
   getStartSNF846 = Err.Description
End Function

Function saveMaxSerialF846(startString As String, qty As Integer, isn As Integer) As String
On Error GoTo errorHandler
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[PacketFrontPrintHandler_F846]"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 8, "save")
    cmd.Parameters.Append cmd.CreateParameter("startString", adVarChar, adParamInput, 4, startString)
    cmd.Parameters.Append cmd.CreateParameter("printDate", adVarChar, adParamInput, 8, Format$(Now, "yyyyMMdd"))
    cmd.Parameters.Append cmd.CreateParameter("count", adInteger, adParamInput, 4, qty)
    cmd.Parameters.Append cmd.CreateParameter("maxSerial", adInteger, adParamOutput)
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 100)
    cmd.Parameters.Append cmd.CreateParameter("txtsn", adInteger, adParamInput, 3, isn)
    cmd.Execute
    saveMaxSerialF846 = "OK"
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
   saveMaxSerialF846 = Err.Description
End Function

Function saveMaxSerial(startString As String, qty As Integer, isn As Integer) As String
On Error GoTo errorHandler
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[PacketFrontPrintHandler_New]"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 8, "save")
    cmd.Parameters.Append cmd.CreateParameter("startString", adVarChar, adParamInput, 5, startString)   '打印保存到PacketFront表，5为startstring开头的取值位数
    cmd.Parameters.Append cmd.CreateParameter("printDate", adVarChar, adParamInput, 8, Format$(Now, "yyyyMMdd"))
    cmd.Parameters.Append cmd.CreateParameter("count", adInteger, adParamInput, 4, qty)
    cmd.Parameters.Append cmd.CreateParameter("maxSerial", adInteger, adParamOutput)
    cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 100)
    cmd.Parameters.Append cmd.CreateParameter("txtsn", adInteger, adParamInput, 3, isn)
    cmd.Execute
    saveMaxSerial = "OK"
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
   saveMaxSerial = Err.Description
End Function

Function getMACOfPackFront(ByVal sn As String) As String
    Dim rec As New ADODB.Recordset
    sql = "select sn, mac, wo from PacketFrontRecords where sn = '" & sn & "'"
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    If rec.EOF = False Then
        getMACOfPackFront = rec!mac
    Else
        getMACOfPackFront = ""
    End If
    
    Set rec = Nothing

End Function
Function savePackFrontRecords(sn As String, mac As String, wo As String) As String
    On Error GoTo errorHandler
     Dim maxSerial As Integer
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandText = "insert PacketFrontRecords(wo,sn,mac) values('" & wo & "','" & sn & "','" & mac & "')"
'    cmd.Parame
'    cmd.Parameters.Append cmd.CreateParameter("@sn", adVarChar, adParamInput, 32, sn)
'    cmd.Parameters.Append cmd.CreateParameter("@mac", adVarChar, adParamInput, 32, mac)
    cmd.Execute
    savePackFrontRecords = "OK"
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
   savePackFrontRecords = Err.Description
   Exit Function
End Function

Function savePackFrontRecordsF846(sn As String, mac As String, wo As String, fenpeihao As String, qty As Integer) As String
    On Error GoTo errorHandler
     Dim maxSerial As Integer
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandText = "insert PacketFrontRecordF846([fenpeihao],[wo],[sn],[mac],[qty],[lastprinttime]) values('" & fenpeihao & "','" & wo & "','" & sn & "','" & mac & "','" & qty & "',getdate())"
'    cmd.Parame
'    cmd.Parameters.Append cmd.CreateParameter("@sn", adVarChar, adParamInput, 32, sn)
'    cmd.Parameters.Append cmd.CreateParameter("@mac", adVarChar, adParamInput, 32, mac)
    cmd.Execute
    savePackFrontRecordsF846 = "OK"
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Exit Function
errorHandler:
   savePackFrontRecordsF846 = Err.Description
   Exit Function
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

'进制转换函数：


'16 to 10
Function HEXTODEC(ByVal X As String) As String
Dim a() As String, i As Long, UNIT As Integer
For i = 1 To Len(X)
If Not IsNumeric("&h" & Mid(X, i, 1)) Then MsgBox "NOT A HEX FORMAT!", 64, "INFO": Exit Function
Next
X = String((6 - Len(X) Mod 6) Mod 6, "0") & X

UNIT = Len(X) \ 6 - 1
ReDim a(UNIT)
For i = 0 To UNIT
a(i) = CLng("&h" & Mid(X, i * 6 + 1, 6))
Next
For i = 0 To UNIT
a(i) = multi(a(i), POWERS(UNIT - i))
HEXTODEC = sums(HEXTODEC, a(i))
Next
End Function




' 10 to 16
Function dectohex(ByVal hugenum As String) As String ' trans hugenum to hex

Do While Len(hugenum) > 2
dectohex = Hex(Val(Right(hugenum, 4)) Mod 16) & dectohex
For i = 1 To 4 'devide hugenum by 16
hugenum = half(hugenum)
Next
Loop
Dim tmp As String
Dim k As Integer

tmp = Hex(Val(hugenum)) & dectohex
For k = 1 To 12
    If Len(tmp) < 12 Then
        tmp = "0" & tmp
    Else
        Exit For
    End If
Next
dectohex = tmp
End Function










