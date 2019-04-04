Attribute VB_Name = "InsFuji"

Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public cnDB2            As New ADODB.Connection
Public rsDB2          As New ADODB.Recordset
Public cnOra            As New ADODB.Connection
Public rsOra          As New ADODB.Recordset

Public strCommand       As String
Public strOraCon        As String
Public strDB2Con        As String
Public ITEM_KEY As String


Public Sub Main()
    Dim strSqlDB2 As String
    Dim strSqlTmp As String
    Dim ii As Double
    
    On Error GoTo Rollback_end
    strOraCon = "Provider=OraOLEDB.Oracle;Password=fujiadmin;Persist Security Info=True;User ID=fujiadmin;Data Source=fujidb"
    strDB2Con = "Provider=IBMDADB2.1;Password=T0mcat4Fun;Persist Security Info=True;User ID=cgsapp;Data Source=CGS"
    
    If fCntoDB2 <> 0 Then
        Call mail_send("NG", Now(), "Connect to DB2 Error")
        Exit Sub
    End If
    
    If fCntoOra <> 0 Then
        Call mail_send("NG", Now(), "Connect to Oracle Error")
        Exit Sub
    End If
    
    strSqlDB2 = "SELECT a.item_key,char(date(a.MODIFIED_TMST)) as MODIFIED_TMST,a.ITEM_TYPE_KEY, b.Part_number,case when a.Quantity>20000 then 20000 else a.Quantity end as Quantity,a.item_id " & _
                " from cgs.item_fuji a join cgs.Part_number b " & _
                " on a.Part_number_key=b.Part_number_key " & _
                " Where a.ITEM_DESC is null with ur "
    
    Set rsDB2 = cnDB2.Execute(strSqlDB2)
    
    If rsDB2.EOF = False Then
R_Ora:
        ii = 0
        cnOra.BeginTrans
        cnDB2.BeginTrans
        
        While Not rsDB2.EOF
            On Error GoTo insErr
            If Len(Trim(rsDB2("Part_number"))) > 64 Then GoTo insErr
            If Len(Trim(rsDB2("item_id"))) = 12 And Left(Trim(rsDB2("item_id")), 2) = "L0" Then
                strSqlTmp = "select DIDDID from T_DID where DIDDID='" & rsDB2("item_id") & "'"
                Set rsOra = cnOra.Execute(strSqlTmp)
                If rsOra.EOF Then
                    strSqlTmp = "insert into T_DID (DIDDID,DIDPTN,DIDBAR,DIDQTY,DIDFMDF) values ('" & rsDB2("item_id") & "','" & Trim(rsDB2("Part_number")) & "','" & rsDB2("Part_number") & "','" & rsDB2("Quantity") & "','" & rsDB2("MODIFIED_TMST") & "' )"
                    cnOra.Execute strSqlTmp
                    strSqlTmp = "insert into T_DID (DIDDID,DIDPTN,DIDBAR,DIDQTY,DIDFMDF) values ('ITEM" & rsDB2("item_key") & "','" & Trim(rsDB2("Part_number")) & "','" & rsDB2("Part_number") & "','" & rsDB2("Quantity") & "','" & rsDB2("MODIFIED_TMST") & "' )"
                    cnOra.Execute strSqlTmp
                    DoEvents
                Else
                    strSqlTmp = "select DIDDID from T_DID where DIDDID='ITEM" & rsDB2("item_key") & "'"
                    Set rsOra = cnOra.Execute(strSqlTmp)
                    If rsOra.EOF Then
                        strSqlTmp = "insert into T_DID (DIDDID,DIDPTN,DIDBAR,DIDQTY,DIDFMDF) values ('ITEM" & rsDB2("item_key") & "','" & Trim(rsDB2("Part_number")) & "','" & rsDB2("Part_number") & "','" & rsDB2("Quantity") & "','" & rsDB2("MODIFIED_TMST") & "' )"
                        cnOra.Execute strSqlTmp
                        DoEvents
                    End If
                End If
            Else
                strSqlTmp = "select DIDDID from T_DID where DIDDID='ITEM" & rsDB2("item_key") & "'"
                Set rsOra = cnOra.Execute(strSqlTmp)
                If rsOra.EOF Then
                    strSqlTmp = "insert into T_DID (DIDDID,DIDPTN,DIDBAR,DIDQTY,DIDFMDF) values ('ITEM" & rsDB2("item_key") & "','" & Trim(rsDB2("Part_number")) & "','" & rsDB2("Part_number") & "','" & rsDB2("Quantity") & "','" & rsDB2("MODIFIED_TMST") & "' )"
                    cnOra.Execute strSqlTmp
                    DoEvents
                End If
            End If
'''            Set rsOra = cnOra.Execute(strSqlTmp)
'''            If rsOra.EOF Then
'''                If Len(Trim(rsDB2("item_id"))) = 12 And Left(Trim(rsDB2("item_id")), 2) = "L0" Then
'''                    strSqlTmp = "insert into T_DID (DIDDID,DIDPTN,DIDBAR,DIDQTY,DIDFMDF) values ('" & rsDB2("item_id") & "','" & Trim(rsDB2("Part_number")) & "','" & rsDB2("Part_number") & "','" & rsDB2("Quantity") & "','" & rsDB2("MODIFIED_TMST") & "' )"
'''                    cnOra.Execute strSqlTmp
'''                    strSqlTmp = "insert into T_DID (DIDDID,DIDPTN,DIDBAR,DIDQTY,DIDFMDF) values ('ITEM" & rsDB2("item_key") & "','" & Trim(rsDB2("Part_number")) & "','" & rsDB2("Part_number") & "','" & rsDB2("Quantity") & "','" & rsDB2("MODIFIED_TMST") & "' )"
'''                    cnOra.Execute strSqlTmp
'''                Else
'''                    strSqlTmp = "insert into T_DID (DIDDID,DIDPTN,DIDBAR,DIDQTY,DIDFMDF) values ('ITEM" & rsDB2("item_key") & "','" & Trim(rsDB2("Part_number")) & "','" & rsDB2("Part_number") & "','" & rsDB2("Quantity") & "','" & rsDB2("MODIFIED_TMST") & "' )"
'''                    cnOra.Execute strSqlTmp
'''                End If
'''                'cnOra.Execute strSqlTmp
''''                If rsDB2("ITEM_TYPE_KEY") = 7 Then
''''                    On Error GoTo insErr
''''                    strSqlTmp = "insert into T_TRAY( DID, POSX, POSY, TRAYCOUNT) values ('ITEM" & rsDB2("item_key") & "',1,1,1 )"
''''                    cnOra.Execute strSqlTmp
''''                End If
'''                DoEvents
'''            End If
            strSqlDB2 = "update cgs.item_fuji set item_desc = 'X' where item_key=" & rsDB2("item_key") & " "
            cnDB2.Execute (strSqlDB2)
            DoEvents

insErr:
            rsDB2.MoveNext
            ii = ii + 1
            If ii >= 2000 Then
                On Error GoTo Rollback_end
                cnOra.CommitTrans
                cnDB2.CommitTrans
                GoTo R_Ora
            End If
        Wend
        On Error GoTo Rollback_end
        cnOra.CommitTrans
        cnDB2.CommitTrans
    End If
    
    cnDB2.Close
    If cnOra.State = 1 Then cnOra.Close
    Set rsDB2 = Nothing
    Exit Sub
    
Rollback_end:
    Call mail_send("NG", Now(), Err.Description)
    cnOra.RollbackTrans
    cnDB2.RollbackTrans
    Exit Sub
   
End Sub


Private Function fCntoDB2() As Integer
On Error GoTo Err_fCntoDB2

    fCntoDB2 = -1
    cnDB2.CursorLocation = adUseClient
    cnDB2.Open strDB2Con
    fCntoDB2 = 0
Exit Function
Err_fCntoDB2:
    fCntoDB2 = -1
End Function

Private Function fCntoOra() As Integer
On Error GoTo Err_fCntoOra
    
    fCntoOra = -1
    
    cnOra.Open strOraCon
    fCntoOra = 0
Exit Function
Err_fCntoOra:
    fCntoOra = -1
End Function

Private Sub mail_send(result As String, lDate As String, MailError As String)
    Dim NameS As String
    Dim Email As Object
    If result = "OK" Then
'''        NameS = "http://schemas.microsoft.com/cdo/configuration/"
'''
'''        Set Email = CreateObject("CDO.Message")
'''        Email.From = "mike.jiang@cn.asteelflash.com"
'''        Email.To = "carson.hu@cn.asteelflash.com,Jimmy.Sun@cn.asteelflash.com"
'''        Email.CC = "mike.jiang@cn.asteelflash.com"
'''        Email.Subject = "H3C Report finished at (" & Now() & ")"
'''        Email.TextBody = "Oh yeah!~  ^_^ *_* ,H3C Report has been finished! Date:" & lDate
'''        Email.Configuration.Fields.Item(NameS & "sendusing") = 2
'''        Email.Configuration.Fields.Item(NameS & "smtpserver") = "sz-ex01.cn1.flashelec.com"
'''
'''        Email.Configuration.Fields.Item(NameS & "smtpserverport") = 25
'''        Email.Configuration.Fields.Item(NameS & "smtpauthenticate") = 1
'''        Email.Configuration.Fields.Item(NameS & "sendusername") = "mike.jiang"
'''        Email.Configuration.Fields.Item(NameS & "sendpassword") = "123~z456"
'''
'''        Email.Configuration.Fields.Update
'''        Email.Send
    Else
        NameS = "http://schemas.microsoft.com/cdo/configuration/"
        
        Set Email = CreateObject("CDO.Message")
        Email.From = "MES_Auto_Mail@asteelflash.com"
        Email.To = "mark.qian@asteelflash.com"
        Email.CC = "mike.jiang@asteelflash.com,carson.hu@asteelflash.com,Jimmy.Sun@asteelflash.com,allen.yan@asteelflash.com"
        Email.Subject = "CGSTOFUJI exception mail(" & Now() & ")"
        Email.TextBody = "CGSTOFUJI has a error,please check! Date:" & lDate & " ERROR POINT:" & MailError
        Email.Configuration.Fields.Item(NameS & "sendusing") = 2
        Email.Configuration.Fields.Item(NameS & "smtpserver") = "10.11.1.115"
    
        Email.Configuration.Fields.Item(NameS & "smtpserverport") = 25
        Email.Configuration.Fields.Item(NameS & "smtpauthenticate") = 1
        Email.Configuration.Fields.Item(NameS & "sendusername") = "mes_auto_mail"
        Email.Configuration.Fields.Item(NameS & "sendpassword") = "afg-app01"
    
        Email.Configuration.Fields.Update
        Email.Send
    End If
End Sub








