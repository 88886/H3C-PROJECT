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
    
    strOraCon = "Provider=OraOLEDB.Oracle;Password=fujiadmin;Persist Security Info=True;User ID=fujiadmin;Data Source=fujidb"
    strDB2Con = "Provider=IBMDADB2.1;Password=T0mcat4Fun;Persist Security Info=True;User ID=cgsapp;Data Source=CGS"
    
    If fCntoDB2 <> 0 Then
        MsgBox "Connect to DB2 Error!", vbOKOnly, "Error"
        Exit Sub
    End If
    
    If fCntoOra <> 0 Then
        MsgBox "Connect to OraCle Frror!", vbOKOnly, "Error"
        Exit Sub
    End If
    
    strSqlDB2 = "SELECT a.item_key, b.Part_number,a.Quantity," & _
                "a.Init_tmst from cgs.item a join cgs.Part_number b " & _
                " on a.Part_number_key=b.Part_number_key " & _
                " Where ITEM_DESC = '' "
    
    Set rsDB2 = cnDB2.Execute(strSqlDB2)
    
    If rsDB2.EOF = False Then
R_Ora:
        ii = 0
        cnOra.BeginTrans
        cnDB2.BeginTrans
        
        While Not rsDB2.EOF
            On Error Resume Next
            If Len(rsDB2("Part_number")) > 64 Then GoTo insErr
            strSqlTmp = "insert into T_DID (DIDDID,DIDPTN,DIDBAR,DIDQTY) values ('ITEM" & rsDB2("item_key") & "','" & rsDB2("Part_number") & "','" & rsDB2("Part_number") & "','" & rsDB2("Quantity") & "' )"
            cnOra.Execute strSqlTmp
            DoEvents
insErr:
            strSqlDB2 = "update cgs.item set item_desc = 'X' where item_key=" & rsDB2("item_key") & " "
            cnDB2.Execute (strSqlDB2)
            DoEvents

            rsDB2.MoveNext
            ii = ii + 1
            If ii >= 5000 Then
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
    cnOra.RollbackTrans
    cnDB2.RollbackTrans
   
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




