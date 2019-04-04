VERSION 5.00
Begin VB.Form Cgs2Fuji 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cgs2Fuji Manual"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "ITEM ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Cgs2Fuji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cnDB2            As New ADODB.Connection
Public rsDB2          As New ADODB.Recordset
Public cnOra            As New ADODB.Connection
Public rsOra          As New ADODB.Recordset

Public strCommand       As String
Public strOraCon        As String
Public strDB2Con        As String
Public ITEM_KEY As String


Private Sub Command1_Click()
    Dim strSqlDB2 As String
    Dim strSqlTmp As String
    Dim item_id As String
    
    If Trim(Text1.Text) = "" Then
        MsgBox "Please enter Item ID!", vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    If Left(UCase(Trim(Text1.Text)), 4) = "ITEM" Then
        item_id = Mid(Trim(Text1.Text), 5)
    Else
        item_id = Trim(Text1.Text)
    End If
    
    strOraCon = "Provider=OraOLEDB.Oracle;Password=fujiadmin;Persist Security Info=True;User ID=fujiadmin;Data Source=fujidb"
    strDB2Con = "Provider=IBMDADB2.1;Password=T0mcat4Fun;Persist Security Info=True;User ID=cgsapp;Data Source=CGS"
    
    If fCntoDB2 <> 0 Then
        MsgBox "Connect to DB2 Error!", vbOKOnly, "Error"
        Exit Sub
    End If
    
    If fCntoOra <> 0 Then
        MsgBox "Connect to OraCle Error!", vbOKOnly, "Error"
        Exit Sub
    End If
    
    strSqlDB2 = "SELECT a.item_key, b.Part_number,case when a.Quantity>20000 then 20000 else a.Quantity end as Quantity," & _
                "a.Init_tmst from cgs.item a join cgs.Part_number b " & _
                " on a.Part_number_key=b.Part_number_key " & _
                " Where item_key = " & item_id & " "
    
    Set rsDB2 = cnDB2.Execute(strSqlDB2)
    
    If rsDB2.EOF = False Then
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
        Wend
        On Error GoTo Rollback_end
        cnOra.CommitTrans
        cnDB2.CommitTrans
        Label2.Caption = "ITEM ID: " & Trim(Text1.Text) & " transfer OK!"
    Else
        Label2.Caption = "ITEM ID: " & Trim(Text1.Text) & " not Find!"
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

Private Sub Text1_Change()
    Label2.Caption = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call Command1_Click
    End If
End Sub
