VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmVerset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�趨���ֶ�Ӧ�汾(Setting Rev)"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVersion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10470
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtsn 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   4080
      Width           =   4095
   End
   Begin VB.TextBox txtpass 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   5760
      Width           =   4095
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "��ѯ(Query)"
      Height          =   735
      Left            =   6720
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "����(Insert)"
      Height          =   735
      Left            =   6720
      TabIndex        =   3
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(Delete)"
      Height          =   735
      Left            =   6720
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "ȷ��(Confirm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtVer 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   6600
      Width           =   4095
   End
   Begin VB.TextBox txtModel 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   4920
      Width           =   4095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgH3C 
      Height          =   3615
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   5
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lblPass 
      Caption         =   "�汾��Rev��:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label lblVer 
      Caption         =   "��ˮ�Ž���ֵ:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblModel 
      Caption         =   "��ˮ����ʼֵ:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblSN 
      Caption         =   "��               ��:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "frmVerset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conn As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim op As String

Private Sub cmdCancel_Click()
   txtsn.Text = ""
   txtModel.Text = ""
   txtVer.Text = ""
   txtpass.Text = ""
End Sub

Private Sub cmdConfirm_Click()
If op = "add" Then
   If txtsn.Text = "" Then
       MsgBox "���ֲ���Ϊ��!", vbOKOnly + vbExclamation, "Error"
       txtsn.SetFocus
       Exit Sub
   End If
End If
   If txtModel.Text = "" Then
      MsgBox "��ˮ����ʼֵ����Ϊ��!", vbOKOnly + vbExclamation, "Error"
      txtModel.SetFocus
      Exit Sub
   End If
   If txtpass.Text = "" Then
      MsgBox "��ˮ�Ž���ֵ����Ϊ��!", vbOKOnly + vbExclamation, "Error"
      txtpass.SetFocus
      Exit Sub
   End If
   If Len(txtModel.Text) <> 20 Then
        MsgBox "��ˮ����ʼֵ���Ȳ���!", vbExclamation + vbOKOnly, "Error"
        txtModel.SetFocus
        Exit Sub
   End If
    If Len(txtpass.Text) <> 20 Then
        MsgBox "��ˮ�Ž���ֵ���Ȳ���!", vbExclamation + vbOKOnly, "Error"
        txtpass.SetFocus
        Exit Sub
   End If
   If Val(Right(txtpass.Text, 6)) < Val(Right(txtModel.Text, 6)) Then
       MsgBox "��ˮ�Ž���ֵ����С����ʼֵ!", vbOKOnly + vbExclamation, "Error"
      txtpass.SetFocus
      Exit Sub
   End If
   If txtVer.Text = "" Then
      MsgBox "�汾����Ϊ��!", vbOKOnly + vbExclamation, "Error"
      txtVer.SetFocus
      Exit Sub
   End If
If op = "add" Then
   sql = "select * from revset  where model='" & txtsn.Text & "' and firstall='" & txtModel.Text & "' and endall='" & txtpass.Text & "' and ver='" & txtVer.Text & "'"
    If rec.State = 1 Then
        rec.Close
    End If
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    If rec.EOF = False Then
       MsgBox "�˻����Ѵ���!", vbOKOnly + vbExclamation, "Error"
       Exit Sub
    End If
    sql = "insert into revset(model,firstno,endno,ver,firstall,endall,normal) values('" & txtsn.Text & "'," & Right(txtModel.Text, 6) & "," & Right(txtpass.Text, 6) & ",'" & txtVer.Text & "','" & txtModel.Text & "','" & txtpass.Text & "','" & Mid(txtpass.Text, 12, 3) & "')"
End If
   conn.Execute sql
   Call cmdQuery_Click
   txtsn.Text = ""
   txtModel.Text = ""
   txtVer.Text = ""
   txtpass.Text = ""
End Sub

Private Sub cmdDelete_Click()
'      sql = "select * from revset where   substring(firstall,12,3)='" & Mid(txtsn.Text, 12, 3) & "'"
''      sql = "select * from revset where model='" & Mid(txtsn.Text, 3, 8) & "' and normal='" & Mid(txtsn.Text, 12, 3) & "'  and  firstno<=" & Right(txtsn.Text, 6) & " and endno>=" & Right(txtsn.Text, 6) & "   "
'      Set rec = conn.Execute(sql)
'      If rec.EOF = False Then
'         Text1.Text = rec.Fields(3)
'      End If
If mfgH3C.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫɾ������!", vbInformation + vbOKOnly, "δѡ����"
      Exit Sub
   End If
      Dim response
response = MsgBox("ȷ��Ҫɾ�������ϣ�", vbYesNo, "delete")
If response = vbNo Then
     Exit Sub
End If
   sql = "delete from revset where model='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & "' and firstall='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 3) & "' and endall='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 4) & "' and ver='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 2) & "'"
   conn.Execute sql
   Call cmdQuery_Click

End Sub

Private Sub cmdInsert_Click()
txtsn.Enabled = True
txtModel.Enabled = True
txtpass.Enabled = True
txtVer.Enabled = True
txtsn.BackColor = &HFFFFFF
txtModel.BackColor = &HFFFFFF
txtpass.BackColor = &HFFFFFF
txtVer.BackColor = &HFFFFFF
txtsn.Text = ""
txtModel.Text = ""
txtpass.Text = ""
txtVer.Text = ""
op = "add"
End Sub

Private Sub cmdQuery_Click()
 If rec.State = 1 Then
   rec.Close
End If
sql = "select model,ver,firstall,endall from revset "
rec.Open sql, conn, adOpenKeyset, adLockOptimistic
Set mfgH3C.DataSource = rec
End Sub

Private Sub Form_Load()
  golPath = Connect.getConnectionstring
  conn.ConnectionString = golPath
  conn.Open
  Call cmdQuery_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   conn.Close
   Set conn = Nothing
End Sub

Private Sub mfgH3C_Click()
 If mfgH3C.RowSel > 0 Then
      txtsn.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 1)
      txtModel.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 3)
        txtpass.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 4)
        txtVer.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 2)
  End If
End Sub

Private Sub txtModel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      If txtModel.Text <> "" Then
         txtpass.SetFocus
      End If
   End If
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtpass.Text <> "" Then
         txtVer.SetFocus
      End If
   End If
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtsn.Text <> "" Then
         txtModel.SetFocus
      End If
   End If
End Sub
