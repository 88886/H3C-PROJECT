VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVerset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设定机种对应版本(Setting Rev)"
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
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExport 
      Caption         =   "逆向导入"
      Height          =   495
      Left            =   8520
      TabIndex        =   15
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdNewQuery 
      Caption         =   "查询(Query)"
      Height          =   735
      Left            =   6720
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
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
      Caption         =   "查询(Query)"
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "新增(Insert)"
      Height          =   735
      Left            =   6720
      TabIndex        =   3
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(Delete)"
      Height          =   735
      Left            =   6720
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
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
      Left            =   8520
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "确定(Confirm)"
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
      Left            =   8520
      TabIndex        =   5
      Top             =   4200
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
      Caption         =   "版本（Rev）:"
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
      Caption         =   "流水号结束值:"
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
      Caption         =   "流水号起始值:"
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
      Caption         =   "机               种:"
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
   
   txtsn.Enabled = False
   txtModel.Enabled = False
   txtVer.Enabled = False
   txtpass.Enabled = False
   
   txtsn.BackColor = &HE0E0E0
   txtModel.BackColor = &HE0E0E0
   txtVer.BackColor = &HE0E0E0
   txtpass.BackColor = &HE0E0E0
   op = ""
End Sub

Private Sub cmdConfirm_Click()
If op = "add" Then
   If txtsn.Text = "" Then
       MsgBox "机种不能为空!", vbOKOnly + vbExclamation, "Error"
       txtsn.SetFocus
       Exit Sub
   End If

   If txtModel.Text = "" Then
      MsgBox "流水号起始值不能为空!", vbOKOnly + vbExclamation, "Error"
      txtModel.SetFocus
      Exit Sub
   End If
   If txtpass.Text = "" Then
      MsgBox "流水号结束值不能为空!", vbOKOnly + vbExclamation, "Error"
      txtpass.SetFocus
      Exit Sub
   End If
   If Len(txtModel.Text) <> 20 Then
        MsgBox "流水号起始值长度不对!", vbExclamation + vbOKOnly, "Error"
        txtModel.SetFocus
        Exit Sub
   End If
    If Len(txtpass.Text) <> 20 Then
        MsgBox "流水号结束值长度不对!", vbExclamation + vbOKOnly, "Error"
        txtpass.SetFocus
        Exit Sub
   End If
   If Val(Right(txtpass.Text, 6)) < Val(Right(txtModel.Text, 6)) Then
       MsgBox "流水号结束值不能小于起始值!", vbOKOnly + vbExclamation, "Error"
      txtpass.SetFocus
      Exit Sub
   End If
   If txtVer.Text = "" Then
      MsgBox "版本不能为空!", vbOKOnly + vbExclamation, "Error"
      txtVer.SetFocus
      Exit Sub
   End If

   sql = "select * from revset  where model='" & txtsn.Text & "' and firstall='" & txtModel.Text & "' and endall='" & txtpass.Text & "' and ver='" & txtVer.Text & "'"
    If rec.State = 1 Then
        rec.Close
    End If
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    If rec.EOF = False Then
       MsgBox "此机种已存在!", vbOKOnly + vbExclamation, "Error"
       Exit Sub
    End If
    
    sql = "insert into revset(model,firstno,endno,ver,firstall,endall,normal) values('" & txtsn.Text & "'," & Right(txtModel.Text, 6) & "," & Right(txtpass.Text, 6) & ",'" & txtVer.Text & "','" & txtModel.Text & "','" & txtpass.Text & "','" & Mid(txtpass.Text, 12, 3) & "')"
    
    conn.Execute sql
End If
   
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
      MsgBox "请选择要删除的行!", vbInformation + vbOKOnly, "未选择行"
      Exit Sub
   End If
      Dim response
response = MsgBox("确定要删除此资料？", vbYesNo, "delete")
If response = vbNo Then
     Exit Sub
End If
   sql = "delete from revset where model='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & "' and firstall='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 3) & "' and endall='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 4) & "' and ver='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 2) & "'"
   conn.Execute sql
   Call cmdQuery_Click

   txtsn.Text = ""
   txtModel.Text = ""
   txtVer.Text = ""
   txtpass.Text = ""
End Sub

Private Sub cmdExport_Click()
    Dim xlConn As New ADODB.Connection
    Dim xlRs As New ADODB.Recordset
    Dim model As String
    Dim sn As String
    Dim ver As String
    Dim strConn As String
    Dim xlCnt As Integer
    'If conn.State = 0 Then
    '    conn.ConnectionString = Connect.getConnectionstring
    '    conn.Open
    'End If
    sql = "delete from tblNiXiangExport"
    conn.Execute sql
    'conn.Close
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\nixiang.xls;Extended Properties='Excel 8.0;HDR=yes;IMEX=1'"
    xlConn.Open strConn
    xlRs.Open "select MODEL,SN,VER from [sheet1$]", xlConn, adOpenStatic, adLockReadOnly
    xlCnt = xlRs.RecordCount
    If xlCnt < 1 Then
        MsgBox ("nixiang.xls文件中无数据！")
        xlConn.Close
        Exit Sub
    Else
        For i = 1 To xlCnt
            If IsNull(xlRs("sn")) <> True Then
                model = xlRs("MODEl")
                sn = xlRs("SN")
                ver = xlRs("VER")

                'If conn.State = 0 Then
                'conn.ConnectionString = Connect.getConnectionstring
                'conn.Open
                'End If
                
                sql = "insert tblNiXiangExport select '" & UCase(Trim(model)) & "','" & UCase(Trim(sn)) & "','" & UCase(Trim(ver)) & "'"
                conn.Execute sql
                'conn.Close
            End If
            xlRs.MoveNext
        Next
        xlConn.Close
        
        If rec.State = 1 Then
            rec.Close
        End If
        Dim strmodel As String
        sql = "select distinct model from tblNiXiangExport "
        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
        If rec.RecordCount > 1 Then
           MsgBox "一次只能维护一个机种!"
           rec.Close
           Exit Sub
        Else
           strmodel = rec.Fields(0)
        End If
        
        If rec.State = 1 Then
            rec.Close
        End If
        sql = "select distinct ver from tblNiXiangExport "
        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
        If rec.RecordCount > 1 Then
            MsgBox "只能维护同一版本!"
            rec.Close
            Exit Sub
        End If
        
        If rec.State = 1 Then
            rec.Close
        End If
        sql = "select sn from tblNiXiangExport where substring(sn,3,8)<>'" & strmodel & "' "
        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
        If rec.RecordCount >= 1 Then
            MsgBox "流水号与机种不对应!"
            rec.Close
            Exit Sub
        End If
        
        Dim Conn2 As New ADODB.Connection
        If Conn2.State = 0 Then
            Conn2.ConnectionString = Connect.getConnectionstring
            Conn2.Open
        End If
        
        Dim str As String
        If rec.State = 1 Then
            rec.Close
        End If
        sql = "select MODEL,SN,VER from tblNiXiangExport "
        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
        Do Until rec.EOF

            str = "Insert into revset(model,firstno,endno,ver,firstall,endall,normal) values('" & rec.Fields(0) & "'," & Right(rec.Fields(1), 6) & "," & Right(rec.Fields(1), 6) & ",'" & rec.Fields(2) & "','" & rec.Fields(1) & "','" & rec.Fields(1) & "','" & Mid(rec.Fields(1), 12, 3) & "')"
            Conn2.Execute str

            rec.MoveNext
        Loop
        
        Conn2.Close
        rec.Close
        
        
        
        MsgBox ("资料导入成功！")
    End If
    
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

Private Sub cmdNewQuery_Click()
    If rec.State = 1 Then
        rec.Close
    End If
    sql = "select model,ver,firstall,endall from revset where 1=1 "
    If txtsn.Text <> "" Then
        sql = sql & " and model ='" & Trim(txtsn.Text) & "'"
    End If
    If txtVer.Text <> "" Then
        sql = sql & " and ver ='" & Trim(txtVer.Text) & "'"
    End If
    
    If txtModel.Text <> "" Then
        sql = sql & " and firstall<='" & Trim(txtModel.Text) & "' and  endall>='" & Trim(txtModel.Text) & "'"
    End If
    If txtpass.Text <> "" Then
        sql = sql & " and firstall<='" & Trim(txtpass.Text) & "' and endall>='" & Trim(txtpass.Text) & "'"
    End If
    
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    Set mfgH3C.DataSource = rec
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
