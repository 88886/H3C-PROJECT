VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmQueryVersion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query Version"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQueryVersion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9825
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(Delete)"
      Height          =   615
      Left            =   3960
      TabIndex        =   14
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修改(Update)"
      Height          =   615
      Left            =   2040
      TabIndex        =   13
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   7920
      TabIndex        =   4
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查询(Query)"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   7560
      Width           =   1815
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgVersion 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6588
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fmVersion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   9735
      Begin VB.TextBox txtTestTime 
         Height          =   450
         Left            =   3480
         TabIndex        =   16
         Top             =   2640
         Width           =   4095
      End
      Begin VB.TextBox txtOperator 
         Height          =   495
         Left            =   3480
         TabIndex        =   12
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtVersion 
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txtMName 
         Height          =   495
         Left            =   3480
         TabIndex        =   9
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtSN 
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblTestTime 
         Caption         =   "测试时间(Test Time):"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label lblOperator 
         Caption         =   "操作人(Operator):"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label lblVersion 
         Caption         =   "版本信息(Version):"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label lblMName 
         Caption         =   "机种名称(Model Name):"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblSN 
         Caption         =   "产品序号(Serial Number):"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "版本信息查询"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "frmQueryVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String

Private Sub cmdDelete_Click()
   If mfgVersion.RowSel <= 0 Then
      MsgBox "请选择要删除的行!"
      Exit Sub
   End If
   sql = "Delete from Version where SN='" & mfgVersion.TextMatrix(mfgVersion.RowSel, 2) & "' and Model='" & mfgVersion.TextMatrix(mfgVersion.RowSel, 1) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
       MsgBox "删除版本资料失败!" & "原因是" & status
   End If
   MsgBox "删除版本资料成功!"
   renovate
End Sub

Private Sub cmdQuery_Click()
   Dim rcd As New ADODB.Recordset
   sql = "select * from Version where 1=1"
   If txtSN.Text <> "" Then
      sql = sql & " and SN like '" & txtSN.Text & "%'"
   End If
   If txtMName.Text <> "" Then
      sql = sql & " and Model='" & txtMName.Text & "'"
   End If
   If txtVersion.Text <> "" Then
      sql = sql & " and Ver='" & txtVersion.Text & "'"
   End If
   If txtOperator.Text <> "" Then
      sql = sql & " and Operator='" & txtOperator.Text & "'"
   End If
   If txtTestTime.Text <> "" Then
      sql = sql & " and TestTime Like '" & txtTestTime.Text & "%'"
   End If
   sql = sql & " order by Model,SN"
   rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgVersion.DataSource = rcd
   With mfgVersion
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 3500
        .ColWidth(3) = 2000
        .ColWidth(4) = 1500
        .ColWidth(5) = 3000
        
        .TextMatrix(0, 1) = "机种名称(Model Name)"
        .TextMatrix(0, 2) = "产品序号(Serial Number)"
        .TextMatrix(0, 3) = "版本(Version)"
        .TextMatrix(0, 4) = "操作人(Operator)"
        .TextMatrix(0, 5) = "操作时间(TestTime)"
   End With
   rcd.Close
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   If mfgVersion.RowSel <= 0 Then
      MsgBox "请选择要修改的行!"
      Exit Sub
   End If
   If txtSN.Text = "" Then
      MsgBox "产品序号不能为空!"
      txtSN.SetFocus
      Exit Sub
   Else
      If Len(txtSN.Text) < 10 Then
         MsgBox "产品序号长度不能小于10"
         txtSN.SetFocus
         Exit Sub
      End If
   End If
   If txtMName.Text <> Mid(txtSN.Text, 3, 8) Then
      MsgBox "机种名称不能对应产品序号!"
      txtMName.SetFocus
      Exit Sub
   End If
   If txtVersion.Text = "" Then
      MsgBox "版本不能为空!"
      txtVersion.SetFocus
      Exit Sub
   End If
   If txtTestTime.Text = "" Then
      MsgBox "测试时间不能为空!"
      txtTestTime.SetFocus
      Exit Sub
   End If
   sql = "Update Version set Ver='" & txtVersion.Text & "',Operator='" & txtOperator.Text & "',TestTime='" & txtTestTime.Text & "' where SN='" & mfgVersion.TextMatrix(mfgVersion.RowSel, 2) & "' and Model='" & txtMName.Text & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
       MsgBox "修改版本资料失败!" & "原因是" & status
   End If
   MsgBox "修改版本资料成功!"
   renovate
End Sub

Private Sub Form_Load()
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   renovate
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub

Private Sub renovate()
   sql = "select * from Version order by Model,SN"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgVersion.DataSource = rec
   With mfgVersion
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 3500
        .ColWidth(3) = 2000
        .ColWidth(4) = 1500
        .ColWidth(5) = 3000
        
        .TextMatrix(0, 1) = "机种名称(Model Name)"
        .TextMatrix(0, 2) = "产品序号(Serial Number)"
        .TextMatrix(0, 3) = "版本(Version)"
        .TextMatrix(0, 4) = "操作人(Operator)"
        .TextMatrix(0, 5) = "操作时间(TestTime)"
   End With
   rec.Close
End Sub

Private Sub mfgVersion_Click()
   If mfgVersion.RowSel > 0 Then
      txtMName.Text = mfgVersion.TextMatrix(mfgVersion.RowSel, 1)
      txtSN.Text = mfgVersion.TextMatrix(mfgVersion.RowSel, 2)
      txtVersion.Text = mfgVersion.TextMatrix(mfgVersion.RowSel, 3)
      txtOperator.Text = mfgVersion.TextMatrix(mfgVersion.RowSel, 4)
      txtTestTime.Text = mfgVersion.TextMatrix(mfgVersion.RowSel, 5)
   End If
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(txtSN.Text) < 10 Then
         MsgBox "产品序号长度不能少于10"
         txtSN.SetFocus
         Exit Sub
      Else
        txtMName.Text = Mid(txtSN.Text, 3, 8)
      End If
   End If
End Sub
