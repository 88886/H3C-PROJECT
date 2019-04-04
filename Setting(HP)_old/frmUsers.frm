VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users Mangaer"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8445
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Retrun)"
      Height          =   615
      Left            =   6120
      TabIndex        =   17
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   3360
      TabIndex        =   16
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "确定(Confirm)"
      Height          =   615
      Left            =   480
      TabIndex        =   15
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(Delete)"
      Height          =   615
      Left            =   6120
      TabIndex        =   14
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修改(Update)"
      Height          =   615
      Left            =   3360
      TabIndex        =   13
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "新增(Insert)"
      Height          =   615
      Left            =   480
      TabIndex        =   12
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Frame fmUser 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8295
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtConfirmPwd 
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
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txtUserPwd 
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
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblConfirmPwd 
         Caption         =   "确认密码(Confirm P):"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label lblUserPwd 
         Caption         =   "用户密码(Password):"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblUserName 
         Caption         =   "用户名称(User Name):"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "用户管理(Users Mangaer)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim r As Integer
Dim total As Integer
Dim op As String

Private Sub enable()         '
   txtUserName.Enabled = True
   txtUserName.BackColor = &HFFFFFF
   txtUserPwd.Enabled = True
   txtUserPwd.BackColor = &HFFFFFF
   txtConfirmPwd.Enabled = True
   txtConfirmPwd.BackColor = &HFFFFFF
   cmdInsert.Enabled = False
   cmdUpdate.Enabled = False
   cmdDelete.Enabled = False
   cmdConfirm.Enabled = True
   cmdCancel.Enabled = True
End Sub

Private Sub unable()
   txtUserName.Enabled = False
   txtUserName.BackColor = &HE0E0E0
   txtUserPwd.Enabled = False
   txtUserPwd.BackColor = &HE0E0E0
   txtConfirmPwd.Enabled = False
   txtConfirmPwd.BackColor = &HE0E0E0
   cmdInsert.Enabled = True
   cmdUpdate.Enabled = True
   cmdDelete.Enabled = True
   cmdConfirm.Enabled = False
   cmdCancel.Enabled = False
End Sub

Private Sub cmdCancel_Click()
   unable
   op = ""
End Sub

Private Sub cmdConfirm_Click()
   If txtUserName.Text = "" Then
      MsgBox "用户名称不能为空!"
      txtUserName.SetFocus
      Exit Sub
   End If
   If txtConfirmPwd.Text <> txtUserPwd.Text Then
      MsgBox "确认密码与用户密码不一致!"
      txtConfirmPwd.SetFocus
      Exit Sub
   End If
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from Users where UserName='" & txtUserName.Text & "'"
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "用户名称已存在!"
         txtUserName.SetFocus
         Exit Sub
      End If
      rcd.Close
      sql = "Insert into Users(UserID,UserName,Password) Values(" & getmaxUserID & ",'" & txtUserName.Text & "','" & txtUserPwd.Text & "')"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "新增用户失败!" & "原因是" & status
      End If
      MsgBox "新增用户成功!"
      renovate
      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update Users set Password='" & txtUserPwd.Text & "' where userid=" & rec.Fields(0) & " and UserName='" & txtUserName.Text & "'"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "修改用户失败!" & "原因是" & status
      End If
      MsgBox "修改用户成功!"
      renovate
      cmdCancel_Click
   End If
End Sub

Private Sub cmdDelete_Click()
   If MsgBox("确定要删除当前用户吗?", vbExclamation + vbOKCancel, "删除用户") = vbOK Then
      sql = "delete from users where userid='" & rec.Fields(0) & "' and UserName='" & rec.Fields(1) & "'"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "删除用户失败!" & "原因是" & status
      End If
      MsgBox "删除用户成功!"
      renovate
   End If
End Sub

Private Sub cmdFirst_Click()
   rec.MoveFirst
   r = 1
   cmdNext.Enabled = True
   cmdLast.Enabled = True
   txtUserName.Text = chknull(rec.Fields(1), "")
   txtUserPwd.Text = chknull(rec.Fields(2), "")
   txtConfirmPwd.Text = chknull(rec.Fields(2), "")
   cmdFirst.Enabled = False
   cmdPrevious.Enabled = False
End Sub

Private Sub cmdInsert_Click()
   enable
   txtUserName.Text = ""
   txtUserPwd.Text = ""
   txtConfirmPwd.Text = ""
   op = "Insert"
End Sub

Private Sub cmdLast_Click()
   rec.MoveLast
   r = rec.RecordCount
   cmdFirst.Enabled = True
   cmdPrevious.Enabled = True
   txtUserName.Text = chknull(rec.Fields(1), "")
   txtUserPwd.Text = chknull(rec.Fields(2), "")
   txtConfirmPwd.Text = chknull(rec.Fields(2), "")
   cmdNext.Enabled = False
   cmdLast.Enabled = False
End Sub

Private Sub cmdNext_Click()
   rec.MoveNext
   r = r + 1
   cmdFirst.Enabled = True
   cmdPrevious.Enabled = True
   If r = total Then
      cmdNext.Enabled = False
      cmdLast.Enabled = False
   End If
   txtUserName.Text = chknull(rec.Fields(1), "")
   txtUserPwd.Text = chknull(rec.Fields(2), "")
   txtConfirmPwd.Text = chknull(rec.Fields(2), "")
End Sub

Private Sub cmdPrevious_Click()
   rec.MovePrevious
   r = r - 1
   cmdNext.Enabled = True
   cmdLast.Enabled = True
   If r = 1 Then
      cmdFirst.Enabled = False
      cmdPrevious.Enabled = False
   End If
   txtUserName.Text = chknull(rec.Fields(1), "")
   txtUserPwd.Text = chknull(rec.Fields(2), "")
   txtConfirmPwd.Text = chknull(rec.Fields(2), "")
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
  enable
  txtUserName.Enabled = False
  txtUserName.BackColor = &HE0E0E0
  op = "Update"
End Sub

Private Sub Form_Load()
   unable
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   renovate
End Sub

Private Sub renovate()
   sql = "select count(*) from Users"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   total = rec.Fields(0)
   rec.Close
   sql = "select * from Users"
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   If rec.EOF = False Then
      txtUserName.Text = rec.Fields(1)
      txtUserPwd.Text = rec.Fields(2)
      txtConfirmPwd.Text = rec.Fields(2)
      cmdFirst.Enabled = False
      cmdPrevious.Enabled = False
      r = 1
      If rec.RecordCount = 1 Then
         cmdNext.Enabled = False
         cmdLast.Enabled = False
      Else
         cmdNext.Enabled = True
         cmdLast.Enabled = True
      End If
   Else
      cmdFirst.Enabled = False
      cmdPrevious.Enabled = False
      cmdNext.Enabled = False
      cmdLast.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub
