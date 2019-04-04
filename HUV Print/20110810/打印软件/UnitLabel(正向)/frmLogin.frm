VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户登录(User Login)"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmLongin 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton cmdOK 
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         Picture         =   "frmLogin.frx":073E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton CmdCancel 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         Picture         =   "frmLogin.frx":1020
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblUser 
         Caption         =   "用户名(User Name):"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblPassward 
         Caption         =   "密    码(Password):"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conn As New ADODB.Connection
Private Sub cmdOK_Click()


  Dim sUserName As String
  Dim sPWD As String
  sPWD = Trim(txtPassword.Text)
  sUserName = Trim(txtUserName.Text)
  
    Dim recBeg As New ADODB.Recordset
    sql = "select * from tblUNIT_UserLog where UserName='" & sUserName & "' "
    recBeg.Open sql, conn, adOpenKeyset, adLockOptimistic
    If recBeg.EOF = False Then
        MsgBox "此账号正在使用中!"
        txtPassword.Text = ""
        txtUserName.Text = ""
        Exit Sub
    
    End If
  
  Dim sSQL As String
  sSQL = "select * from users where username='" & sUserName & "' and Password='" & sPWD & "' "
  Dim rsResult As Recordset
  Set rsResult = conn.Execute(sSQL)
  If rsResult.EOF = True Then
    MsgBox "用户名密码不正确!"
    Exit Sub
  End If
  golUSERID = Trim(rsResult.Fields(0).Value)
  golUSERNAME = Trim(rsResult.Fields(1).Value)
  Main_Scan_SN.Visible = True
  conn.Close
  Set conn = Nothing
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  golPath = Connect.getConnectionstring
  conn.ConnectionString = golPath
  conn.Open
End Sub


