VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Label Manager"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9270
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdExit 
      Caption         =   "�뿪��ǩ����ϵͳ(Exit System)"
      Height          =   855
      Left            =   3000
      TabIndex        =   4
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton cmdVersion 
      Caption         =   "�汾��Ϣ��ѯ(Version Query )"
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton cmdVariables 
      Caption         =   "��ǩ�����趨(Variables Setting)"
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdUsers 
      Caption         =   "�û����Ϲ���(Users Manager)"
      Height          =   855
      Left            =   3000
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "��ǩ����(Label Manager)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
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
      Width           =   9255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   End
End Sub

Private Sub cmdUsers_Click()
   frmUsers.Show
End Sub

Private Sub cmdVariables_Click()
   frmVariables.Show
End Sub

Private Sub cmdVersion_Click()
   frmQueryVersion.Show
End Sub
