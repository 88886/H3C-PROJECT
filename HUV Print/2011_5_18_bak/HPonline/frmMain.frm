VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Main Form"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8730
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdHP 
      Caption         =   "HP����"
      Height          =   615
      Left            =   3960
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmd3COM 
      Caption         =   "3COM ��ǩ��ӡ"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C 
      Caption         =   "H3C ��ǩ��ӡ"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdHUAWEI 
      Caption         =   "HUAWEI ��ǩ��ӡ"
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C_3COM 
      Caption         =   "H3C-3COM ��ǩ��ӡ"
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "��  ��(Exit)"
      Height          =   735
      Left            =   6840
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "2.28"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6480
      Width           =   615
   End
   Begin VB.Image img3COM 
      Enabled         =   0   'False
      Height          =   1245
      Left            =   1680
      Picture         =   "frmMain.frx":13652
      Top             =   600
      Width           =   1725
   End
   Begin VB.Image imgH3C 
      Height          =   810
      Left            =   1680
      Picture         =   "frmMain.frx":14246
      Top             =   2160
      Width           =   1785
   End
   Begin VB.Image imgHUAWEI 
      Height          =   1185
      Left            =   1680
      Picture         =   "frmMain.frx":14C99
      Top             =   3120
      Width           =   1740
   End
   Begin VB.Image imgH3C_3COM 
      Height          =   1785
      Left            =   1680
      Picture         =   "frmMain.frx":1587F
      Top             =   4560
      Width           =   1785
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "��ǩ��ӡѡ��(Label Printed Select)"
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
      TabIndex        =   5
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd3COM_Click()
   '����
End Sub

Private Sub cmdExit_Click()
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   If conn1.State = 1 Then
      conn1.Close
   Set conn1 = Nothing
   End If
   End
End Sub

Private Sub cmdH3C_3COM_Click()
   frmH3C_3COMPrint.Show
End Sub

Private Sub cmdH3C_Click()
   frmH3CPrint.Show
End Sub

Private Sub cmdHP_Click()
    frmChunHP.Show
End Sub

Private Sub cmdHUAWEI_Click()
   frmHUAWEIPrint.Show
End Sub



Private Sub img3COM_Click()
   cmd3COM_Click
End Sub

Private Sub imgH3C_3COM_Click()
   cmdH3C_3COM_Click
End Sub

Private Sub imgH3C_Click()
   cmdH3C_Click
End Sub

Private Sub imgHUAWEI_Click()
   cmdHUAWEI_Click
End Sub

