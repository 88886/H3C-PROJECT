VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Main Form"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12105
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
   ScaleHeight     =   7905
   ScaleWidth      =   12105
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command7 
      Caption         =   "SRIE��ǩ��ӡ"
      Height          =   800
      Left            =   9120
      TabIndex        =   16
      Top             =   960
      Width           =   2275
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��NEC��ǩ��ӡ"
      Height          =   800
      Left            =   6480
      TabIndex        =   15
      Top             =   960
      Width           =   2275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�󻪱�ǩ��ӡ"
      Height          =   800
      Left            =   6480
      TabIndex        =   14
      Top             =   2040
      Width           =   2275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��������"
      Height          =   800
      Left            =   3840
      TabIndex        =   13
      Top             =   5280
      Width           =   2275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UNIS������ӡ"
      Height          =   800
      Left            =   6480
      TabIndex        =   12
      Top             =   4200
      Width           =   2275
   End
   Begin VB.CommandButton cmdHPnew 
      Caption         =   "HP����(�£�"
      Height          =   800
      Left            =   6480
      TabIndex        =   11
      Top             =   6360
      Visible         =   0   'False
      Width           =   2275
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "��H3C ��ǩ��ӡ"
      Height          =   800
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   2275
   End
   Begin VB.CommandButton cmdDoubleSNHP 
      Caption         =   "HP˫SN����"
      Height          =   800
      Left            =   6480
      TabIndex        =   9
      Top             =   5280
      Width           =   2275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ӱ�ǩ��ӡ"
      Height          =   800
      Left            =   -1200
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   2275
   End
   Begin VB.CommandButton cmdHP 
      Caption         =   "HP����"
      Height          =   800
      Left            =   3840
      TabIndex        =   7
      Top             =   4200
      Width           =   2275
   End
   Begin VB.CommandButton cmd3COM 
      Caption         =   "3COM ��ǩ��ӡ"
      Height          =   800
      Left            =   3840
      TabIndex        =   4
      Top             =   960
      Width           =   2275
   End
   Begin VB.CommandButton cmdH3C 
      Caption         =   "H3C ��ǩ��ӡ"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3840
      TabIndex        =   3
      Top             =   2040
      Width           =   2275
   End
   Begin VB.CommandButton cmdHUAWEI 
      Caption         =   "HUAWEI ��ǩ��ӡ"
      Height          =   800
      Left            =   3840
      TabIndex        =   0
      Top             =   3120
      Width           =   2275
   End
   Begin VB.CommandButton cmdH3C_3COM 
      Caption         =   "H3C-3COM ��ǩ��ӡ"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3840
      TabIndex        =   2
      Top             =   6360
      Width           =   2275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "��  ��(Exit)"
      Height          =   850
      Left            =   9840
      TabIndex        =   1
      Top             =   6840
      Width           =   2000
   End
   Begin VB.Label Label1 
      Caption         =   "2014.10.25 ME"
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
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Image img3COM 
      Enabled         =   0   'False
      Height          =   1245
      Left            =   1800
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
      Width           =   12135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd3COM_Click()
   frmHW3COM.Show
End Sub

Private Sub cmdDoubleSNHP_Click()
    frmDoubleSNHP.Show
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

Private Sub cmdHPnew_Click()
    frmChunHPnew.Show
End Sub

Private Sub cmdHUAWEI_Click()
   frmHUAWEIPrint.Show
End Sub

Private Sub cmdNEC_Click()
    frmNEC.Show
End Sub

Private Sub Command1_Click()
   frmHUVPrint.Show
End Sub

Private Sub Command2_Click()
    frmNewH3CPrint.Show
End Sub

Private Sub Command3_Click()
    frmNewUnisPrint.Show
End Sub

Private Sub Command4_Click()
    frmChunConsen.Show
End Sub

Private Sub Command5_Click()
    frmDaHua.Show
End Sub

Private Sub Command6_Click()
    frmNewNECPrint.Show
End Sub

Private Sub Command7_Click()
    frmSRIEPrint.Show
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

