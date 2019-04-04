VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Main Form"
   ClientHeight    =   6780
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
   ScaleHeight     =   6780
   ScaleWidth      =   12105
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command6 
      Caption         =   "新NEC标签打印"
      Height          =   735
      Left            =   9120
      TabIndex        =   16
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "大华标签打印"
      Height          =   615
      Left            =   6480
      TabIndex        =   15
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "海康发货"
      Height          =   615
      Left            =   3960
      TabIndex        =   14
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UNIS发货打印"
      Height          =   495
      Left            =   6480
      TabIndex        =   13
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdHPnew 
      Caption         =   "HP发货(新）"
      Height          =   615
      Left            =   3960
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "新H3C 标签打印"
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdDoubleSNHP 
      Caption         =   "HP双SN发货"
      Height          =   615
      Left            =   6480
      TabIndex        =   10
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "宇视标签打印"
      Height          =   615
      Left            =   -480
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdNEC 
      Caption         =   "NEC标签打印"
      Height          =   735
      Left            =   6480
      TabIndex        =   8
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdHP 
      Caption         =   "HP发货"
      Height          =   615
      Left            =   3960
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmd3COM 
      Caption         =   "3COM 标签打印"
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C 
      Caption         =   "H3C 标签打印"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdHUAWEI 
      Caption         =   "HUAWEI 标签打印"
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C_3COM 
      Caption         =   "H3C-3COM 标签打印"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退  出(Exit)"
      Height          =   735
      Left            =   6720
      TabIndex        =   1
      Top             =   5880
      Width           =   1575
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
      Top             =   6480
      Width           =   1695
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
      Caption         =   "标签打印选择(Label Printed Select)"
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

