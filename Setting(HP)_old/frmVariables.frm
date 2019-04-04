VERSION 5.00
Begin VB.Form frmVariables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Variables Setting"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVariables.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9285
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd3COM 
      Caption         =   "3COM和NEC电源代码设定"
      Height          =   735
      Left            =   1560
      TabIndex        =   10
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdH3C 
      Caption         =   "H3C 变量设定"
      Height          =   735
      Left            =   1560
      TabIndex        =   9
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdHUAWEI 
      BackColor       =   &H00404040&
      Caption         =   "HUAWEI 变量设定"
      Height          =   735
      Left            =   1560
      TabIndex        =   8
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton cmdH3C_3COM 
      Caption         =   "H3C-3COM 变量设定"
      Height          =   735
      Left            =   1560
      TabIndex        =   7
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton bthHP 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HP变量设定"
      Height          =   735
      Left            =   5040
      TabIndex        =   6
      Top             =   2400
      Width           =   2145
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "HUV 变量设定"
      Height          =   735
      Left            =   5040
      TabIndex        =   5
      Top             =   1080
      Width           =   2145
   End
   Begin VB.CommandButton cmdH3C2D 
      Caption         =   "H3C 2D 打印变量设定"
      Height          =   735
      Left            =   5040
      TabIndex        =   4
      Top             =   3720
      Width           =   2145
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "新H3C变量设定"
      Height          =   735
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   2145
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返    回(Return)"
      Height          =   735
      Left            =   7560
      TabIndex        =   1
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "2013.09.03"
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
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "标签变量设定(Variables Setting)"
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
      Width           =   9255
   End
End
Attribute VB_Name = "frmVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bthHP_Click()
  frmHPSetting.Show 1
End Sub

Private Sub cmd3COM_Click()
    frmOthers.Show 1
End Sub

Private Sub cmdH3C_3COM_Click()
   frmH3C_3COMSetting.Show 1
End Sub

Private Sub cmdH3C_Click()
   frmH3CSetting.Show 1
End Sub

Private Sub cmdH3C2D_Click()
    frmH3C2DSetting.Show 1
End Sub

Private Sub cmdHUAWEI_Click()
   frmHUAWEISetting.Show 1
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   frmHUVSetting.Show 1
End Sub

Private Sub Command2_Click()
    frmNewH3CSetting.Show 1
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


