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
   Begin VB.CommandButton cmdModelVer 
      Caption         =   "发货标签软件  版本维护"
      Height          =   735
      Left            =   6600
      TabIndex        =   9
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdECO 
      Caption         =   "条码ECO版本 防呆"
      Height          =   735
      Left            =   6600
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdCustomType 
      Caption         =   "整机序列号品牌维护"
      Height          =   735
      Left            =   6600
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdEAN 
      Caption         =   "EAN变量设定"
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返    回(Return)"
      Height          =   735
      Left            =   6840
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdH3C_3COM 
      Caption         =   "H3C-3COM 变量设定"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdHUAWEI 
      Caption         =   "非3COM类     变量设定"
      Height          =   735
      Left            =   3960
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C 
      Caption         =   "H3C 变量设定"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmd3COM 
      Caption         =   "3COM 变量设定"
      Height          =   735
      Left            =   3960
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
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
   Begin VB.Image imgH3C_3COM 
      Height          =   1785
      Left            =   1680
      Picture         =   "frmVariables.frx":2E1A
      Top             =   4560
      Width           =   1785
   End
   Begin VB.Image imgHUAWEI 
      Height          =   1185
      Left            =   1680
      Picture         =   "frmVariables.frx":9F63
      Top             =   3120
      Width           =   1740
   End
   Begin VB.Image imgH3C 
      Height          =   810
      Left            =   1680
      Picture         =   "frmVariables.frx":AB49
      Top             =   2160
      Width           =   1785
   End
   Begin VB.Image img3COM 
      Enabled         =   0   'False
      Height          =   1245
      Left            =   1680
      Picture         =   "frmVariables.frx":B59C
      Top             =   600
      Width           =   1725
   End
End
Attribute VB_Name = "frmVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd3COM_Click()
    frmH3C_3COMSetting.Show 1
End Sub

Private Sub cmdCustomType_Click()
    frmCustomType.Show 1
End Sub

Private Sub cmdEAN_Click()
    frmEANSetting.Show 1
End Sub

Private Sub cmdECO_Click()
    'frmECO.Show 1
    frmECO_Enable.Show 1
End Sub

Private Sub cmdH3C_3COM_Click()
   frmH3C_3COMSetting.Show 1
End Sub

Private Sub cmdH3C_Click()
   frmH3CSetting.Show 1
End Sub

Private Sub cmdHUAWEI_Click()
   frmHUAWEISetting.Show 1
End Sub

Private Sub cmdModelVer_Click()
    frmModelVer.Show 1
End Sub

Private Sub cmdReturn_Click()
   Unload Me
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
