VERSION 5.00
Begin VB.Form FrmNewH3C 
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17460
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   17460
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      TabIndex        =   6
      Top             =   4680
      Width           =   17295
      Begin VB.TextBox txtVer 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7920
         TabIndex        =   30
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1800
         TabIndex        =   29
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   1800
         TabIndex        =   28
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtEPN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7920
         TabIndex        =   27
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtDes 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   1800
         TabIndex        =   26
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtOS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtGW 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7920
         TabIndex        =   24
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtHV 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   1800
         TabIndex        =   23
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox txtMS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   1800
         TabIndex        =   22
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtNAL 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7920
         TabIndex        =   21
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7920
         TabIndex        =   20
         Top             =   3720
         Width           =   3135
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   19
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox chkChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8160
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无CE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   2160
         Width           =   855
      End
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "外尺寸(MM):"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkTurkey 
         BackColor       =   &H80000005&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8160
         TabIndex        =   12
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox chkNonTurkey 
         BackColor       =   &H80000005&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         TabIndex        =   11
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chkVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本信息:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8160
         TabIndex        =   9
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   495
         Left            =   9600
         TabIndex        =   8
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox chkCEAddr 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE Addr"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品描述:"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(中文):"
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(英文):"
         Height          =   375
         Left            =   5760
         TabIndex        =   41
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblGW 
         BackColor       =   &H00FFFFFF&
         Caption         =   "毛重(kg):"
         Height          =   375
         Left            =   5760
         TabIndex        =   40
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息CE:"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息WEEE:"
         Height          =   375
         Left            =   5760
         TabIndex        =   38
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblHV 
         BackColor       =   &H00FFFFFF&
         Caption         =   "硬件版本:"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息ChinaRoHS:"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "执行标准:"
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网许可号:"
         Height          =   375
         Left            =   5760
         TabIndex        =   34
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
         Height          =   375
         Left            =   5760
         TabIndex        =   33
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label lblRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息RoHS:"
         Height          =   495
         Left            =   5760
         TabIndex        =   32
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblTurkeyRohs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "土耳其RoHs:"
         Height          =   375
         Left            =   5760
         TabIndex        =   31
         Top             =   1680
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   6360
      TabIndex        =   5
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   8880
      TabIndex        =   4
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   9720
      Width           =   1815
   End
   Begin VB.TextBox lblMSday 
      Height          =   450
      Left            =   120
      TabIndex        =   2
      Top             =   9480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox lblNALday 
      Height          =   450
      Left            =   480
      TabIndex        =   1
      Top             =   9480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   105
      Picture         =   "FrmNewH3C.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   21315
      TabIndex        =   0
      Top             =   0
      Width           =   21375
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "H3C 标签："
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "FrmNewH3C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
