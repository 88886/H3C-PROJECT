VERSION 5.00
Begin VB.Form frmUNISshenchan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UNIS 生产"
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   18615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUNISshenchan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11550
   ScaleWidth      =   18615
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      Picture         =   "frmUNISshenchan.frx":13652
      ScaleHeight     =   4545
      ScaleWidth      =   11865
      TabIndex        =   104
      Top             =   240
      Width           =   11895
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   16335
      Begin VB.CheckBox chkN4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N4"
         Height          =   375
         Left            =   5280
         TabIndex        =   105
         Top             =   5280
         Width           =   735
      End
      Begin VB.ComboBox txtSZ 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmUNISshenchan.frx":3F7C5
         Left            =   14520
         List            =   "frmUNISshenchan.frx":3F7CF
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   5280
         Width           =   1575
      End
      Begin VB.PictureBox Picture13 
         Height          =   495
         Left            =   10200
         Picture         =   "frmUNISshenchan.frx":3F7D9
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   101
         Top             =   5160
         Width           =   495
      End
      Begin VB.CheckBox chkNonEAC 
         BackColor       =   &H0000C000&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   12000
         TabIndex        =   100
         Top             =   5280
         Width           =   735
      End
      Begin VB.CheckBox chkEAC 
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11160
         TabIndex        =   99
         Top             =   5280
         Width           =   735
      End
      Begin VB.CheckBox chkHPEAddr 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HPE Addr"
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   98
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CheckBox chkY2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y2"
         Height          =   375
         Left            =   3280
         TabIndex        =   96
         Top             =   5280
         Width           =   735
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y*"
         Height          =   375
         Left            =   2280
         TabIndex        =   95
         Top             =   5280
         Width           =   735
      End
      Begin VB.CheckBox chkN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N*"
         Height          =   375
         Left            =   4280
         TabIndex        =   94
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox txtUkraineID 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   3840
         TabIndex        =   93
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CheckBox ChkEP 
         BackColor       =   &H00FFFFFF&
         Caption         =   "非环保"
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
         Left            =   5640
         TabIndex        =   92
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkUkraine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   56
         Top             =   4800
         Width           =   855
      End
      Begin VB.CheckBox chkNonUkraine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   3120
         TabIndex        =   55
         Top             =   4800
         Width           =   615
      End
      Begin VB.CheckBox chkSVPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "是"
         Enabled         =   0   'False
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
         Left            =   14880
         TabIndex        =   54
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox chkNonSVPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "否"
         Enabled         =   0   'False
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
         Left            =   15480
         TabIndex        =   53
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txt5000 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   7800
         TabIndex        =   52
         Top             =   4920
         Width           =   2175
      End
      Begin VB.TextBox txtNAL2Title 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7800
         TabIndex        =   51
         Top             =   4275
         Width           =   2175
      End
      Begin VB.TextBox txtNAL2 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7800
         TabIndex        =   50
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txtNAL1Title 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7800
         TabIndex        =   49
         Top             =   3195
         Width           =   2175
      End
      Begin VB.TextBox txtRCM 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   48
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CheckBox chkRCM 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   11160
         TabIndex        =   47
         Top             =   3600
         Width           =   855
      End
      Begin VB.CheckBox chkNonRCM 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   12000
         TabIndex        =   46
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtGost 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   45
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CheckBox chkGost 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   11160
         TabIndex        =   44
         Top             =   4200
         Width           =   855
      End
      Begin VB.CheckBox chkNonGost 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   12000
         TabIndex        =   43
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox txtKC 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   42
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CheckBox chkKC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   11160
         TabIndex        =   41
         Top             =   4800
         Width           =   855
      End
      Begin VB.CheckBox chkNonKC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   12000
         TabIndex        =   40
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox txtICT 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   39
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CheckBox chkICT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   11160
         TabIndex        =   38
         Top             =   3000
         Width           =   855
      End
      Begin VB.CheckBox chkNonICT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   12000
         TabIndex        =   37
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtCTick 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   36
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CheckBox chkCTick 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   11160
         TabIndex        =   35
         Top             =   2400
         Width           =   855
      End
      Begin VB.CheckBox chkNonCTick 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   12000
         TabIndex        =   34
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtATick 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   33
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox chkATick 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   11160
         TabIndex        =   32
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox chkNonATick 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   12000
         TabIndex        =   31
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox chkCEAddr 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE Addr"
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   30
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   29
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox chkRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   28
         Top             =   4320
         Width           =   855
      End
      Begin VB.CheckBox chkVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本信息:"
         Enabled         =   0   'False
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
         Left            =   7440
         TabIndex        =   27
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkNonTurkey 
         BackColor       =   &H80000005&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   26
         Top             =   3360
         Width           =   615
      End
      Begin VB.CheckBox chkTurkey 
         BackColor       =   &H80000005&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   25
         Top             =   3360
         Width           =   735
      End
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "外尺寸(MM):"
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE"
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   23
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无CE"
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   22
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox chkWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   21
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   20
         Top             =   3840
         Width           =   615
      End
      Begin VB.CheckBox chkChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   19
         Top             =   2880
         Width           =   855
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   18
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11880
         TabIndex        =   17
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtNAL1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7800
         TabIndex        =   16
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox txtMS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7800
         TabIndex        =   15
         Top             =   2130
         Width           =   2175
      End
      Begin VB.TextBox txtHV 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11280
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtGW 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7800
         TabIndex        =   13
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtOS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtProductID 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   11
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtEPN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7800
         TabIndex        =   10
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   9
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtSN 
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
         Height          =   405
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtVer 
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
         Height          =   405
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtQty 
         Height          =   405
         Left            =   6240
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtPC 
         Height          =   450
         Left            =   11880
         TabIndex        =   5
         Text            =   "N/A"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SZ:"
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
         Left            =   12960
         TabIndex        =   102
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "环保属性:"
         Height          =   375
         Left            =   480
         TabIndex        =   97
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ukraine:"
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
         Left            =   840
         TabIndex        =   91
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "软件版本打印"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13320
         TabIndex        =   90
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "5000米状态:"
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
         Left            =   5520
         TabIndex        =   89
         Top             =   4965
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "电源代码:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10800
         TabIndex        =   88
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网型号2:"
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
         Left            =   5520
         TabIndex        =   87
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网2:"
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
         Left            =   6000
         TabIndex        =   86
         Top             =   3750
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网型号1:"
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
         Left            =   5640
         TabIndex        =   85
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE:"
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
         Left            =   1080
         TabIndex        =   84
         Top             =   2325
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RCM ID:"
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
         Left            =   12960
         TabIndex        =   83
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RCM:"
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
         Left            =   10200
         TabIndex        =   82
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gost ID:"
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
         Left            =   12960
         TabIndex        =   81
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gost:"
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
         Left            =   10200
         TabIndex        =   80
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "KC ID:"
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
         Left            =   12960
         TabIndex        =   79
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "KC:"
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
         Left            =   10200
         TabIndex        =   78
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ICT ID:"
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
         Left            =   12960
         TabIndex        =   77
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ICT:"
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
         Left            =   10200
         TabIndex        =   76
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CTick ID:"
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
         Left            =   12960
         TabIndex        =   75
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CTick:"
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
         Left            =   10200
         TabIndex        =   74
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ATick ID:"
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
         Left            =   12960
         TabIndex        =   73
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ATick:"
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
         Left            =   10200
         TabIndex        =   72
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblTurkeyRohs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Turkey RoHs:"
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
         Left            =   240
         TabIndex        =   71
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lblRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RoHS:"
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
         Left            =   840
         TabIndex        =   70
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
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
         Left            =   11160
         TabIndex        =   69
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblNAL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网1:"
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
         Left            =   6120
         TabIndex        =   68
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "执行标准:"
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
         Left            =   5760
         TabIndex        =   67
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ChinaRoHS:"
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
         Left            =   240
         TabIndex        =   66
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblHV 
         BackColor       =   &H00FFFFFF&
         Caption         =   "硬件版本:"
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
         Left            =   10200
         TabIndex        =   65
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label lblWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WEEE:"
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
         Left            =   840
         TabIndex        =   64
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label lblCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE:"
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
         Left            =   1080
         TabIndex        =   63
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblGW 
         BackColor       =   &H00FFFFFF&
         Caption         =   "毛重(kg):"
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
         Left            =   5640
         TabIndex        =   62
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(英文):"
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
         Left            =   5640
         TabIndex        =   61
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(中文):"
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
         Left            =   240
         TabIndex        =   60
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品序列号:"
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
         Left            =   240
         TabIndex        =   59
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品ID:"
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
         Left            =   240
         TabIndex        =   58
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "数量:"
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
         Left            =   5520
         TabIndex        =   57
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Top             =   10920
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   9000
      TabIndex        =   2
      Top             =   10920
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   6480
      TabIndex        =   1
      Top             =   10920
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "UNIS 生产："
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmUNISshenchan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim mydoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects

Public Sub Reset()
    For Each ctr In Me.Controls
        If TypeOf ctr Is TextBox Then
                ctr.Text = ""
                ctr.Enabled = True
                ctr.BackColor = &HFFFFFF
             ElseIf TypeOf ctr Is CheckBox Then
                ctr.Value = 0
        End If
    Next
    Me.txtPC.Text = "N/A"
End Sub


Private Sub chkCE_Click()
   If chkCE.Value = 1 Then
      chkNonCE.Value = 0
   Else
      chkNonCE.Value = 1
   End If
End Sub

Private Sub chkEP_Click()
    If MsgBox("勾选非环保属性变更不能更改,确认请点击Yes,否则请点击No", vbYesNoCancel, "操作提醒") = vbYes Then
        If Me.chkEP.Value = 1 Then
            Me.chkEP.Enabled = False
            chkNonWEEE.Value = 1
            chkNonTurkey.Value = 1
            chkNonRoHS.Value = 1
            chkChinaRoHS.Value = 1
            chkNonCE.Value = 1
            chkCEAddr.Value = 0
            chkHPEAddr.Value = 0
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub chkNonCE_Click()
   If chkNonCE.Value = 1 Then
      chkCE.Value = 0
   Else
      chkCE.Value = 1
   End If
End Sub

Private Sub chkNonChinaRoHS_Click()
   If chkNonChinaRoHS.Value = 1 Then
      chkChinaRoHS.Value = 0
   Else
      chkChinaRoHS.Value = 1
   End If
End Sub

Private Sub chkNonRoHS_Click()
    If chkNonRoHS.Value = 1 Then
        chkRoHS.Value = 0
    End If
End Sub

Private Sub chkNonTurkey_Click()
    If chkNonTurkey.Value = 1 Then
      chkTurkey.Value = 0
   Else
      chkTurkey.Value = 1
   End If
End Sub

Private Sub chkRoHS_Click()
    If chkRoHS.Value = 1 Then
        chkNonRoHS.Value = 0
    End If
End Sub

Private Sub chkTurkey_Click()
    If chkTurkey.Value = 1 Then
      chkNonTurkey.Value = 0
    Else
      chkNonTurkey.Value = 1
    End If
End Sub

Private Sub chkNonWEEE_Click()
   If chkNonWEEE.Value = 1 Then
      chkWEEE.Value = 0
   Else
      chkWEEE.Value = 1
   End If
End Sub
Private Sub chkEAC_Click()
   If chkEAC.Value = 1 Then
      chkNonEAC.Value = 0
   Else
      chkNonEAC.Value = 1
   End If
End Sub

Private Sub chkNonEAC_Click()
   If chkNonEAC.Value = 1 Then
      chkEAC.Value = 0
   Else
      chkEAC.Value = 1
   End If
End Sub

Private Sub chkOS_Click()
   If chkOS.Value = 1 Then
      txtOS.Enabled = True
      txtOS.BackColor = &H80000005
   Else
      txtOS.Enabled = False
      txtOS.BackColor = &HC0C0C0
   End If
End Sub

Private Sub chkChinaRoHS_Click()
   If chkChinaRoHS.Value = 1 Then
      chkNonChinaRoHS.Value = 0
   Else
      chkNonChinaRoHS.Value = 1
   End If
End Sub


Private Sub chkWEEE_Click()
   If chkWEEE.Value = 1 Then
      chkNonWEEE.Value = 0
   Else
      chkNonWEEE.Value = 1
   End If
End Sub

Private Sub cmdCancel_Click()
   Reset
   txtSN.SetFocus
   Me.chkEP.Enabled = True
End Sub

Private Sub cmdPrint_Click()
'On Error Resume Next
   If Me.chkN.Value = 0 And Me.chkY.Value = 0 And Me.chkY2.Value = 0 And Me.chkN4.Value = 0 Then
        MsgBox "环保属性未输入,不能打印!", vbInformation + vbOKOnly, "未输入环保属性"
        txtSN.SetFocus
        Exit Sub
    End If
    If Me.chkN.Value + Me.chkY.Value + Me.chkY2.Value + Me.chkN4.Value > 1 Then
        MsgBox "环保属性输入多个,不能打印!", vbInformation + vbOKOnly, "输入环保属性多个"
        txtSN.SetFocus
        Exit Sub
    End If
    
    Dim Pb As String
    If chkY2.Value = 1 Then
         Pb = "Y2"
    ElseIf chkY.Value = 1 Then
         Pb = "Y*"
    ElseIf chkN.Value = 1 Then
         Pb = "N*"
    ElseIf chkN4.Value = 1 Then
         Pb = "N4"
    End If
   
   If Trim(txtSN.Text) = "" Then
      MsgBox "产品编码未输入,不能打印!", vbInformation + vbOKOnly, "未输入产品编码"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If txtQty.Text = "" Then
      MsgBox "数量未输入，不能打印！", vbInformation + vbOKOnly, "未输入数量"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If CInt(txtQty.Text) = 0 Then
      MsgBox "请输入正确的数量！", vbInformation + vbOKOnly, "数量不对"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If Trim(txtGW.Text) = "" Then
      MsgBox "产品重量未带出,不能打印!", vbInformation + vbOKOnly, "未带出毛重"
      txtGW.SetFocus
      Exit Sub
   End If
   
   If Trim(txtVer.Text) = "" Then
      MsgBox "版本未输入,不能打印!", vbInformation + vbOKOnly, "未输入版本"
      txtVer.SetFocus
      Exit Sub
   End If
   
   If txtHV.Text = "" Then
      MsgBox "硬件版本未输入,不能打印!", vbInformation + vbOKOnly, "未输入硬件版本"
      txtHV.SetFocus
      Exit Sub
   End If
   

     
'   If Trim(txtPO.Text) <> "/" Then
'        If IsNumeric(txtPO.Text) = False Then
'            MsgBox "PO号必须是数字!", vbInformation + vbOKOnly, "PO号"
'            txtPO.SetFocus
'            Exit Sub
'        End If
'        If Len(Trim(txtPO.Text)) <> 10 Then
'            MsgBox "PO号必须是10位!", vbInformation + vbOKOnly, "PO号"
'            txtPO.SetFocus
'            Exit Sub
'        End If
'   End If

     '======Add by mike 2015.3.24 for data upload to FTPC============
    If UploadH3CInfo(Pb, Trim(txtSN.Text), Trim(txtVer.Text), Trim(txt5000.Text), Trim(txtPC.Text), "CHINA", "NXFHShenChanPrint") = False Then
        MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        Exit Sub
    End If
    
    If UploadH3C_PB(Pb, Trim(txtSN.Text), Trim(txtVer.Text), Trim(txt5000.Text), Trim(txtPC.Text), "CHINA", "NXFHShenChanPrint") = False Then
        MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
   
   
   Dim i, qty As Integer
   Dim leftstr, rightstr, str As String
   qty = CInt(txtQty.Text)
    leftstr = UCase(Left(Trim(txtSN.Text), 14))
    rightstr = Right(Trim(txtSN.Text), 6)
   
'   If Asc(Mid(Right(txtSN.Text, 6), 1, 1)) >= 48 And Asc(Mid(Right(txtSN.Text, 6), 1, 1)) <= 57 Then
'    leftstr = UCase(Left(Trim(txtSN.Text), 14))
'    rightstr = Right(Trim(txtSN.Text), 6)
'   Else
'    leftstr = UCase(Left(Trim(txtSN.Text), 15))
'    rightstr = Right(Trim(txtSN.Text), 5)
'   End If

   
     OpenLppx
     
   'For i = 0 To qty - 1
   
     
   
      'str = leftstr & Right("000000" & CStr(CLng(rightstr) + i), Len(rightstr))
      str = leftstr & rightstr
      
   
 
  myVars.Item("SN").Value = UCase(txtSN.Text)
   myVars.Item("Part Number").Value = Mid(UCase(txtSN.Text), 3, 8)
   myVars.Item("Host Rev").Value = Trim(Me.txtHV.Text)
   If txtSZ.Text <> "SZ" Then
        myObjs("SZ").Top = 10000
   End If
'   If chkVer.Value = 0 Then
'     myVars.Item("Host Rev").Value = ""
'   Else
'       If txtVer.Text = "" Or txtVer.Text = "/" Or txtVer.Text = "N/A" Then
'          myVars.Item("Host Rev").Value = ""
'       Else
'          myVars.Item("Host Rev").Value = Trim(Me.txtVer.Text)
'       End If
'   End If
   
   If Me.chkSVPrint.Value = 1 Then
      If UCase(Trim(Me.txtVer.Text)) = "/" Or UCase(Trim(Me.txtVer.Text)) = "N/A" Then
        myVars.Item("Software").Value = ""
      Else
        myVars.Item("Software").Value = UCase(Trim(Me.txtVer.Text))
      End If
    
   Else
        myVars.Item("Software").Value = ""
   End If
   
''   If Me.chkPCPrint.Value = 1 Then
''        myVars.Item("Code").Value =
   
   myVars.Item("Product Name1").Value = Trim(txtCPN.Text)
   myVars.Item("Product Name2").Value = Trim(txtEPN.Text)
   myVars.Item("Product ID").Value = Trim(txtProductID.Text)

   'If chkOS.Value = 0 Or txtOS.Text = "/" Then
   '   myObjs("OD").Top = 10000
   '   myVars.Item("OD").Value = ""
   'Else
   '   myVars.Item("OD").Value = txtOS.Text
   'End If
   If Trim(txtOS.Text) = "N/A" Then
        myVars.Item("Size").Value = ""
   Else
        myVars.Item("Size").Value = Trim(txtOS.Text)
   End If
   
   If Trim(txtGW.Text) = "/" Or Trim(txtGW.Text) = "N/A" Then
        myVars.Item("Weight").Value = ""
   Else
        myVars.Item("Weight").Value = txtGW.Text
   End If
    
   
   If chkNonCE.Value = 1 Then
      myObjs("CE").Top = 10000
      myObjs("CE address").Top = 10000
      myObjs("HPE addr").Top = 10000
   Else
      If chkCEAddr.Value = 0 Then
        myObjs("CE address").Top = 10000
      End If
      
      If chkHPEAddr.Value = 0 Then
        myObjs("HPE addr").Top = 10000
      End If
   End If
   If chkNonWEEE.Value = 1 Then
      myObjs("WEEE").Top = 10000
   End If
   If chkNonEAC.Value = 1 Then
      myObjs("EAC").Top = 10000
   End If
   If chkNonChinaRoHS.Value = 1 Then
      myObjs("China RoHS").Top = 10000
   End If
   If chkNonRoHS.Value = 1 Then
      myObjs("H3C RoHS").Top = 10000
   End If
   If Me.chkNonTurkey.Value = 1 Then
      myObjs("Turkey RoHS").Top = 10000
   End If
   
   If Me.chkNonUkraine.Value = 1 Then
      myObjs("Ukraine").Top = 10000
      myVars.Item("Ukraine ID").Value = ""
   Else
      myVars.Item("Ukraine ID").Value = Me.txtUkraineID.Text
   End If
   
   If Me.chkATick.Value = 1 Then
      myVars.Item("A-Tick ID").Value = Trim(Me.txtATick.Text)
   Else
      myObjs("A-Tick").Top = 10000
      myVars.Item("A-Tick ID").Value = ""
   End If
   
   If Me.chkCTick.Value = 1 Then
      myVars.Item("C-Tick ID").Value = Trim(Me.txtCTick.Text)
   Else
      myObjs("C-Tick").Top = 10000
      myVars.Item("C-Tick ID").Value = ""
   End If
   
   If Me.chkICT.Value = 1 Then
      If Trim(Me.txtICT.Text) <> "/" And Trim(Me.txtICT.Text) <> "N/A" Then
        myVars.Item("ICT ID").Value = "HP" + " " + Trim(Me.txtICT.Text)
      Else
        myVars.Item("ICT ID").Value = "HP"
      End If
   Else
      myObjs("ICT").Top = 10000
      myVars.Item("ICT ID").Value = ""
   End If
   
   If Me.chkRCM.Value = 1 Then
      myVars.Item("RCM ID").Value = Trim(Me.txtRCM.Text)
   Else
      myObjs("RCM").Top = 10000
      myVars.Item("RCM ID").Value = ""
   End If
   
   If Me.chkGost.Value = 1 Then
      myVars.Item("Gost ID").Value = Trim(Me.txtGost.Text)
   Else
      myObjs("Gost").Top = 10000
      myVars.Item("Gost ID").Value = ""
   End If
   
   If Me.chkKC.Value = 1 Then
      myVars.Item("KC ID").Value = Trim(Me.txtKC.Text)
   Else
      myObjs("KC").Top = 10000
      myVars.Item("KC ID").Value = ""
   End If
   
   'If optH3CRoHS.Value = True Then
   '   myObjs("3COM RoHS").Top = 10000
   'End If
   'If opt3COMRoHS.Value = True Then
   '   myObjs("3COM RoHS").Top = 2300
   '   myObjs("H3C RoHS").Top = 10000
   'End If
   'If optNonRoHS.Value = True Then
   '   myObjs("H3C RoHS").Top = 10000
   '   myObjs("3COM RoHS").Top = 10000
   'End If
   If Trim(txtMS.Text) = "/" Or Trim(txtMS.Text) = "N/A" Then
        myObjs.Item("MS Title").Top = 10000
        myVars.Item("MS").Value = ""
   Else
        myVars.Item("MS").Value = txtMS.Text
   End If
   
   If Trim(Me.txtNAL1.Text) <> "/" And Trim(Me.txtNAL1.Text) <> "N/A" And Trim(Me.txtNAL1Title.Text) <> "/" And Trim(Me.txtNAL1Title.Text) <> "N/A" Then
           myVars.Item("NAL").Value = txtNAL1Title.Text & "(" & Trim(Me.txtNAL1.Text) & ")" '
   Else
        myVars.Item("NAL").Value = ""
        myObjs.Item("NAL Title1").Top = 10000
   End If
   
   If Trim(Me.txtNAL2.Text) <> "/" And Trim(Me.txtNAL2.Text) <> "N/A" And Trim(Me.txtNAL2Title.Text) <> "/" And Trim(Me.txtNAL2Title.Text) <> "N/A" Then
        myVars.Item("NAL2").Value = txtNAL2Title.Text & "(" & Me.txtNAL2.Text & ")"
   Else
        myVars.Item("NAL2").Value = ""
        myObjs.Item("NAL Title2").Top = 10000
   End If
   
   If txtHV.Text = "" Or txtHV.Text = "/" Or txtHV.Text = "N/A" Then
      myObjs("BHver").Top = 10000
      myVars.Item("HVer").Value = "N/A"
   Else
''      myObjs("THver").Top = 10000
''      myVars.Item("HVer").Value = UCase(Trim(Replace(txtHV.Text, vbCrLf, "")))
   End If
   
   If Trim(Me.txtPC.Text) <> "" And Trim(Me.txtPC.Text) <> "/" And UCase(Trim(Me.txtPC.Text)) <> "N/A" Then
        myVars.Item("Plant Code").Value = UCase(Trim(Me.txtPC.Text))
   Else
        myVars.Item("Plant Code").Value = ""
   End If

   myVars.Item("Y2").Value = Pb
   
   
'   If Trim(txtPO.Text) = "/" Then
'        myObjs("Text14(22)").Top = 100000
'        myVars.Item("PO").Value = ""
'   Else
'        myVars.Item("PO").Value = Trim(txtPO.Text)
'   End If
   
   
   mydoc.PrintLabel 1
   mydoc.FormFeed
   'Next
   
   UnloadLppx

   cmdCancel_Click
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
    Me.txtVer.Text = Clipboard.GetText()
End Sub

Private Sub Form_Load()
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub




Private Sub txtHV_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 13) Then
     txtMS.SetFocus
  End If
End Sub

Private Sub txtMS_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtNAL.SetFocus
  End If
End Sub



Private Sub txtNAL_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtRemark.SetFocus
  End If
End Sub

Private Sub txtQty_Change()
If txtQty.Text <> "" Then
    If Asc(Right(txtQty.Text, 1)) > 57 Or Asc(Right(txtQty.Text, 1)) < 48 Then
       MsgBox "只能输入数字！", vbInformation + vbOKOnly, "输入不正确"
       SendKeys "{backspace}"
       txtQty.SetFocus
       Exit Sub
    End If
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtVer.SetFocus
  End If
End Sub



Private Sub txtRemark_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     cmdPrint_Click
  End If
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      If Len(Trim(txtSN.Text)) < 10 Then
         MsgBox "产品序号长度不能小于10!"
         txtSN.SetFocus
         Exit Sub
      End If

      sql = "select top 1 * from revset where model='" & Mid(Trim(txtSN.Text), 3, 8) & "' and firstall<='" & Trim(txtSN.Text) & "' and endall>='" & Trim(txtSN.Text) & "'order by ver desc"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "此产品序号未收集版本!"
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
         txtHV.Text = Trim(rec.Fields(3))
      End If
      rec.Close
      
      
             sql = "SELECT [ID],[Part_Number],[Part_Revision],[ProductID],[CPN],[EPN],[Des],[Size],[GW],[MS],case when n.[NAL1] is null then '/' when  n.[NAL1] is not null then n.[NAL1] end as 'n.[NAL1]',case when n.[NAL1_Title] is null then '/' when  n.[NAL1_Title] is not null then n.[NAL1_Title] end as 'n.[NAL1_Title]',[NAL2],[NAL2_Title]" & _
        ",case when [CE] = 0 then 'Non CE' when CE = 1 then 'CE' when CE = 2 then 'CE+CE Addr'when CE = 3 then 'CE+HPE Addr' end as 'CE'" & _
        ",case when WEEE is null then 'N/A' when WEEE = 0 then 'No' when WEEE = 1 then 'Yes' end as 'WEEE'" & _
        ",case when ChinaRoHS is null then 'N/A' when ChinaRoHS = 0 then 'No' when ChinaRoHS = 1 then 'Yes' end as 'ChinaRoHS'" & _
        ",case when [RoHS] is null then 'N/A' when RoHS = 0 then 'No' when RoHS = 1 then 'Yes' end as 'RoHS'" & _
        ",case when [TurkeyRoHS] is null then 'N/A' when [TurkeyRoHS] = 0 then 'No' when TurkeyRoHS = 1 then 'Yes' end as '[TurkeyRoHS]'" & _
        ",case when Ukraine is null then 'N/A' when Ukraine = 0 then 'No' when Ukraine = 1 then 'Yes' end as 'Ukraine'" & _
        ",[Ukraine_ID]" & _
        ",case when ATick is null then 'N/A' when ATick = 0 then 'No' when ATick = 1 then 'Yes' end as 'ATick'" & _
        ",[ATick_ID]" & _
        ",case when CTick is null then 'N/A' when CTick = 0 then 'No' when CTick = 1 then 'Yes' end as 'CTick'" & _
        ",[CTick_ID]" & _
        ",case when ICT is null then 'N/A' when ICT = 0 then 'No' when ICT = 1 then 'Yes' end as 'ICT'" & _
        ",[ICT_ID]" & _
        ",case when RCM is null then 'N/A' when RCM = 0 then 'No' when RCM = 1 then 'Yes' end as 'RCM'" & _
        ",[RCM_ID]" & _
        ",case when Gost is null then 'N/A' when Gost = 0 then 'No' when Gost = 1 then 'Yes' end as 'Gost'" & _
        ",[Gost_ID]" & _
        ",case when KC is null then 'N/A' when KC = 0 then 'No' when KC = 1 then 'Yes' end as 'KC'" & _
        ",[KC_ID]" & _
        ",case when Print_SV is null then 'N/A' when Print_SV = 0 then 'No' when Print_SV = 1 then 'Yes' end as 'Print_SV'" & _
        ",case when Print_Power is null then 'N/A' when Print_Power = 0 then 'No' when Print_Power = 1 then 'Yes' end as 'Print_Power'" & _
        ",[5000_State],[Remark],case when EAC is null then 'N/A' when EAC = 0 then 'No' when EAC = 1 then 'Yes' end as 'EAC' FROM [Print].[dbo].[tblUnisNew] " & _
        "LEFT JOIN [Print].[dbo].[tblNetMaintain] n ON Part_Number = n.partNumber and n.valid =1 where Part_Number = '" & _
        Mid(txtSN.Text, 3, 8) & "' and Part_Revision ='" & txtHV.Text & "'"
      
      
      'sql = "select ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark from tblUnisNew where SN='" & Mid(txtSN.Text, 3, 8) & "' and HV='" & txtHV.Text & "'"
'       sql = "SELECT [ID],[Part_Number],[Part_Revision],[ProductID],[CPN],[EPN],[Des],[Size],[GW],[MS],[NAL1],[NAL1_Title],[NAL2],[NAL2_Title]" & _
'        ",[CE], WEEE,ChinaRoHS,[RoHS],[TurkeyRoHS],Ukraine,ATick,[ATick_ID],CTick,[CTick_ID],ICT,[ICT_ID],RCM,[RCM_ID],Gost,[Gost_ID],KC,[KC_ID],Print_SV,Print_Power,[5000_State],[Remark] FROM [Print].[dbo].[tblUnisNew] where Part_Number = '" & _
'         Mid(txtSN.Text, 3, 8) & "' and Part_Revision ='" & txtHV.Text & "'"
      
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      If rec.EOF = True Then
         MsgBox "此产品编码未进行设置!"
         txtVer.Text = ""
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
        Me.txtProductID.Text = rec.Fields(3)
        txtCPN.Text = rec.Fields(4)
        txtEPN.Text = rec.Fields(5)
''        txtDes.Text = rec.Fields(5)
        
'        Dim psv As String
'        psv = rec.Fields(17)
'        If UCase(psv) = "N" Then
'            chkVer.Value = 0
'        Else
'            chkVer.Value = 1
'        End If
        
        chkOS.Value = 1
        txtOS.Text = rec.Fields(7)
        '============edit by ben end  =============
        If IsNull(rec.Fields(8)) = True Then
            txtGW.Text = ""
        Else
            txtGW.Text = rec.Fields(8)
        End If
        txtMS.Text = rec.Fields(9)
        Me.txtNAL1.Text = rec.Fields(10)
        Me.txtNAL1Title.Text = rec.Fields(11)
        Me.txtNAL2.Text = rec.Fields(12)
        Me.txtNAL2Title.Text = rec.Fields(13)
        If UCase(Trim(rec.Fields(14))) = "CE" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
           chkCEAddr.Value = 0
           chkHPEAddr.Value = 0
        ElseIf rec.Fields(14) = "Non CE" Then
           chkCE.Value = 0
           chkNonCE.Value = 1
           chkCEAddr.Value = 0
           chkHPEAddr.Value = 0
        ElseIf rec.Fields(14) = "CE+CE Addr" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
           chkCEAddr.Value = 1
           chkHPEAddr.Value = 0
        ElseIf rec.Fields(14) = "CE+HPE Addr" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
           chkCEAddr.Value = 0
           chkHPEAddr.Value = 1
        End If
        If UCase(Trim(rec.Fields(15))) = "YES" Then
           chkWEEE.Value = 1
           chkNonWEEE.Value = 0
        ElseIf rec.Fields(15) = "No" Or rec.Fields(15) = "N/A" Then
           chkWEEE.Value = 0
           chkNonWEEE.Value = 1
        End If
        If UCase(Trim(rec.Fields(16))) = "YES" Then
           chkChinaRoHS.Value = 1
           chkNonChinaRoHS.Value = 0
        ElseIf rec.Fields(16) = "No" Or rec.Fields(16) = "N/A" Then
           chkChinaRoHS.Value = 0
           chkNonChinaRoHS.Value = 1
        End If
        If Me.chkEP.Value = 0 Then
            If UCase(Trim(rec.Fields(17))) = "YES" Then
                chkRoHS.Value = 1
                chkNonRoHS.Value = 0
            ElseIf rec.Fields(17) = "No" Or rec.Fields(17) = "N/A" Then
                chkRoHS.Value = 0
                chkNonRoHS.Value = 1
            End If
            If UCase(Trim(rec.Fields(18))) = "YES" Then
                chkTurkey.Value = 1
                chkNonTurkey.Value = 0
            ElseIf rec.Fields(18) = "No" Or rec.Fields(18) = "N/A" Then
                chkTurkey.Value = 0
                chkNonTurkey.Value = 1
            End If
            If UCase(Trim(rec.Fields(19))) = "YES" Then
                Me.chkUkraine.Value = 1
                Me.chkNonUkraine.Value = 0
                Me.txtUkraineID.Text = rec.Fields(20)
            ElseIf rec.Fields(19) = "No" Or rec.Fields(19) = "NO" Then
                Me.chkUkraine.Value = 0
                Me.chkNonUkraine.Value = 1
                Me.txtUkraineID.Text = ""
            End If
        Else
            chkNonWEEE.Value = 1
            chkNonTurkey.Value = 1
            chkNonRoHS.Value = 1
            chkChinaRoHS.Value = 1
            chkNonCE.Value = 1
            chkCEAddr.Value = 0
            chkHPEAddr.Value = 0
        End If
        If UCase(Trim(rec.Fields(21))) = "YES" Then
            Me.chkATick.Value = 1
            Me.chkNonATick.Value = 0
        Else
            Me.chkATick.Value = 0
            Me.chkNonATick.Value = 1
            Me.txtATick.Text = ""
        End If
        Me.txtATick.Text = Trim(rec.Fields(22))
        
        If UCase(Trim(rec.Fields(23))) = "YES" Then
            Me.chkCTick.Value = 1
            Me.chkNonCTick.Value = 0
        Else
            Me.chkCTick.Value = 0
            Me.chkNonCTick.Value = 1
            Me.txtCTick.Text = ""
        End If
        Me.txtCTick.Text = rec.Fields(24)
        
        If UCase(Trim(rec.Fields(25))) = "YES" Then
            Me.chkICT.Value = 1
            Me.chkNonICT.Value = 0
        Else
            Me.chkICT.Value = 0
            Me.chkNonICT.Value = 1
            Me.txtICT.Text = ""
        End If
        Me.txtICT.Text = rec.Fields(26)
        
        If UCase(Trim(rec.Fields(27))) = "YES" Then
            Me.chkRCM.Value = 1
            Me.chkNonRCM.Value = 0
        Else
            Me.chkRCM.Value = 0
            Me.chkNonRCM.Value = 1
            Me.txtRCM.Text = ""
        End If
        Me.txtRCM.Text = rec.Fields(28)
        
        If UCase(Trim(rec.Fields(29))) = "YES" Then
            Me.chkGost.Value = 1
            Me.chkNonGost.Value = 0
        Else
            Me.chkGost.Value = 0
            Me.chkNonGost.Value = 1
            Me.txtGost.Text = ""
        End If
        Me.txtGost.Text = rec.Fields(30)
        
         If UCase(Trim(rec.Fields(31))) = "YES" Then
            Me.chkKC.Value = 1
            Me.chkNonKC.Value = 0
        Else
            Me.chkKC.Value = 0
            Me.chkNonKC.Value = 1
            Me.txtKC.Text = ""
        End If
        Me.txtKC.Text = rec.Fields(32)
        
        If UCase(Trim(rec.Fields(33))) = "YES" Then
            Me.chkSVPrint.Value = 1
            Me.chkNonSVPrint.Value = 0
        Else
            Me.chkSVPrint.Value = 0
            Me.chkNonSVPrint.Value = 1
        End If
        
'        If UCase(Trim(rec.Fields(34))) = "YES" Then
'            Me.chkPCPrint.Value = 1
'            Me.chkNonPCPrint.Value = 0
'        Else
'            Me.chkPCPrint.Value = 0
'            Me.chkNonPCPrint.Value = 1
'        End If
        
        Me.txt5000.Text = rec.Fields(35)
        
        txtHV.Text = rec.Fields(2)
        Me.txtPC.Text = "N/A"
        txtRemark.Text = rec.Fields(36)
        
        If UCase(Trim(rec.Fields("EAC"))) = "YES" Then
           chkEAC.Value = 1
           chkNonEAC.Value = 0
        ElseIf rec.Fields("EAC") = "No" Or rec.Fields("EAC") = "N/A" Then
           chkEAC.Value = 0
           chkNonEAC.Value = 1
        End If
      End If
        
'      End If
      rec.Close
      
       '============add by carson start for TR5=============
'        Dim conSZ As ADODB.Connection
'        Dim rsSZ As ADODB.Recordset
'        Set conSZ = New ADODB.Connection
'        Set rsSZ = New ADODB.Recordset
'        conSZ.ConnectionString = "Provider=SQLOLEDB;User ID=sa;PWD=Flash123;Initial Catalog=afg_active_90;Data Source=10.11.1.130"
'        conSZ.ConnectionTimeout = 50
'        conSZ.Open
''        Dim stringSQL As String
'        Set rsSZ.ActiveConnection = conSZ
'        rsSZ.CursorType = adOpenDynamic
'
'        stringSQL = " select TOP 1 'SZ' from C_NoTR5_Part where EFFE_FLAG='1' AND  Part_Number ='" & Mid(txtSN.Text, 3, 8) & "'  "
'
'        rsSZ.Open stringSQL
'        If rsSZ.EOF = True Then
'            txtSZ.Text = ""
'        Else
'            txtSZ.Text = rsSZ.Fields(0)
'        End If
'        rsSZ.Close
      '============add by carson end  =============
      txtQty.SetFocus
   End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
'============add by carson start for OldBox for nixiang 20170627=============
        Dim OldFlag As Boolean
        
        Dim conSZ As ADODB.Connection
        Dim rsSZ As ADODB.Recordset
        Set conSZ = New ADODB.Connection
        Set rsSZ = New ADODB.Recordset
        conSZ.ConnectionString = "Provider=SQLOLEDB;User ID=sa;PWD=Flash123;Initial Catalog=afg_active_90;Data Source=10.11.1.130"
        conSZ.ConnectionTimeout = 50
        conSZ.Open
'        Dim stringSQL As String
        Set rsSZ.ActiveConnection = conSZ
        rsSZ.CursorType = adOpenDynamic

        stringSQL = " select TOP 1 1 from C_OldBox_Part where EFFE_FLAG='1' AND  Part_Number ='" & Mid(txtSN.Text, 3, 8) & "'  "

        rsSZ.Open stringSQL
        If rsSZ.EOF = True Then
            OldFlag = False
        Else
            OldFlag = True
        End If
        rsSZ.Close
  '============add by carson end  =============
   Me.MousePointer = vbHourglass
'   \\sz-fs01\Public\Manufacture\标签模板\H3C发货标签NEW\逆向H3CNEW
   If OldFlag = True Then
        Set mydoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\H3C发货标签NEW\逆向H3CNEW\" & "UNIS-生产.lab")
   Else
        Set mydoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\H3C发货标签NEW\逆向H3CNEW\新H3C地址模板\" & "UNIS-生产.lab")
   End If
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = mydoc.Variables
   Set myObjs = mydoc.DocObjects
End Sub

Private Sub txtVer_GotFocus()
    Me.txtVer.Text = Clipboard.GetText()
    Clipboard.Clear
End Sub

Private Sub txtVer_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtSN.SetFocus
  End If
End Sub

Private Sub txtVer_KeyUp(KeyCode As Integer, Shift As Integer)
'      If KeyAscii <> 13 Then
''        If (Len(Me.txtVer.Text) > 10 Or Me.txtVer = "/") Then
''        Else
''            MsgBox "只可以粘贴，不能手动输入"
''            Me.txtVer.Text = ""
''            Exit Sub
''        End If
'    End If
End Sub
