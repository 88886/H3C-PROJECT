VERSION 5.00
Begin VB.Form frmNewH3CPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New H3C Label Print"
   ClientHeight    =   12195
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   16605
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewH3CPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12195
   ScaleWidth      =   16605
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox lblNALday 
      Height          =   450
      Left            =   240
      TabIndex        =   44
      Top             =   11640
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox lblMSday 
      Height          =   450
      Left            =   240
      TabIndex        =   43
      Top             =   11040
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   13920
      TabIndex        =   34
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   13920
      TabIndex        =   33
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   13920
      TabIndex        =   32
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   120
      TabIndex        =   19
      Top             =   5040
      Width           =   16335
      Begin VB.CheckBox chkHPE1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HPE1 Addr"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   105
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CheckBox chkN4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N4"
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
         TabIndex        =   104
         Top             =   6240
         Width           =   855
      End
      Begin VB.CheckBox chkY2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y2"
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
         TabIndex        =   103
         Top             =   6240
         Width           =   855
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y*"
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
         TabIndex        =   102
         Top             =   6240
         Width           =   855
      End
      Begin VB.CheckBox chkN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N*"
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
         Left            =   4680
         TabIndex        =   101
         Top             =   6240
         Width           =   855
      End
      Begin VB.TextBox txtSZ 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   450
         Left            =   7800
         TabIndex        =   99
         Top             =   5640
         Width           =   2175
      End
      Begin VB.PictureBox Picture13 
         Height          =   495
         Left            =   5520
         Picture         =   "frmNewH3CPrint.frx":13652
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   98
         Top             =   5040
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
         Left            =   7320
         TabIndex        =   97
         Top             =   5160
         Width           =   735
      End
      Begin VB.CheckBox chkEAC 
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   96
         Top             =   5160
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
         TabIndex        =   95
         Top             =   2760
         Width           =   1575
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
         TabIndex        =   92
         Top             =   5760
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
         Left            =   3840
         TabIndex        =   91
         Top             =   5760
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
         TabIndex        =   89
         Top             =   720
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
         TabIndex        =   88
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txt5000 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   7800
         TabIndex        =   87
         Top             =   4440
         Width           =   2175
      End
      Begin VB.CheckBox chkPCPrint 
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
         Left            =   11880
         TabIndex        =   84
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox chkNonPCPrint 
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
         Left            =   12600
         TabIndex        =   83
         Top             =   720
         Width           =   615
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
         TabIndex        =   82
         Top             =   3795
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
         TabIndex        =   79
         Top             =   3240
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
         TabIndex        =   78
         Top             =   2715
         Width           =   2175
      End
      Begin VB.TextBox txtRCM 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   74
         Top             =   3120
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
         TabIndex        =   73
         Top             =   3120
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
         TabIndex        =   72
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtGost 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   69
         Top             =   3720
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
         TabIndex        =   68
         Top             =   3720
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
         TabIndex        =   67
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txtKC 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   64
         Top             =   4320
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
         TabIndex        =   63
         Top             =   4320
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
         TabIndex        =   62
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox txtICT 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   59
         Top             =   2450
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
         TabIndex        =   58
         Top             =   2520
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
         TabIndex        =   57
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtCTick 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   54
         Top             =   1920
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
         TabIndex        =   53
         Top             =   1920
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
         TabIndex        =   52
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtATick 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   14520
         TabIndex        =   49
         Top             =   1440
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
         TabIndex        =   48
         Top             =   1440
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
         TabIndex        =   47
         Top             =   1440
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
         TabIndex        =   45
         Top             =   2760
         Width           =   1215
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
         TabIndex        =   42
         Top             =   5160
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
         TabIndex        =   41
         Top             =   5280
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
         Left            =   5640
         TabIndex        =   40
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
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
         TabIndex        =   39
         Top             =   4320
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
         TabIndex        =   38
         Top             =   4320
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
         TabIndex        =   35
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
         TabIndex        =   7
         Top             =   2280
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
         TabIndex        =   8
         Top             =   2280
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
         TabIndex        =   9
         Top             =   4800
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
         TabIndex        =   10
         Top             =   4800
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
         TabIndex        =   11
         Top             =   3840
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
         TabIndex        =   12
         Top             =   3840
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
         Left            =   14040
         TabIndex        =   16
         Top             =   240
         Width           =   2175
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
         TabIndex        =   14
         Top             =   2160
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
         TabIndex        =   13
         Top             =   1650
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
         Left            =   11880
         TabIndex        =   15
         Top             =   240
         Width           =   855
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
         TabIndex        =   6
         Top             =   1200
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   720
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
         TabIndex        =   2
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
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtVer 
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
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label9 
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
         Left            =   6480
         TabIndex        =   100
         Top             =   5760
         Width           =   495
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "环保属性:"
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
         Left            =   720
         TabIndex        =   94
         Top             =   6240
         Width           =   1215
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
         TabIndex        =   93
         Top             =   5760
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
         Top             =   835
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
         TabIndex        =   86
         Top             =   4485
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "电源代码打印"
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
         Left            =   10080
         TabIndex        =   85
         Top             =   795
         Width           =   1695
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
         TabIndex        =   81
         Top             =   3840
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
         TabIndex        =   80
         Top             =   3270
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
         TabIndex        =   77
         Top             =   2760
         Width           =   1335
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
         TabIndex        =   75
         Top             =   3120
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
         TabIndex        =   71
         Top             =   3120
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
         TabIndex        =   70
         Top             =   3720
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
         TabIndex        =   66
         Top             =   3720
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
         TabIndex        =   65
         Top             =   4320
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
         TabIndex        =   61
         Top             =   4320
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
         TabIndex        =   60
         Top             =   2520
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
         TabIndex        =   56
         Top             =   2520
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
         TabIndex        =   55
         Top             =   1920
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
         TabIndex        =   51
         Top             =   1920
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
         TabIndex        =   50
         Top             =   1440
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
         TabIndex        =   46
         Top             =   1440
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
         Left            =   360
         TabIndex        =   37
         Top             =   4320
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
         Height          =   375
         Left            =   960
         TabIndex        =   36
         Top             =   5280
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
         Left            =   12960
         TabIndex        =   31
         Top             =   240
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
         TabIndex        =   30
         Top             =   2280
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
         TabIndex        =   29
         Top             =   1680
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
         Left            =   360
         TabIndex        =   28
         Top             =   3840
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
         Left            =   10080
         TabIndex        =   27
         Top             =   270
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
         TabIndex        =   26
         Top             =   4800
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
         TabIndex        =   25
         Top             =   2520
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
         TabIndex        =   24
         Top             =   1200
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
         TabIndex        =   23
         Top             =   720
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      Picture         =   "frmNewH3CPrint.frx":156D5
      ScaleHeight     =   4545
      ScaleWidth      =   12105
      TabIndex        =   18
      Top             =   360
      Width           =   12135
   End
   Begin VB.Label Label10 
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
      Left            =   5880
      TabIndex        =   76
      Top             =   9480
      Width           =   1335
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
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmNewH3CPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private PRI_pb As String

Dim rec1 As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim hpsn, SN21 As String
Dim checkhp As New ADODB.Recordset
Public HP_pack_label As Boolean
Dim str As String

Function GetPb() As String
    GetPb = PRI_pb
End Function
Sub SetPb(ByVal PB As String)
    PRI_pb = PB
End Sub
Sub InitPb(ByVal sn As String)
    If (sn = "") Then
        chkY2.Value = 0
        chkY.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
    Else
        
    End If
End Sub


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
End Sub

Private Sub chkCE_Click()
   If chkCE.Value = 1 Then
      chkNonCE.Value = 0
   Else
      chkNonCE.Value = 1
   End If
End Sub

Private Sub chkNonCE_Click()
   If chkNonCE.Value = 1 Then
      chkCE.Value = 0
   Else
      chkCE.Value = 1
   End If
End Sub

Private Sub chkChinaRoHS_Click()
   If chkChinaRoHS.Value = 1 Then
      chkNonChinaRoHS.Value = 0
   Else
      chkNonChinaRoHS.Value = 1
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
    Else
        chkRoHS.Value = 1
    End If
End Sub

Private Sub chkRoHS_Click()
    If chkRoHS.Value = 1 Then
        chkNonRoHS.Value = 0
    Else
        chkNonRoHS.Value = 1
    End If
End Sub

Private Sub chkTurkey_Click()
    If chkTurkey.Value = 1 Then
        chkNonTurkey.Value = 0
    Else
        chkNonTurkey.Value = 1
    End If
End Sub

Private Sub chkNonTurkey_Click()
     If chkNonTurkey.Value = 1 Then
        chkTurkey.Value = 0
     Else
        chkTurkey.Value = 1
     End If
End Sub

Private Sub chkWEEE_Click()
   If chkWEEE.Value = 1 Then
      chkNonWEEE.Value = 0
   Else
      chkNonWEEE.Value = 1
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

Private Sub chkY_Click()
    If chkY.Value = 1 Then
        chkY2.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub chkY2_Click()
    If chkY2.Value = 1 Then
        chkY.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub chkN_Click()
    If chkN.Value = 1 Then
        chkY2.Value = 0
        chkY.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub chkN4_Click()
    If chkN4.Value = 1 Then
        chkY.Value = 0
        chkN.Value = 0
        chkY2.Value = 0
    End If
End Sub

Private Sub cmdCancel_Click()
   Reset
   txtSN.SetFocus
End Sub

Private Sub cmdPrint_Click()
   Dim pc As String
   If Me.chkPCPrint.Value = 1 Then
        pc = getPowerCode(Trim(Me.txtSN.Text))
        If pc = "" Then
            Exit Sub
        End If
   End If
   
   If Trim(txtSN.Text) = "" Then
      MsgBox "产品编码未输入,不能打印!", vbInformation + vbOKOnly, "未输入产品编码"
      txtSN.SetFocus
      Exit Sub
   End If
   If Trim(txtVer.Text) = "" Then
      MsgBox "软件版本未带出,不能打印,请重新输入产品编码!", vbInformation + vbOKOnly, "未带出版本"
      txtSN.SetFocus
      Exit Sub
   End If
   If Trim(txtHV.Text) = "" Then
      MsgBox "产品没有硬件版本,不能打印!", vbInformation + vbOKOnly, "没有硬件版本"
      txtHV.SetFocus
      Exit Sub
   End If
   If Trim(txtGW.Text) = "" Then
      MsgBox "产品重量未带出,不能打印!", vbInformation + vbOKOnly, "未带出毛重"
      txtGW.SetFocus
      Exit Sub
   End If
   If Trim(Me.txt5000.Text) = "" Then
      MsgBox "5000米状态未带出,不能打印!", vbInformation + vbOKOnly, "未带出5000米状态"
      Exit Sub
   End If
   
   '===============add by ben start===============
'   If Trim(txtOS.Text) = "" Then
'      MsgBox "外尺寸未带出，不能打印!", vbInformation + vbOKOnly, "未带出外尺寸"
'      txtSN.SetFocus
'      Exit Sub
'   End If
   '===============add by ben end  ===============
   
'===============add by ben 2012-02-05 start===============
    If reprint = False Then
        If Connect.isPrintedLabel(Me.txtSN.Text, Me.Name) Then
            MsgBox ("此序列号已打印！")
            txtSN.SetFocus
            Exit Sub
        End If
    End If
'===============add by ben 2012-02-05 end=================


Dim PB As String
If (chkY2.Value = 1) Then
    PB = "Y2"
ElseIf (chkY.Value = 1) Then
    PB = "Y*"
ElseIf (chkN.Value = 1) Then
    PB = "N*"
ElseIf (chkN4.Value = 1) Then
    PB = "N4"
End If

   If chkPCPrint.Value = 1 Then
        If UploadH3CInfo2(True, Trim(Me.txtSN.Text), Trim(Me.txtVer.Text), Trim(Me.txt5000.Text), pc, "CHINA", golUSERNAME, PB) = False Then
             MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
             txtSN.SetFocus
             Exit Sub
        End If
   Else
       If UploadH3CInfo2(False, Trim(Me.txtSN.Text), Trim(Me.txtVer.Text), Trim(Me.txt5000.Text), pc, "CHINA", golUSERNAME, PB) = False Then
             MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
             txtSN.SetFocus
             Exit Sub
        End If
   End If
   

If UploadH3C_PB(PB, Trim(UCase(txtSN.Text)), Trim(UCase(txtVer.Text)), "NA", "N/A", "CHINA", "frmNewH3CPrint") = False Then
    MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
    txtSN.SetFocus
    Exit Sub
End If
   
   
   OpenLppx
   
   myVars.Item("Y2").Value = PB
   
   myVars.Item("SN").Value = UCase(txtSN.Text)
   myVars.Item("Part Number").Value = Mid(UCase(txtSN.Text), 3, 8)
   myVars.Item("Host Rev").Value = Trim(Me.txtHV.Text)
   
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
        myVars.Item("Software").Value = Trim(Me.txtVer.Text)
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
      myObjs("HPE1 addr").Top = 10000
   Else
      If chkCEAddr.Value = 0 Then
        myObjs("CE address").Top = 10000
      End If
      
      If chkHPEAddr.Value = 0 Then
        myObjs("HPE addr").Top = 10000
      End If
      
      If chkHPE1.Value = 0 Then
        myObjs("HPE1 addr").Top = 10000
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
   
   If Me.chkPCPrint.Value = 1 Then
        myVars.Item("Plant Code").Value = pc
   Else
        myVars.Item("Plant Code").Value = ""
   End If
   
   If txtSZ.Text <> "SZ" Then
        myObjs("SZ").Top = 10000
   End If
   
   
   
   
'   myVars.Item("Remark").Value = UCase(txtRemark.Text)
   
   sql = "Insert Into tblHPonline_PrintLog(SN,PTime,Printer) values ('" & UCase(txtSN.Text) & "',getdate(),'" & golUSERNAME & "')"
   conn.Execute sql
   
   Dim smodel As String
   smodel = Mid(Trim(txtSN.Text), 3, 8)
   
   'myApp.Visible = True
   myDoc.PrintLabel 1
   myDoc.FormFeed
   
'===============add by ben 2012-02-05 start===============
                Call Connect.addPrintedLabel(Me.txtSN.Text, Me.Name)
'===============add by ben 2012-02-05 end=================
   
   UnloadLppx
   
   SN21 = txtSN.Text
   cmdCancel_Click
   If HP_pack_label = True Then
    frmH3CPrint.Hide
    'add hp print
    
    FormHPFahuo.txtSN = hpsn
    FormHPFahuo.txtModel_hid = smodel
    FormHPFahuo.txtHPSN = UCase(Trim(SN21))
    
    
    FormHPFahuo.Show
    Call FormHPFahuo.cmdMPrint_Click
   End If

End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
Dim u As New Unit
If u.Init("210231A2Y5B154000010") Then
    MsgBox ("sn = " & u.SerialNumber & ", wo = " & u.WorkOrder & ", pb = " & u.PB)
End If
Set u = Nothing

End Sub

Private Sub Form_Load()
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
'   If conn1.State = 0 Then
'      conn1.ConnectionString = "Provider=SQLOLEDB;User ID=datasweep;PWD=datasweep;Initial Catalog=dsActive;Data Source=DS-DB"
'      conn1.Open
'   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   If conn1.State = 1 Then
      conn1.Close
      Set conn1 = Nothing
   End If
End Sub

Public Function get_nextchar(strRemark As String, pipei As String) As String

    If InStr(strRemark, pipei) > 0 Then
        get_nextchar = UCase(Mid(strRemark, InStr(strRemark, pipei) + Len(pipei), 1))
    End If

End Function

Public Function get_ver(strVer As String) As String

    If InStr(strVer, "-") > 1 Then
        get_ver = Mid(strVer, 1, InStr(strVer, "-") - 1)
    Else
        get_ver = strVer
    End If
    

End Function



Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      If Len(txtSN.Text) < 10 Then
         MsgBox "产品序号长度不能小于10!"
         txtSN.SetFocus
         Exit Sub
      End If
      
      
'        Dim u As New Unit
'        If (u.Init(txtSN.Text)) Then
'            If (u.Pb = "NPb") Then
'              chkY2.Value = 1
'              chkY.Value = 0
'            Else
'              chkY2.Value = 0
'              chkY.Value = 1
'            End If
'        Else
'            MsgBox ("无效条码")
'            txtSN.Text = ""
'            txtSN.SetFocus
'            Exit Sub
'        End If
    Dim lh As New Label_History
    Dim sn As String
    sn = txtSN.Text
    If (lh.Init(sn)) Then
        If lh.PB = "Y*" Then
            chkY.Value = 1
            chkY2.Value = 0
            chkN4.Value = 0
            chkN.Value = 0
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
        ElseIf lh.PB = "Y2" Then
            chkY.Value = 0
            chkY2.Value = 1
            chkN4.Value = 0
            chkN.Value = 0
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
        ElseIf lh.PB = "N*" Then
            chkY.Value = 0
            chkY2.Value = 0
            chkN4.Value = 0
            chkN.Value = 1
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
        ElseIf lh.PB = "N4" Then
            chkY.Value = 0
            chkY2.Value = 0
            chkN4.Value = 1
            chkN.Value = 0
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
        End If
    Else
        chkY.Enabled = True
        chkY2.Enabled = True
        chkN.Enabled = True
        chkN4.Enabled = True
    End If
      
      hpsn = ""
      If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
      End If
      Dim checkhp As New ADODB.Recordset
      'Edited by mike 2010.06.11
       
      HP_pack_label = False
      sql = "select * from hp where charindex(h3c_bom_code,'" & Trim(txtSN.Text) & "')<>0 "
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      If Not rec.EOF Then
        If rec("pack_label") = "Y" Then HP_pack_label = True
      End If
      If rec.State = 1 Then rec.Close

      If HP_pack_label = True Then

        If conn11.State = 1 Then
             conn11.Close
        End If
        
        strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
            'con13.ConnectionTimeout = 50
        conn11.Open ConnectionString:=strConn
'        conn11.Open
        sql = " SELECT component_SN from dc_component_sn where unit_key IN (SELECT unit_key from UNIT WHERE SERIAL_NUMBER = '" & Trim(txtSN.Text) & "')" & " AND Remark = 'HP'"
        checkhp.Open sql, conn11, adOpenKeyset, adLockReadOnly
        If checkhp.EOF = True Then
            MsgBox ("没有对应的HP条码！")
            txtSN.Text = ""
            txtSN.SetFocus
            checkhp.Close
            Exit Sub
        Else
            hpsn = checkhp.Fields(0)
            checkhp.Close
        End If

        If conn11.State = 1 Then
             conn11.Close
        End If
        
      End If
      
      '=========================================================================
            Dim con13 As ADODB.Connection
            Dim rs13 As ADODB.Recordset
            Dim com As ADODB.Command
            
            Dim part_number As String
            Dim part_revision As String
            Dim order_number As String

            Set con13 = New ADODB.Connection
            Set rs13 = New ADODB.Recordset
            strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
            'con13.ConnectionTimeout = 50
            con13.Open ConnectionString:=strConn
            Set com = New ADODB.Command
            com.ActiveConnection = con13
            'str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtSN.Text) & "'"
            'str = " select top 1 a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "'"
            str = " select top 1 part_number,part_revision,creation_time,order_number from (" & _
            "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
            "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
            "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK') as t where order_number is not null order by t.creation_time desc "
            com.CommandText = str
            rs13.Open Source:=com
            'rs13.Open str
            If rs13.EOF = True Then
               
                MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
                rs13.Close
                cmdCancel_Click
                Exit Sub
               
            Else
                txtHV.Text = rs13.Fields(1)
                part_number = rs13.Fields(0)
                part_revision = rs13.Fields(1)
                order_number = rs13.Fields(3)
            End If
            If rs13.State = 1 Then
                rs13.Close
            End If
            If con13.State = 1 Then
                con13.Close
            End If
            
            'add by allen yan 2014/05/20
            'the main purpose of this function is to block the ECO versions that are disabled.
            If IsValidECOVersion(part_number, Me.txtHV.Text) = False Then
                cmdCancel_Click
                Exit Sub
            End If

       
      
      
      'Dim rcSet As New ADODB.Recordset
      'sql = "select top 1 * from revset where model='" & Mid(txtSN.Text, 3, 8) & "' and firstall<='" & txtSN.Text & "' and endall>='" & txtSN.Text & "'order by ver desc"
      'rcSet.Open sql, conn, adOpenKeyset, adLockOptimistic
      'If rcSet.EOF Then
      '   rcSet.Close
      'Else
      '  txtHV.Text = rcSet.Fields(3)
      'End If
      'If rcSet.State = 1 Then
      '   rcSet.Close
      'End If
      
      '===============================
      
        '=========================================================================
       '============add by carson start for TR5=============
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

        stringSQL = " select TOP 1 'SZ' from C_NoTR5_Part where EFFE_FLAG='1' AND  Part_Number ='" & Mid(txtSN.Text, 3, 8) & "'  "

        rsSZ.Open stringSQL
        If rsSZ.EOF = True Then
            txtSZ.Text = ""
        Else
            txtSZ.Text = rsSZ.Fields(0)
        End If
        rsSZ.Close
      '============add by carson end  =============
      
      '============add by ben start=============
        Dim con4 As ADODB.Connection
        Dim rs4 As ADODB.Recordset
        Dim con5 As ADODB.Connection
        Dim rs5 As ADODB.Recordset
        
        Dim flagTaskOrder As Boolean, flagHasBOM As Boolean, stringBOM As String
        
        Set con5 = New ADODB.Connection
        Set rs5 = New ADODB.Recordset
        con5.ConnectionString = "Provider=SQLOLEDB;User ID=sa;PWD=Flash123;Initial Catalog=afg_active_90;Data Source=10.11.1.130"
        con5.ConnectionTimeout = 50
        con5.Open
'        Dim stringSQL As String
        
        Set rs5.ActiveConnection = con5
        rs5.CursorType = adOpenForwardOnly
    
        stringSQL = "select 1 from UNIT a left join UDA_Unit b on a.unit_key = b.object_key "
        stringSQL = stringSQL + "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK' and a.serial_number = '" + UCase(Trim(Me.txtSN.Text)) + "'"
        
        If rs5.State = 1 Then rs5.Close
        rs5.Open stringSQL
        If rs5.EOF = True Then
            flagTaskOrder = True
        Else
            flagTaskOrder = False
        End If
        
        If flagTaskOrder = True Then
            txtOS.Text = ""
        Else
'            stringSQL = "select A.bom_name from BOM as A where A.bom_name in ( " & _
'            "select W.order_number + '_' + U.part_number from unit as U with(nolock) " & _
'            "left join work_order as W with(nolock) " & _
'            "on U.order_key = W.order_key where U.serial_number = '" & txtSN.Text & "' " & _
'                 ")  or A.bom_name in (" & _
'            "select '_DEL_' + W.order_number + '_' + U.part_number from unit as U with(nolock) " & _
'            "left join work_order as W with(nolock) " & _
'            "on U.order_key = W.order_key  where U.serial_number = '" & txtSN.Text & "' " & _
'            ") "
            stringSQL = "select A.bom_name from BOM as A where A.bom_name in ('" & order_number & "_" & part_number & "')  or A.bom_name in (" & _
            "'_DEL_" & order_number & "_" & part_number & "') "
            If rs5.State = 1 Then rs5.Close
            rs5.Open stringSQL
            If rs5.EOF = False Then
                flagHasBOM = True
                stringBOM = rs5.Fields(0)
            Else
                flagHasBOM = False
            End If
            
            If flagHasBOM = True Then
'
                stringSQL = " select C.size_of_part from [BOM] as A with (nolock) left join [BOM_PART_LIST] as B with (nolock) " & _
                "on A.bom_key = B.bom_key left join [BOM_PART_3003] as C with (nolock) " & _
                "on B.part_number = C.part_number where A.bom_name = '" & stringBOM & "' and C.size_of_part is not null "
                If rs5.State = 1 Then rs5.Close
                rs5.Open stringSQL
                If rs5.EOF = True Then
                    txtOS.Text = ""
                Else
                    txtOS.Text = rs5.Fields(0)
                End If
            Else
                If reprint = True Then
                    
                Else
                    MsgBox "此正常品缺少工单BOM，禁止打印!"
                    txtSN.Text = ""
                    txtSN.SetFocus
                    rs4.Close
                    Exit Sub
                End If
            End If
            
        End If

      '============add by ben end  =============
      

      
      '===============================
      
      Dim rcDavid As New ADODB.Recordset
      sql = "select case when Print_SV = 1 then 'Y' else 'N' end from tblH3CNew where  Part_Number ='" & Mid(txtSN.Text, 3, 8) & "'  and Part_Revision = '" & txtHV.Text & "'"
      rcDavid.Open sql, conn, adOpenKeyset, adLockReadOnly
      
      If rcDavid.EOF Then
            MsgBox "此产品序号未收集版本!"
            txtSN.Text = ""
            txtSN.SetFocus
            rcDavid.Close
            Exit Sub
      Else
            If rcDavid.Fields(0) = "N" Then
                txtVer.Text = "N/A"
            Else
                '--------------
                Set con = New ADODB.Connection
                con.CursorLocation = adUseClient
                con.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
                con.ConnectionTimeout = 100
                
                
                sql = "select * from tblSoftVersion where model='" & Mid(txtSN.Text, 3, 8) & "'"
    
                If con.State = 1 Then
                    con.Close
                End If
   
                con.Open
    
                Set rs3 = New ADODB.Recordset
                rs3.ActiveConnection = con
                rs3.Open sql, con, adOpenKeyset, adLockReadOnly
                
                If rs3.EOF Then
                    MsgBox "此产品序号未进行发货标签软件版本维护!"
                    txtSN.Text = ""
                    txtSN.SetFocus
                    rs3.Close
                    rcDavid.Close
                    Exit Sub
                Else
                    If rs3.Fields("searchFlag") = "Y" Then
                        Set con2 = New ADODB.Connection
                        con2.CursorLocation = adUseClient
                        con2.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=dataT"
                        con2.ConnectionTimeout = 100
                        
                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (rtrim(remark) <> '' and left(equipment,3) <> 'MTP' and remark is not null AND testtime >= dateadd(month,-3,getdate())) ORDER BY testtime DESC "
'                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ_ATE where barcode='" & Trim(txtSN.Text) & "' AND (ISNULL(remark, '') <> '') ORDER BY testtime DESC "
                        If con2.State = 1 Then
                            con2.Close
                        End If
                        con2.Open
                        Set rs2 = New ADODB.Recordset
                        rs2.ActiveConnection = con2
                        rs2.Open sql, con2, adOpenKeyset, adLockReadOnly
                        If rs2.EOF Then
                            MsgBox "查询软件版本资料时错误!"
                            txtSN.Text = ""
                            txtSN.SetFocus
                            rs2.Close
                            rs3.Close
                            rcDavid.Close
                            Exit Sub
                        Else
                            Dim stmp As String
                            Dim stmp2 As String
                            Dim stmp3 As String
                            Dim nowver As String
                            Dim beforver As String
                            Dim endDate As String
                            
                            'stmp2,stmp3 is ME settings sv
                            stmp2 = rs3.Fields("nowVer")
                            stmp3 = rs3.Fields("beforeVer")
                            
                            nowver = Mid(stmp2, 2)
                            beforver = Mid(stmp3, 2)
                            nowver = get_ver(nowver)
                            beforver = get_ver(beforver)
                            
                            endDate = rs3.Fields("endDate")
                            
                            'stmp is test sv
                            stmp = rs2.Fields("remark")
'update by allen.yan for the DongXu 2014/10/9
'exactly match first, if not then try faintly match
'先采用精确匹配，若匹配不成功，使用模糊匹配
                            If stmp2 = stmp Then
                                txtVer.Text = stmp
                            ElseIf stmp = smtp3 Then
                                If DateDiff("d", Now, CDate(endDate)) < 0 Then
                                    MsgBox "查询软件版本资料时错误(超过有效期)!"
                                    txtSN.Text = ""
                                    txtSN.SetFocus
                                    rs2.Close
                                    rs3.Close
                                    rcDavid.Close
                                    Exit Sub
                                Else
                                    txtVer.Text = stmp3
                                End If
                            ElseIf InStr(stmp, nowver) > 0 Then
                                Dim ttt As String
                                '查询测试记录中，ME维护的软件版本后的下一位字符
                                ttt = get_nextchar(stmp, nowver)
                                
                                If ttt = "L" Or ttt = "P" Then
                                    MsgBox "查询软件版本资料时错误(版本匹配错误)!"
                                    txtSN.Text = ""
                                    txtSN.SetFocus
                                    rs2.Close
                                    rs3.Close
                                    rcDavid.Close
                                    Exit Sub
                                Else
                                    txtVer.Text = stmp2
                                End If
                                    
                            Else '''''''''stmp2 = stmp
                                If Trim(beforver) = "" Then
                                    MsgBox "查询软件版本资料时错误(版本匹配错误)!"
                                    txtSN.Text = ""
                                    txtSN.SetFocus
                                    rs2.Close
                                    rs3.Close
                                    rcDavid.Close
                                    Exit Sub
                                Else
                                    '***********
                                    If InStr(stmp, beforver) > 0 Then
                                        Dim st As String
                                        st = get_nextchar(stmp, beforver)
                                        If st = "L" Or st = "P" Then
                                            MsgBox "查询软件版本资料时错误(版本匹配错误)!"
                                            txtSN.Text = ""
                                            txtSN.SetFocus
                                            rs2.Close
                                            rs3.Close
                                            rcDavid.Close
                                            Exit Sub
                                        Else
                                            If DateDiff("d", Now, CDate(endDate)) < 0 Then
                                                MsgBox "查询软件版本资料时错误(超过有效期)!"
                                                txtSN.Text = ""
                                                txtSN.SetFocus
                                                rs2.Close
                                                rs3.Close
                                                rcDavid.Close
                                                Exit Sub
                                            Else
                                                txtVer.Text = stmp3
                                            End If
                                        End If
    
                                    Else
                                            MsgBox "查询软件版本资料时错误(版本匹配错误)!"
                                            txtSN.Text = ""
                                            txtSN.SetFocus
                                            rs2.Close
                                            rs3.Close
                                            rcDavid.Close
                                            Exit Sub
                                    End If
                                    '**********
                                End If
                            End If ''''''''stmp2 = stmp end
                            
                        End If 'rs2 end
                        
                        rs2.Close
                        con2.Close
                        
                    Else '''''If rs3.Fields("searchFlag") = "Y" de else
                        If rs3.Fields("searchFlag") = "N" Then
    '=====================================================================
    
                            Dim stmp2_2 As String
                            Dim stmp3_2 As String
                            Dim endDate_2 As String
                            Dim nowver_2 As String
                            Dim beforver_2 As String
                            Dim stmp_2 As String
                            
                            stmp2_2 = rs3.Fields("nowVer")
                            stmp3_2 = rs3.Fields("beforeVer")
                            endDate_2 = rs3.Fields("endDate")
                            nowver_2 = Trim(stmp2_2)
                            beforver_2 = Trim(stmp3_2)

    
                            sql = "select top 1 ver from version where SN='" & txtSN.Text & "' order by testtime desc"
                            rec.Open sql, conn, adOpenKeyset, adLockReadOnly
                            If rec.EOF = True Then
                                MsgBox "此产品序号未收集版本!"
                                txtSN.Text = ""
                                txtSN.SetFocus
                                rec.Close
                                rs3.Close
                                rcDavid.Close
                    
                                Exit Sub
                            Else
                                Dim rcd As New ADODB.Recordset
                                sql = "select max(testtime) from version where sn='" & Trim(txtSN.Text) & "'"
                                rcd.Open sql, conn, adOpenKeyset, adLockReadOnly
                                If rcd.EOF = True Then
                                    MsgBox "此产品序号未收集版本!"
                                    txtSN.Text = ""
                                    txtSN.SetFocus
                                    rcd.Close
                                    rec.Close
                                    rs3.Close
                                    rcDavid.Close
                                    Exit Sub
                                Else
                                    Dim rs8 As New ADODB.Recordset
                                    sql = "select ver from version where testtime='" & rcd.Fields(0) & "' and sn='" & Trim(txtSN.Text) & "'"
                                    rs8.Open sql, conn, adOpenKeyset, adLockReadOnly
                                    If rs8.EOF = False Then
                    '               txtVer.Text = rs8.Fields(0)
                                        stmp_2 = rs8.Fields(0)
                                        If checkVersion(stmp_2, beforver_2, nowver_2, endDate_2) Then
                                            txtVer.Text = rs8.Fields(0)
                                        Else
                                            txtSN.Text = ""
                                            txtSN.SetFocus
                                            rs8.Close
                                            rcd.Close
                                            rec.Close
                                            rs3.Close
                                            rcDavid.Close
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox "此产品序号未收集版本!"
                                        txtSN.Text = ""
                                        txtSN.SetFocus
                                        rs8.Close
                                        rcd.Close
                                        rec.Close
                                        rs3.Close
                                        rcDavid.Close
                                        Exit Sub
                                    End If
                                    rs8.Close
                                End If 'rcd.EOF = True
                                rcd.Close
                            End If 'rec.EOF = True
                            rec.Close
      '==============================================
                        End If 'rs3.Fields("searchFlag") = "N"
                        
                    End If ''''''If rs3.Fields("searchFlag") = "Y" end
                    
                End If 'rs3.EOF
                
                rs3.Close
                con.Close
                
                '--------------
            End If 'rcDavid.Fields(0) = "N"
      End If 'rcDavid.EOF
      
      
      
      '===========================================================
      
       sql = "SELECT [ID],[Part_Number],[Part_Revision],[ProductID],[CPN],[EPN],[Des],[Size],[GW],[MS],[NAL1],[NAL1_Title],[NAL2],[NAL2_Title]" & _
        ",case when [CE] = 0 then 'Non CE' when CE = 1 then 'CE' when CE = 2 then 'CE+CE Addr' when CE = 3 then 'CE+HPE Addr' when CE = 4 then 'CE+HPE1 Addr' end as 'CE'" & _
        ",case when WEEE is null then 'N/A' when WEEE = 0 then 'No' when WEEE = 1 then 'Yes' end as 'WEEE'" & _
        ",case when ChinaRoHS is null then 'N/A' when ChinaRoHS = 0 then 'No' when ChinaRoHS = 1 then 'Yes' end as 'ChinaRoHS'" & _
        ",case when [RoHS] is null then 'N/A' when RoHS = 0 then 'No' when RoHS = 1 then 'Yes' end as 'RoHS'" & _
        ",case when [TurkeyRoHS] is null then 'N/A' when [TurkeyRoHS] = 0 then 'No' when TurkeyRoHS = 1 then 'Yes' end as '[TurkeyRoHS]'" & _
        ",case when Ukraine is null then 'N/A' when Ukraine = 0 then 'No' when Ukraine = 1 then 'Yes' end as 'Ukraine'" & _
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
        ",[5000_State],[Remark],case when EAC is null then 'N/A' when EAC = 0 then 'No' when EAC = 1 then 'Yes' end as 'EAC' FROM [Print].[dbo].[tblH3CNew] where Part_Number = '" & _
        Mid(txtSN.Text, 3, 8) & "' and Part_Revision ='" & txtHV.Text & "'"
      
      
      'sql = "select ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark from tblH3CNew where SN='" & Mid(txtSN.Text, 3, 8) & "' and HV='" & txtHV.Text & "'"
'       sql = "SELECT [ID],[Part_Number],[Part_Revision],[ProductID],[CPN],[EPN],[Des],[Size],[GW],[MS],[NAL1],[NAL1_Title],[NAL2],[NAL2_Title]" & _
'        ",[CE], WEEE,ChinaRoHS,[RoHS],[TurkeyRoHS],Ukraine,ATick,[ATick_ID],CTick,[CTick_ID],ICT,[ICT_ID],RCM,[RCM_ID],Gost,[Gost_ID],KC,[KC_ID],Print_SV,Print_Power,[5000_State],[Remark] FROM [Print].[dbo].[tblH3CNew] where Part_Number = '" & _
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
        'txtOS.Enabled = True
        'txtMN.Text = rec.Fields(5)
        'txtOS.BackColor = &HC0C0C0
        '============edit by ben start=============
        If flagTaskOrder = True Then
            txtOS.Text = rec.Fields(7)
        Else
            If flagHasBOM = True Then
                If txtOS.Text = "" Then
                    txtOS.Text = rec.Fields(7)
                Else
                    If Trim(txtOS.Text) <> Trim(rec.Fields(7)) Then
                        MsgBox "后台尺寸数据维护不一致,请找ME确认!"
                        txtSN.Text = ""
                        txtSN.SetFocus
                        rec.Close
                        Exit Sub
                    End If
                End If
            Else
                If reprint = True Then
                    txtOS.Text = rec.Fields(7)
                Else
                    MsgBox "此正常品缺少工单BOM，禁止打印!"
                    txtSN.Text = ""
                    txtSN.SetFocus
                    rec.Close
                    Exit Sub
                End If
            End If
        End If
'        txtOS.Text = rec.Fields(6)
        '============edit by ben end  =============
        If IsNull(rec.Fields(8)) = True Then
            MsgBox "此正常品缺少毛重数据，禁止打印!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
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
           chkHPE1.Value = 0
        ElseIf rec.Fields(14) = "Non CE" Then
           chkCE.Value = 0
           chkNonCE.Value = 1
           chkCEAddr.Value = 0
           chkHPEAddr.Value = 0
           chkHPE1.Value = 0
        ElseIf rec.Fields(14) = "CE+CE Addr" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
           chkCEAddr.Value = 1
           chkHPEAddr.Value = 0
           chkHPE1.Value = 0
        ElseIf rec.Fields(14) = "CE+HPE Addr" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
           chkCEAddr.Value = 0
           chkHPEAddr.Value = 1
           chkHPE1.Value = 0
        ElseIf rec.Fields(14) = "CE+HPE1 Addr" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
           chkCEAddr.Value = 0
           chkHPEAddr.Value = 0
           chkHPE1.Value = 1
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
        ElseIf rec.Fields(19) = "No" Or rec.Fields(19) = "No" Then
            Me.chkUkraine.Value = 0
            Me.chkNonUkraine.Value = 1
        End If
        If UCase(Trim(rec.Fields(20))) = "YES" Then
            Me.chkATick.Value = 1
            Me.chkNonATick.Value = 0
        Else
            Me.chkATick.Value = 0
            Me.chkNonATick.Value = 1
            Me.txtATick.Text = ""
        End If
        Me.txtATick.Text = Trim(rec.Fields(21))
        
        If UCase(Trim(rec.Fields(22))) = "YES" Then
            Me.chkCTick.Value = 1
            Me.chkNonCTick.Value = 0
        Else
            Me.chkCTick.Value = 0
            Me.chkNonCTick.Value = 1
            Me.txtCTick.Text = ""
        End If
        Me.txtCTick.Text = rec.Fields(23)
        
        If UCase(Trim(rec.Fields(24))) = "YES" Then
            Me.chkICT.Value = 1
            Me.chkNonICT.Value = 0
        Else
            Me.chkICT.Value = 0
            Me.chkNonICT.Value = 1
            Me.txtICT.Text = ""
        End If
        Me.txtICT.Text = rec.Fields(25)
        
        If UCase(Trim(rec.Fields(26))) = "YES" Then
            Me.chkRCM.Value = 1
            Me.chkNonRCM.Value = 0
        Else
            Me.chkRCM.Value = 0
            Me.chkNonRCM.Value = 1
            Me.txtRCM.Text = ""
        End If
        Me.txtRCM.Text = rec.Fields(27)
        
        If UCase(Trim(rec.Fields(28))) = "YES" Then
            Me.chkGost.Value = 1
            Me.chkNonGost.Value = 0
        Else
            Me.chkGost.Value = 0
            Me.chkNonGost.Value = 1
            Me.txtGost.Text = ""
        End If
        Me.txtGost.Text = rec.Fields(29)
        
         If UCase(Trim(rec.Fields(30))) = "YES" Then
            Me.chkKC.Value = 1
            Me.chkNonKC.Value = 0
        Else
            Me.chkKC.Value = 0
            Me.chkNonKC.Value = 1
            Me.txtKC.Text = ""
        End If
        Me.txtKC.Text = rec.Fields(31)
        
        If UCase(Trim(rec.Fields(32))) = "YES" Then
            Me.chkSVPrint.Value = 1
            Me.chkNonSVPrint.Value = 0
        Else
            Me.chkSVPrint.Value = 0
            Me.chkNonSVPrint.Value = 1
        End If
        
        If UCase(Trim(rec.Fields(33))) = "YES" Then
            Me.chkPCPrint.Value = 1
            Me.chkNonPCPrint.Value = 0
        Else
            Me.chkPCPrint.Value = 0
            Me.chkNonPCPrint.Value = 1
        End If
        
        Me.txt5000.Text = rec.Fields(34)
        
        txtHV.Text = rec.Fields(2)
        txtRemark.Text = rec.Fields(35)
        
        If UCase(Trim(rec.Fields(36))) = "YES" Then
           chkEAC.Value = 1
           chkNonEAC.Value = 0
        ElseIf rec.Fields(36) = "No" Or rec.Fields(36) = "N/A" Then
           chkEAC.Value = 0
           chkNonEAC.Value = 1
        End If
        
      End If
      '==================================================
       If rec.State = 1 Then
            rec.Close
       End If
       
       
       If chkY2.Value + chkY.Value + chkN.Value + chkN4.Value > 0 Then
           cmdPrint_Click
       Else
            MsgBox "Please select the value of PB"
       End If
       
   End If
   
  
   
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   
   'sql99 = "select order_number from work_order  where order_key in (select order_key from unit where serial_number='" & txtSN.Text & "')"
   'rec1.Open sql99, conn1, adOpenKeyset, adLockOptimistic
   'If rec1.EOF = True Then
   '     Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "H3C.lab")
   'Else
   '     If Trim(rec1.Fields(0) > "30000000") And Trim(rec1.Fields(0) < "40000000") Then
   '         Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "NPI-H3C.lab")
   '     Else
   '         Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "H3C.lab")
   '     End If
   'End If
   'rec1.Close

   'Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\H3C发货标签NEW\" & "H3CNEW 0401.lab")
   'Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\H3C发货标签NEW\" & "H3CNEW 0802.lab")
'   sql99 = "select top 1 1 from tblH3CNew  where FenXiao=1 and  Part_Number ='" & Mid(txtSN.Text, 3, 8) & "'  and Part_Revision = '" & txtHV.Text & "'"
    sql99 = "select top 1 isnull(FenXiao,0),isnull(Server,0) from tblH3CNew  where Part_Number ='" & Mid(txtSN.Text, 3, 8) & "'  and Part_Revision = '" & txtHV.Text & "'"
   rec1.Open sql99, conn, adOpenKeyset, adLockOptimistic
   If rec1.EOF = False Then
        If rec1.Fields(0) = True And rec1.Fields(1) = True Then
            MsgBox "分销产品和服务器产品维护同时选中，不允许打印，请联系ME"
            Set myDoc = myApp.Documents.Open("xx.lab")
        ElseIf rec1.Fields(0) = False And rec1.Fields(1) = False Then
            Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\H3C发货标签NEW\" & "H3CNEW 0802.lab")
        ElseIf rec1.Fields(0) = True And rec1.Fields(1) = False Then
            Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\H3C发货标签NEW\" & "分销H3CNEW.lab")
        ElseIf rec1.Fields(0) = False And rec1.Fields(1) = True Then
            Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\H3C发货标签NEW\" & "服务器产品H3CNEW.lab")
        Else 'never process
            MsgBox "分销产品和服务器产品维护错误，不允许打印，请联系ME"
            Set myDoc = myApp.Documents.Open("xx.lab")
        End If
   Else
        MsgBox "打印参数维护错误，不允许打印，请联系ME"
        Set myDoc = myApp.Documents.Open("xx.lab")
   End If
   rec1.Close
   
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub OpenLppx_hp()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP发货标签.lab")
   
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub txtSZ_Change()

End Sub
