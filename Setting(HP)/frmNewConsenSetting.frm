VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmNewConsenSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New HIK Setting"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   17325
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   17325
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fmH3C 
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   16935
      Begin VB.TextBox txtMaterial 
         Height          =   450
         Left            =   1800
         TabIndex        =   114
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtOrder 
         Height          =   450
         Left            =   5760
         TabIndex        =   112
         Top             =   960
         Width           =   2055
      End
      Begin VB.PictureBox Picture15 
         Height          =   495
         Left            =   8280
         Picture         =   "frmNewConsenSetting.frx":0000
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   110
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox chkNonCCC 
         BackColor       =   &H0000C000&
         Caption         =   "��"
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
         TabIndex        =   109
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox chkCCC 
         Caption         =   "��"
         Height          =   375
         Left            =   9240
         TabIndex        =   108
         Top             =   1920
         Width           =   735
      End
      Begin VB.PictureBox Picture14 
         Height          =   975
         Left            =   12240
         Picture         =   "frmNewConsenSetting.frx":0706
         ScaleHeight     =   915
         ScaleWidth      =   2355
         TabIndex        =   107
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox chkNonLaser 
         BackColor       =   &H0000C000&
         Caption         =   "��"
         Height          =   375
         Left            =   15840
         TabIndex        =   106
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox chkLaser 
         Caption         =   "��"
         Height          =   375
         Left            =   14880
         TabIndex        =   105
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtModel 
         Height          =   450
         Left            =   9840
         TabIndex        =   104
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtPower 
         Height          =   450
         Left            =   14520
         TabIndex        =   102
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtType 
         Height          =   450
         Left            =   5760
         TabIndex        =   100
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkEAC 
         Caption         =   "��"
         Height          =   375
         Left            =   14280
         TabIndex        =   98
         Top             =   5280
         Width           =   735
      End
      Begin VB.CheckBox chkNonEAC 
         BackColor       =   &H0000C000&
         Caption         =   "��"
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
         Left            =   15240
         TabIndex        =   97
         Top             =   5280
         Width           =   735
      End
      Begin VB.PictureBox Picture13 
         Height          =   495
         Left            =   13440
         Picture         =   "frmNewConsenSetting.frx":1DDB
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   96
         Top             =   5160
         Width           =   495
      End
      Begin VB.CheckBox chkHPEAddr 
         Caption         =   "HPE Addr"
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
         TabIndex        =   95
         Top             =   8280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtUkraine 
         Height          =   495
         Left            =   10440
         TabIndex        =   94
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkCEAddr 
         Caption         =   "CE Addr"
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
         TabIndex        =   70
         Top             =   7800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkRoHS 
         Caption         =   "��"
         Height          =   375
         Left            =   14280
         TabIndex        =   69
         Top             =   6000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkNonSVPrint 
         BackColor       =   &H0000C000&
         Caption         =   "��"
         Height          =   495
         Left            =   15120
         TabIndex        =   68
         Top             =   7800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkSVPrint 
         Caption         =   "��"
         Height          =   495
         Left            =   14280
         TabIndex        =   67
         Top             =   7800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkTurkeyRohs 
         Caption         =   "��"
         Height          =   375
         Left            =   14400
         TabIndex        =   66
         Top             =   8520
         Width           =   735
      End
      Begin VB.CheckBox chkChinaRoHS 
         Caption         =   "��"
         Height          =   375
         Left            =   1800
         TabIndex        =   65
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txXXXXXX 
         Height          =   495
         Left            =   25200
         TabIndex        =   64
         Text            =   "sdgfdsfadsfadsf"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtHV 
         Height          =   495
         Left            =   -600
         TabIndex        =   63
         Top             =   7560
         Width           =   1935
      End
      Begin VB.TextBox txtMS 
         Height          =   450
         Left            =   8880
         TabIndex        =   62
         Top             =   4200
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CheckBox chkWEEE 
         Caption         =   "��"
         Height          =   375
         Left            =   5160
         TabIndex        =   61
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox chkCE 
         Caption         =   "CE"
         Height          =   375
         Left            =   3960
         TabIndex        =   60
         Top             =   8160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtGW 
         Height          =   450
         Left            =   9000
         TabIndex        =   59
         Top             =   3720
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtSize 
         Height          =   450
         Left            =   5160
         TabIndex        =   58
         Top             =   3720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtEPN 
         Height          =   450
         Left            =   1680
         TabIndex        =   57
         Top             =   3720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtCPN 
         Height          =   450
         Left            =   8880
         TabIndex        =   56
         Top             =   7440
         Width           =   2775
      End
      Begin VB.TextBox txtSN 
         Height          =   450
         Left            =   1800
         TabIndex        =   55
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtProductID 
         Height          =   495
         Left            =   3480
         TabIndex        =   54
         Top             =   7320
         Width           =   2895
      End
      Begin VB.TextBox txtNAL2Title 
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
         Left            =   1680
         TabIndex        =   53
         Top             =   5040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtNAL2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   52
         Top             =   5040
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox cb5000 
         Height          =   450
         Left            =   7680
         TabIndex        =   51
         Top             =   8280
         Width           =   2895
      End
      Begin VB.CheckBox chkPWPrint 
         Caption         =   "��"
         Height          =   330
         Left            =   14280
         TabIndex        =   50
         Top             =   7200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkNonPWPrint 
         BackColor       =   &H0000C000&
         Caption         =   "��"
         Height          =   375
         Left            =   15120
         TabIndex        =   49
         Top             =   7200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   3240
         Picture         =   "frmNewConsenSetting.frx":3E5E
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   48
         Top             =   8040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkUkraine 
         Caption         =   "��"
         Height          =   375
         Left            =   8640
         TabIndex        =   47
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkNonUkraine 
         BackColor       =   &H0000C000&
         Caption         =   "��"
         Height          =   375
         Left            =   9480
         TabIndex        =   46
         Top             =   5160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H0000C000&
         Caption         =   "��"
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
         TabIndex        =   45
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H0000C000&
         Caption         =   "��"
         Height          =   375
         Left            =   2760
         TabIndex        =   44
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox chkNonTurkeyRohs 
         BackColor       =   &H0000C000&
         Caption         =   "��"
         Height          =   375
         Left            =   15360
         TabIndex        =   43
         Top             =   8520
         Width           =   735
      End
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H0000C000&
         Caption         =   "��"
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
         Left            =   15240
         TabIndex        =   42
         Top             =   6000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         Height          =   615
         Left            =   840
         Picture         =   "frmNewConsenSetting.frx":490C
         ScaleHeight     =   555
         ScaleWidth      =   675
         TabIndex        =   41
         Top             =   1800
         Width           =   735
      End
      Begin VB.PictureBox Picture3 
         Height          =   615
         Left            =   4320
         Picture         =   "frmNewConsenSetting.frx":5E72
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   40
         Top             =   1800
         Width           =   615
      End
      Begin VB.PictureBox Picture4 
         Height          =   735
         Left            =   10560
         Picture         =   "frmNewConsenSetting.frx":7804
         ScaleHeight     =   675
         ScaleWidth      =   3555
         TabIndex        =   39
         Top             =   8280
         Width           =   3615
      End
      Begin VB.PictureBox Picture5 
         Height          =   495
         Left            =   12720
         Picture         =   "frmNewConsenSetting.frx":F492
         ScaleHeight     =   435
         ScaleWidth      =   1395
         TabIndex        =   38
         Top             =   6000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox Picture7 
         Height          =   495
         Left            =   120
         Picture         =   "frmNewConsenSetting.frx":11830
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   37
         Top             =   5160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture8 
         Height          =   495
         Left            =   120
         Picture         =   "frmNewConsenSetting.frx":12586
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   36
         Top             =   5880
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture9 
         Height          =   615
         Left            =   120
         Picture         =   "frmNewConsenSetting.frx":12E50
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   35
         Top             =   6600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Picture10 
         Height          =   495
         Left            =   2880
         Picture         =   "frmNewConsenSetting.frx":140B2
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   34
         Top             =   9120
         Width           =   615
      End
      Begin VB.PictureBox Picture11 
         Height          =   495
         Left            =   6480
         Picture         =   "frmNewConsenSetting.frx":14C58
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   33
         Top             =   5880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Picture12 
         Height          =   615
         Left            =   6600
         Picture         =   "frmNewConsenSetting.frx":159BA
         ScaleHeight     =   555
         ScaleWidth      =   435
         TabIndex        =   32
         Top             =   5160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtATick 
         Height          =   495
         Left            =   3600
         TabIndex        =   31
         Top             =   5160
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtCTick 
         Height          =   495
         Left            =   3600
         TabIndex        =   30
         Top             =   5880
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtNAL1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   29
         Top             =   4440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H0000C000&
         Caption         =   "�� CE"
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
         Left            =   5160
         TabIndex        =   28
         Top             =   8160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtRCM 
         Height          =   495
         Left            =   6720
         TabIndex        =   27
         Top             =   9120
         Width           =   1815
      End
      Begin VB.TextBox txtNAL1Title 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   26
         Top             =   4440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtRemark 
         Height          =   495
         Left            =   8880
         TabIndex        =   25
         Top             =   4680
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.PictureBox Picture6 
         Height          =   615
         Left            =   7680
         Picture         =   "frmNewConsenSetting.frx":16620
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   24
         Top             =   5040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkNonATick 
         BackColor       =   &H0000C000&
         Caption         =   "��"
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
         Left            =   1680
         TabIndex        =   23
         Top             =   5160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkATick 
         Caption         =   "��"
         Height          =   375
         Left            =   840
         TabIndex        =   22
         Top             =   5160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkNonCTick 
         BackColor       =   &H0000C000&
         Caption         =   "��"
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
         Left            =   1680
         TabIndex        =   21
         Top             =   6000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkCTick 
         Caption         =   "��"
         Height          =   375
         Left            =   840
         TabIndex        =   20
         Top             =   6000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkNonICT 
         BackColor       =   &H0000C000&
         Caption         =   "��"
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
         Left            =   1680
         TabIndex        =   19
         Top             =   6720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkICT 
         Caption         =   "��"
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   6720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtICT 
         Height          =   495
         Left            =   3600
         TabIndex        =   17
         Top             =   6600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox chkNonKC 
         BackColor       =   &H0000C000&
         Caption         =   "��"
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
         Left            =   8160
         TabIndex        =   16
         Top             =   6600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkKC 
         Caption         =   "��"
         Height          =   375
         Left            =   7320
         TabIndex        =   15
         Top             =   6600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkNonGost 
         BackColor       =   &H0000C000&
         Caption         =   "��"
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
         Left            =   8160
         TabIndex        =   14
         Top             =   5880
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkGost 
         Caption         =   "��"
         Height          =   375
         Left            =   7320
         TabIndex        =   13
         Top             =   5880
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkNonRCM 
         BackColor       =   &H0000C000&
         Caption         =   "��"
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
         Left            =   4560
         TabIndex        =   12
         Top             =   9120
         Width           =   735
      End
      Begin VB.CheckBox chkRCM 
         Caption         =   "��"
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   9120
         Width           =   735
      End
      Begin VB.TextBox txtGost 
         Height          =   495
         Left            =   10320
         TabIndex        =   10
         Top             =   5880
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtKC 
         Height          =   495
         Left            =   10320
         TabIndex        =   9
         Top             =   6600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "���ϴ���:"
         Height          =   375
         Left            =   360
         TabIndex        =   113
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "���۶���:"
         Height          =   375
         Left            =   4200
         TabIndex        =   111
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "��Ʒ�ͺ�:"
         Height          =   375
         Left            =   8280
         TabIndex        =   103
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "��Դά��:"
         Height          =   375
         Left            =   12840
         TabIndex        =   101
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "��Ʒ���:"
         Height          =   375
         Left            =   4200
         TabIndex        =   99
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblPrintSV 
         Caption         =   "��ӡ����汾:"
         Height          =   495
         Left            =   12240
         TabIndex        =   93
         Top             =   7800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblOS 
         Caption         =   "�ߴ�(mm):"
         Height          =   375
         Left            =   3720
         TabIndex        =   92
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblRemark 
         Caption         =   "��ע:"
         Height          =   495
         Left            =   24120
         TabIndex        =   91
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblHV 
         Caption         =   "Ӳ���汾:"
         Height          =   495
         Left            =   -1920
         TabIndex        =   90
         Top             =   7560
         Width           =   1455
      End
      Begin VB.Label lblMS 
         Caption         =   "�����׼:"
         Height          =   375
         Left            =   7200
         TabIndex        =   89
         Top             =   4200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblGW 
         Caption         =   "ë��(kg):"
         Height          =   375
         Left            =   7320
         TabIndex        =   88
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblEPN 
         Caption         =   "Ӣ������:"
         Height          =   375
         Left            =   240
         TabIndex        =   87
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblCPN 
         Caption         =   "��������:"
         Height          =   375
         Left            =   7440
         TabIndex        =   86
         Top             =   7440
         Width           =   1335
      End
      Begin VB.Label lblSN 
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   360
         TabIndex        =   85
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "��Ʒ����:"
         Height          =   495
         Left            =   1800
         TabIndex        =   84
         Top             =   7320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "�����ͺ�1"
         Height          =   375
         Left            =   240
         TabIndex        =   83
         Top             =   4440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "����1:"
         Height          =   375
         Left            =   3840
         TabIndex        =   82
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "�����ͺ�2"
         Height          =   375
         Left            =   240
         TabIndex        =   81
         Top             =   5040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "����2:"
         Height          =   375
         Left            =   3840
         TabIndex        =   80
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "5000��״̬:"
         Height          =   495
         Left            =   7560
         TabIndex        =   79
         Top             =   7920
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "��ӡ��Դ����:"
         Height          =   375
         Left            =   12240
         TabIndex        =   78
         Top             =   7200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "ATick ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   77
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "CTick ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   76
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "RCM ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   75
         Top             =   9120
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Gost ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   74
         Top             =   6000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "KC ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   73
         Top             =   6720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "ICT ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   72
         Top             =   6720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "��ע:"
         Height          =   495
         Left            =   7560
         TabIndex        =   71
         Top             =   4680
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgH3C 
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   10186
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14400
      TabIndex        =   6
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   12840
      TabIndex        =   5
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "ȷ��(Confirm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11160
      TabIndex        =   4
      Top             =   9840
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(Delete)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14400
      TabIndex        =   3
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�޸�(Update)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   12840
      TabIndex        =   2
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "����(Insert)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11160
      TabIndex        =   1
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "��ѯ(Query)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9480
      TabIndex        =   0
      Top             =   9480
      Width           =   1455
   End
End
Attribute VB_Name = "frmNewConsenSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim op As String
Dim xlApp As New Excel.Application
Dim xlBook As New Excel.Workbook
Dim xlSheet As New Excel.Worksheet
Dim query As Boolean

Private Sub Reset()
    Dim ctr As Control

    If op = "Insert" Then
        cmdQuery.Enabled = True
        cmdInsert.Enabled = False
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
        cmdConfirm.Enabled = True
        cmdCancel.Enabled = True
        For Each ctr In Me.Controls
            If TypeOf ctr Is TextBox Then
                    ctr.Text = ""
                    ctr.Enabled = True
                    ctr.BackColor = &HFFFFFF
                 ElseIf TypeOf ctr Is ComboBox Then
                    If ctr.Style = 2 Then
                       ctr.ListIndex = -1
                    Else
                       ctr.Text = ""
                    End If
                 ElseIf TypeOf ctr Is CheckBox Then
                    ctr.Value = 0
            End If
        Next
        Me.txtSize.Text = "N/A"
        Me.txtRemark = "N/A"
        Me.chkChinaRoHS.Value = 0
        Me.chkTurkeyRohs.Value = 1
        Me.chkRoHS.Value = 1
        Me.chkLaser.Value = 0
        Me.chkCCC.Value = 0
        Me.chkWEEE.Value = 0
        Me.chkEAC.Value = 1
        Me.txtNAL2.Text = "/"
        Me.txtNAL2Title = "/"
        Me.txtSN.Enabled = True
        Me.txtSN.SetFocus
    ElseIf op = "Cancel" Then
        For Each ctr In Me.Controls
        If TypeOf ctr Is TextBox Then
                ctr.Enabled = False
                ctr.BackColor = &HFFFFFF
             ElseIf TypeOf ctr Is ComboBox Then
                ctr.Enabled = False
                ctr.BackColor = &HFFFFFF
             ElseIf TypeOf ctr Is CheckBox Then
                ctr.Enabled = True
        End If
    Next
        cmdQuery.Enabled = True
        cmdInsert.Enabled = True
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
        cmdConfirm.Enabled = True
        cmdCancel.Enabled = False
    ElseIf op = "Update" Then
        For Each ctr In Me.Controls
        If TypeOf ctr Is TextBox Then
                ctr.Enabled = True
             ElseIf TypeOf ctr Is ComboBox Then
                ctr.Enabled = True
                End If
    Next
        txtSN.Enabled = False
        txtSN.BackColor = &HE0E0E0
    End If
End Sub
Private Sub enable()
   txtSN.Enabled = True
   txtSN.BackColor = &HFFFFFF
   txtCPN.Enabled = True
   txtCPN.BackColor = &HFFFFFF
   txtEPN.Enabled = True
   txtEPN.BackColor = &HFFFFFF
   txtDes.Enabled = True
   txtDes.BackColor = &HFFFFFF
   txtSize.Enabled = True
   txtSize.BackColor = &HFFFFFF
   txtGW.Enabled = True
   txtGW.BackColor = &HFFFFFF
   
   chkCE.Enabled = True
   chkCEAddr.Enabled = True
   chkHPEAddr.Enabled = True
      
   chkNonCE.Enabled = True

   chkEAC.Enabled = True
   chkNonEAC.Enabled = True

   chkTurkeyRohs.Enabled = True
   chkNonTurkeyRohs.Enabled = True
   
   chkRoHS.Enabled = True
   chkNonRoHS.Enabled = True
   
   
   chkCCC.Enabled = True
   chkNonCCC.Enabled = True
   chkChinaRoHS.Enabled = True
   chkNonChinaRoHS.Enabled = True
   chkWEEE.Enabled = True
   chkNonWEEE.Enabled = True
   chkLaser.Enabled = True
   chkNonLaser.Enabled = True
   
   txtType.Enabled = True
   txtType.BackColor = &HFFFFFF
   txtModel.Enabled = True
   txtModel.BackColor = &HFFFFFF
   txtPower.Enabled = True
   txtPower.BackColor = &HFFFFFF
   txtMaterial.Enabled = True
   txtMaterial.BackColor = &HFFFFFF
   txtOrder.Enabled = True
   txtOrder.BackColor = &HFFFFFF

   'optH3CRoHS.Enabled = True
   'opt3COMRoHS.Enabled = True
   'optNonRoHS.Enabled = True
   
   
   txtMS.Enabled = True
   txtMS.BackColor = &HFFFFFF

   
   chkSVPrint.Enabled = True
   chkNonSVPrint.Enabled = True
   
   txtHV.Enabled = True
   txtHV.BackColor = &HFFFFFF
   txtRemark.Enabled = True
   txtRemark.BackColor = &HFFFFFF
   
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = False
   cmdUpdate.Enabled = False
   cmdDelete.Enabled = False
   cmdConfirm.Enabled = True
   cmdCancel.Enabled = True
End Sub

Private Sub unable()
   
   

End Sub

Private Sub chkCCC_Click()
    If chkCCC.Value = 1 Then
        chkNonCCC.Value = 0
    Else
        chkNonCCC.Value = 1
    End If
End Sub


Private Sub chkNonCCC_Click()
    If chkNonCCC.Value = 1 Then
        chkCCC.Value = 0
    Else
        chkCCC.Value = 1
    End If
End Sub
Private Sub chkChinaROHS_Click()
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

Private Sub chkNonWEEE_Click()
   If chkNonWEEE.Value = 1 Then
      chkWEEE.Value = 0
   Else
      chkWEEE.Value = 1
   End If
End Sub

Private Sub chkNonChinaRoHS_Click()
   If chkNonChinaRoHS.Value = 1 Then
      chkChinaRoHS.Value = 0
   Else
      chkChinaRoHS.Value = 1
   End If
End Sub

Private Sub chkLaser_Click()
    If chkLaser.Value = 1 Then
        chkNonLaser.Value = 0
    Else
        chkNonLaser.Value = 1
    End If
End Sub
Private Sub chkNonLaser_Click()
    If chkNonLaser.Value = 1 Then
        chkLaser.Value = 0
    Else
        chkLaser.Value = 1
    End If
End Sub


Private Sub chkATick_Click()
    If Me.chkATick.Value = 1 Then
        Me.chkNonATick.Value = 0
        Me.txtATick.Enabled = True
        Me.txtATick.Text = "N279"
    End If
End Sub



Private Sub chkCE_Click()
   If chkCE.Value = 1 Then
      chkNonCE.Value = 0
   Else
      chkNonCE.Value = 1
   End If
End Sub

Private Sub chkCEAddr_Click()
   If chkCEAddr.Value = 1 Then
      chkCE.Value = 1
      chkNonCE.Value = 0
      chkHPEAddr.Value = 0
   End If
End Sub
Private Sub chkHPEAddr_Click()
   If chkHPEAddr.Value = 1 Then
      chkCE.Value = 1
      chkNonCE.Value = 0
      chkCEAddr.Value = 0
   End If
End Sub

Private Sub chkCTick_Click()
    If Me.chkCTick.Value = 1 Then
        Me.chkNonCTick = 0
        Me.txtCTick.Enabled = True
        Me.txtCTick.Text = "N151"
    End If
End Sub

Private Sub chkGost_Click()
    If chkGost.Value = 1 Then
        Me.chkNonGost.Value = 0
        Me.txtGost.Enabled = True
        Me.txtGost.Text = "ME77"
    End If
End Sub



Private Sub chkICT_Click()
    If Me.chkICT.Value = 1 Then
        Me.chkNonICT.Value = 0
        Me.txtICT.Enabled = True
        Me.txtICT.Text = "B00502010"
    End If
End Sub

Private Sub chkKC_Click()
    If Me.chkKC.Value = 1 Then
        Me.chkNonKC.Value = 0
        Me.txtKC.Enabled = True
    End If
End Sub

Private Sub chkNonATick_Click()
    If Me.chkNonATick.Value = 1 Then
        Me.chkATick.Value = 0
        Me.txtATick.Enabled = False
        Me.txtATick.Text = ""
    End If
    
End Sub



Private Sub chkNonCE_Click()
   If chkNonCE.Value = 1 Then
      chkCE.Value = 0
      chkCEAddr.Value = 0
      chkHPEAddr.Value = 0
   Else
      chkCE.Value = 1
   End If
End Sub



Private Sub chkNonCTick_Click()
    If Me.chkNonCTick.Value = 1 Then
        Me.chkCTick.Value = 0
        Me.txtCTick.Text = ""
        Me.txtCTick.Enabled = False
    End If
End Sub

Private Sub chkNonGost_Click()
    If Me.chkNonGost.Value = 1 Then
        Me.chkGost.Value = 0
        Me.txtGost.Text = ""
        Me.txtGost.Enabled = False
    End If
End Sub

Private Sub chkNonICT_Click()
    If Me.chkNonICT.Value = 1 Then
        Me.chkICT.Value = 0
        Me.txtICT.Text = ""
        Me.txtICT.Enabled = False
    End If
End Sub

Private Sub chkNonKC_Click()
    If Me.chkNonKC.Value = 1 Then
        Me.chkKC.Value = 0
        Me.txtKC.Text = ""
        Me.txtKC.Enabled = False
    End If
End Sub

Private Sub chkNonPWPrint_Click()
    If Me.chkNonPWPrint.Value = 1 Then
        Me.chkPWPrint.Value = 0
    End If
End Sub

Private Sub chkNonRCM_Click()
    If Me.chkNonRCM.Value = 1 Then
        Me.chkRCM.Value = 0
        Me.txtRCM.Text = ""
        Me.txtRCM.Enabled = False
    End If
End Sub

Private Sub chkNonSVPrint_Click()
    If Me.chkNonSVPrint.Value = 1 Then
        Me.chkSVPrint.Value = 0
    End If
End Sub

Private Sub chkNonUkraine_Click()
    If Me.chkNonUkraine.Value = 1 Then
        Me.chkUkraine.Value = 0
    End If
End Sub


Private Sub chkPWPrint_Click()
    If Me.chkPWPrint.Value = 1 Then
        Me.chkNonPWPrint.Value = 0
    End If
End Sub

Private Sub chkRCM_Click()
    If Me.chkRCM.Value = 1 Then
        Me.chkNonRCM.Value = 0
        Me.txtRCM.Enabled = True
        Me.txtRCM.Text = "N279"
    End If
End Sub

Private Sub chkRoHS_Click()
    If chkRoHS.Value = 1 Then
        chkNonRoHS.Value = 0
    Else
        chkNonRoHS.Value = 1
    End If
End Sub

Private Sub chkNonRoHS_Click()
    If chkNonRoHS.Value = 1 Then
        chkRoHS.Value = 0
    Else
        chkRoHS.Value = 1
    End If
End Sub

Private Sub chkSVPrint_Click()
   If chkSVPrint.Value = 1 Then
      chkNonSVPrint.Value = 0
   End If
End Sub

Private Sub chkUkraine_Click()
    If Me.chkUkraine.Value = 1 Then
        Me.chkNonUkraine.Value = 0
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

Private Sub chkTurkeyROHS_Click()
    If chkTurkeyRohs.Value = 1 Then
        chkNonTurkeyRohs.Value = 0
    Else
        chkNonTurkeyRohs.Value = 1
    End If
End Sub

Private Sub chkNonTurkeyRohs_Click()
    If chkNonTurkeyRohs.Value = 1 Then
        chkTurkeyRohs.Value = 0
    Else
        chkTurkeyRohs.Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
   op = "Cancel"
   Reset
End Sub

Private Sub cmdConfirm_Click()
   If Trim(txtSN.Text) = "" Then
      MsgBox "��Ʒ���벻��Ϊ��!!", vbExclamation + vbOKOnly, "��Ʒ�����"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If Trim(txtType.Text) = "" Then
      MsgBox "��Ʒ�����Ϊ��!!", vbExclamation + vbOKOnly, "��Ʒ����"
      txtType.SetFocus
      Exit Sub
   End If
   
   If Trim(txtModel.Text) = "" Then
      MsgBox "��Ʒ�ͺŲ���Ϊ��!!", vbExclamation + vbOKOnly, "��Ʒ�ͺſ�"
      txtModel.SetFocus
      Exit Sub
   End If
   
'   If Trim(txtPower.Text) = "" Then
'      MsgBox "��Դά������Ϊ��!!", vbExclamation + vbOKOnly, "��Դά����"
'      txtPower.SetFocus
'      Exit Sub
'   End If
   If Trim(txtMaterial.Text) = "" Then
      MsgBox "���ϴ��벻��Ϊ��!!", vbExclamation + vbOKOnly, "���ϴ����"
      txtMaterial.SetFocus
      Exit Sub
   End If
'   If Trim(txtOrder.Text) = "" Then
'      MsgBox "���۶�������Ϊ��!!", vbExclamation + vbOKOnly, "���۶�����"
'      txtOrder.SetFocus
'      Exit Sub
'   End If
'   If txtCPN.Text = "" Then
'       MsgBox "��Ʒ����(����)����Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ����(����)��"
'       txtCPN.SetFocus
'       Exit Sub
'   End If
'   If txtEPN.Text = "" Then
'      MsgBox "��Ʒ����(Ӣ��)����Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ����(Ӣ��)��"
'      txtEPN.SetFocus
'      Exit Sub
'   End If
'   If txtProductID.Text = "" Then
'      MsgBox "��Ʒ��������Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ������"
'      txtProductID.SetFocus
'      Exit Sub
'   End If
'   If txtSize.Text = "" Then
'      MsgBox "��ߴ粻��Ϊ��!", vbExclamation + vbOKOnly, "��ߴ��"
'      txtSize.SetFocus
'      Exit Sub
'   End If
'   If txtSize.Text = "/" Then
'      MsgBox "����ߴ���ά��N/A!", vbExclamation + vbOKOnly, "����ߴ�"
'      txtSize.SetFocus
'      Exit Sub
'   End If
'   If txtSize.Text = "n/a" Then
'      txtSize.Text = UCase(txtSize.Text)
'   End If

'   If txtSize.Text <> "N/A" Then
'
'        txtSize.Text = LTrim(RTrim(txtSize.Text))
'
'        If Right(txtSize.Text, 2) <> "mm" Then
'            MsgBox "��ߴ��ʽ����!", vbExclamation + vbOKOnly, "��ߴ����"
'            txtSize.SetFocus
'            Exit Sub
'        End If
'
'        If InStr(txtSize.Text, "mmm") > 0 Then
'            MsgBox "��ߴ��ʽ����!", vbExclamation + vbOKOnly, "��ߴ����"
'            txtSize.SetFocus
'            Exit Sub
'        End If
'   End If
'
'
'   If Trim(txtGW.Text) <> "" Then
'        If UCase(Right(Trim(txtGW.Text), 2)) <> "KG" Then
'           MsgBox "ë�ر�����ϵ�λkg!", vbExclamation + vbOKOnly, "ë�ص�λ��"
'           txtGW.SetFocus
'           Exit Sub
'        End If
'        If Len(Me.txtGW.Text) < 6 Or Mid(Right(Trim(Me.txtGW.Text), 5), 1, 1) <> "." Then
'            MsgBox "ë�����ݳ���Ӧ����6λ���Ұ���С���㣬��x.xxkg!", vbExclamation + vbOKOnly, "ë�ظ�ʽ����ȷ"
'            txtGW.SetFocus
'            Exit Sub
'        End If
'   End If
   
  
   
   
'   If txtMS.Text = "" Then
'      MsgBox "�����׼����Ϊ��!", vbExclamation + vbOKOnly, "�����׼��"
'      txtMS.SetFocus
'      Exit Sub
'   End If
'
'   If txtHV.Text = "" Then
'      MsgBox "Ӳ���汾����Ϊ��!", vbExclamation + vbOKOnly, "Ӳ���汾��"
'      txtHV.SetFocus
'      Exit Sub
'   End If
'   If chkSVPrint.Value = 0 And chkNonSVPrint.Value = 0 Then
'      MsgBox "�Ƿ��ӡ����汾����Ϊ��!", vbExclamation + vbOKOnly, "������汾��"
'      txtHV.SetFocus
'      Exit Sub
'   End If
'
'   If (Trim(Me.txtNAL1.Text) = "/" And Trim(Me.txtNAL1Title.Text) = "/") Or (Trim(Me.txtNAL1.Text) <> "/" And Trim(Me.txtNAL1Title.Text) <> "/") Then
'   Else
'       MsgBox "����1���ݸ�ʽ����ȷ!", vbExclamation + vbOKOnly, "����1���ݸ�ʽ����ȷ"
'       Exit Sub
'   End If
'
'   If (Trim(Me.txtNAL2.Text) = "/" And Trim(Me.txtNAL2Title.Text) = "/") Or (Trim(Me.txtNAL2.Text) <> "/" And Trim(Me.txtNAL2Title.Text) <> "/") Then
'   Else
'       MsgBox "����2���ݸ�ʽ!", vbExclamation + vbOKOnly, "����2���ݸ�ʽ����ȷ"
'       Exit Sub
'   End If
'
'   If (Trim(Me.txtNAL2.Text) <> "/" And Trim(Me.txtNAL2Title.Text) <> "/") And (Trim(Me.txtNAL1.Text) = "/" Or Trim(Me.txtNAL1Title.Text) = "/") Then
'      MsgBox "�����������Ƚ�����1��������", vbExclamation + vbOKOnly, "������д����1�������Ϣ"
'      Exit Sub
'   End If
'
'
'   If (Len(Trim(txtNAL1.Text)) <> 14 And InStr(1, txtNAL1.Text, "-") < 0) Then
'        MsgBox "�������ݸ�ʽ����ȷ����ȷ��!", vbExclamation + vbOKOnly, "�������ݸ�ʽ����ȷ"
'        txtNAL1.SetFocus
'        Exit Sub
'   End If
'
'   If (Len(Trim(txtNAL2.Text)) <> 14 And InStr(1, txtNAL2.Text, "-") < 0) And Trim(txtNAL2.Text) <> "/" Then
'        MsgBox "�������ݸ�ʽ����ȷ����ȷ��!", vbExclamation + vbOKOnly, "�������ݸ�ʽ����ȷ"
'        txtNAL2.SetFocus
'        Exit Sub
'   End If
   
   
   Dim CE, WEEE, ChinaRoHS, RoHS, TurkeyRoHS, SVPrint, ATick, CTick, ICT, RCM, Gost, KC, ATick_ID, CTick_ID, ICT_ID, RCM_ID, Gost_ID, KC_ID, PWPrint, EAC, CCC, Laser As String
   Dim ftStatus, NAL1, NAL2, NAL1Title, NAL2Title, Ukraine, Ukraine_ID As String
'   If Trim(Me.txtNAL1.Text) = "/" Or UCase(Trim(Me.txtNAL1.Text)) = "NA" Or UCase(Trim(Me.txtNAL1.Text)) = "N/A" Then
'      NAL1 = "/"
'   Else
'      NAL1 = Trim(Me.txtNAL1.Text)
'   End If
'
'   If Trim(Me.txtNAL2.Text) = "/" Or UCase(Trim(Me.txtNAL2.Text)) = "NA" Or UCase(Trim(Me.txtNAL2.Text)) = "N/A" Then
'      NAL2 = "/"
'   Else
'      NAL2 = Trim(Me.txtNAL2.Text)
'   End If
'   If Trim(Me.txtNAL1Title.Text) = "/" Or UCase(Trim(Me.txtNAL1Title.Text)) = "NA" Or UCase(Trim(Me.txtNAL1Title.Text)) = "N/A" Then
'      NAL1Title = "/"
'   Else
'      NAL1Title = Trim(Me.txtNAL1Title.Text)
'   End If
'   If Trim(Me.txtNAL2Title.Text) = "/" Or UCase(Trim(Me.txtNAL2Title.Text)) = "NA" Or UCase(Trim(Me.txtNAL2Title.Text)) = "N/A" Then
'      NAL2Title = "/"
'   Else
'      NAL2Title = Trim(Me.txtNAL2Title.Text)
'   End If
'
'   If Me.chkUkraine.Value = 1 Then
'        Ukraine = "1"
'        Ukraine_ID = Trim(Me.txtUkraine.Text)
'   Else
'        Ukraine = "0"
'        Ukraine_ID = ""
'   End If
   
'   If chkCE.Value = 1 Then
''      CE = "CE"
'      If chkHPEAddr.Value = 1 Then
'        CE = "3"
'      ElseIf chkCEAddr.Value = 1 Then
'        CE = "2"
'      Else
'        CE = "1"
'      End If
'   ElseIf chkNonCE.Value = 1 Then
'      CE = "0"
'   End If
   If chkCCC.Value = 1 Then
      CCC = "1"
   ElseIf chkNonCCC.Value = 1 Then
      CCC = "0"
   End If
   If chkChinaRoHS.Value = 1 Then
      ChinaRoHS = "1"
   Else
      ChinaRoHS = "0"
   End If

   If chkWEEE.Value = 1 Then
      WEEE = "1"
   ElseIf chkNonWEEE.Value = 1 Then
      WEEE = "0"
   End If
   If chkLaser.Value = 1 Then
      Laser = "1"
   ElseIf chkNonLaser.Value = 1 Then
      Laser = "0"
   End If
   
'   If chkEAC.Value = 1 Then
'      EAC = "1"
'   Else
'      EAC = "0"
'   End If
'   If chkChinaRoHS.Value = 1 Then
'      ChinaRoHS = "1"
'   Else
'      ChinaRoHS = "0"
'   End If
'   If chkTurkeyRohs.Value = 1 Then
'      TurkeyRoHS = "1"
'   Else
'      TurkeyRoHS = "0"
'   End If
'   If chkSVPrint.Value = 1 Then
'      SVPrint = "1"
'   Else
'      SVPrint = "0"
'   End If
'
'   If chkRoHS.Value = 1 Then
'      RoHS = "1"
'   Else
'      RoHS = "0"
'   End If
   
'    If Me.chkATick.Value + Me.chkNonATick.Value = 0 Or Me.chkCTick.Value + Me.chkNonCTick.Value = 0 Or Me.chkICT.Value + Me.chkNonICT.Value = 0 Then
'         MsgBox "ATick,CTick ����ICT û��ѡ����ȷ��!", vbExclamation + vbOKOnly, "ATick,CTick ����ICT û��ѡ��"
'         Exit Sub
'    End If
'
'    If (Me.chkATick.Value = 1 And Len(Trim(Me.txtATick.Text)) = 0) Or (Me.chkCTick.Value = 1 And Len(Trim(Me.txtCTick.Text)) = 0) Or (chkICT.Value = 1 And Len(Trim(Me.txtICT.Text)) = 0) Then
'         MsgBox "ATick,CTick ����ICT �ı���û��ֵ����ȷ��!", vbExclamation + vbOKOnly, "ATick,CTick ����ICT û��ֵ"
'         Exit Sub
'    End If
'
'    If Me.chkRCM.Value + Me.chkNonRCM.Value = 0 Or Me.chkGost.Value + Me.chkNonGost.Value = 0 Or Me.chkKC.Value + Me.chkNonKC.Value = 0 Then
'         MsgBox "RCM,Gost ����KC û��ѡ����ȷ��!", vbExclamation + vbOKOnly, "RCM,Gost ����KC û��ѡ��"
'         Exit Sub
'    End If
'
'    If (Me.chkRCM.Value = 1 And Len(Trim(Me.txtRCM.Text)) = 0) Or (Me.chkGost.Value = 1 And Len(Trim(Me.txtGost.Text)) = 0) Or (Me.chkKC.Value = 1 And Len(Trim(Me.txtKC.Text)) = 0) Then
'         MsgBox "RCM,Gost ����KC �ı���û��ֵ����ȷ��!", vbExclamation + vbOKOnly, "RCM,Gost ����ICT û��ֵ"
'         Exit Sub
'    End If
'
'    If Me.chkATick.Value = 1 Then
'        ATick = "1"
'        If Trim(Me.txtATick.Text) = "/" Or UCase(Trim(Me.txtATick.Text)) = "N/A" Or UCase(Trim(Me.txtATick.Text)) = "NA" Then
'            ATick_ID = ""
'        Else
'            ATick_ID = Trim(Me.txtATick.Text)
'        End If
'    Else
'        ATick = "0"
'        ATick_ID = ""
'    End If
'
'    If Me.chkCTick.Value = 1 Then
'        CTick = "1"
'        If Trim(Me.txtCTick.Text) = "/" Or UCase(Trim(Me.txtCTick.Text)) = "N/A" Or UCase(Trim(Me.txtCTick.Text)) = "NA" Then
'            CTick_ID = ""
'        Else
'            CTick_ID = Trim(Me.txtCTick.Text)
'        End If
'    Else
'        CTick = "0"
'        CTick_ID = ""
'    End If
'
'    If Me.chkICT.Value = 1 Then
'        ICT = "1"
'        If Trim(Me.txtICT.Text) = "/" Or UCase(Trim(Me.txtICT.Text)) = "N/A" Or UCase(Trim(Me.txtICT.Text)) = "NA" Then
'            ICT_ID = ""
'        Else
'            ICT_ID = Trim(Me.txtICT.Text)
'        End If
'    Else
'        ICT = "0"
'        ICT_ID = ""
'    End If
'
'    If Me.chkRCM.Value = 1 Then
'        RCM = "1"
'        If Trim(Me.txtRCM.Text) = "/" Or UCase(Trim(Me.txtRCM.Text)) = "N/A" Or UCase(Trim(Me.txtRCM.Text)) = "NA" Then
'            RCM_ID = ""
'        Else
'            RCM_ID = Trim(Me.txtRCM.Text)
'        End If
'    Else
'        RCM = "0"
'        RCM_ID = ""
'    End If
'
'    If Me.chkGost.Value = 1 Then
'        Gost = "1"
'        If Trim(Me.txtGost.Text) = "/" Or UCase(Trim(Me.txtGost.Text)) = "N/A" Or UCase(Trim(Me.txtGost.Text)) = "NA" Then
'            Gost_ID = ""
'        Else
'            Gost_ID = Trim(Me.txtGost.Text)
'        End If
'    Else
'        Gost = "0"
'        Gost_ID = ""
'    End If
'
'    If Me.chkKC.Value = 1 Then
'        KC = "1"
'        If Trim(Me.txtKC.Text) = "/" Or UCase(Trim(Me.txtKC.Text)) = "N/A" Or UCase(Trim(Me.txtKC.Text)) = "NA" Then
'            KC_ID = ""
'        Else
'            KC_ID = Trim(Me.txtKC.Text)
'        End If
'    Else
'        KC = "0"
'        KC_ID = ""
'    End If
'
'    If Me.chkPWPrint.Value = 1 Then
'        PWPrint = "1"
'    Else
'        PWPrint = "0"
'    End If
'
'    If Me.cb5000.ListIndex <= -1 Then
'        MsgBox "5000״̬û��ѡ��!", vbExclamation + vbOKOnly, "��ѡ��5000״̬��һ��ѡ��"
'         Me.cb5000.SetFocus
'         Exit Sub
'    End If
'
''    Y��N��NA,TBD
'    If Me.cb5000.ListIndex = 0 Then
'        ftStatus = "Y"
'    ElseIf Me.cb5000.ListIndex = 1 Then
'        ftStatus = "N"
'    ElseIf Me.cb5000.ListIndex = 2 Then
'        ftStatus = "NA"
'    ElseIf Me.cb5000.ListIndex = 3 Then
'        ftStatus = "TBD"
'    End If
    
'    FTStatus = CStr(cb5000.ListIndex)
    

  
'  txtGW.Text = LCase(Trim(txtGW.Text))
   
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from tblConsenNew where Part_Number ='" & Trim(txtSN.Text) & "' "
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "��Ʒ�����Ѵ���!", vbExclamation + vbOKOnly, "��Ʒ����ظ�"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close

      sql = "insert [tblConsenNew]([Part_Number],[Part_Type],[Part_Model],[Power],[Material],[SalesOrder],[ChinaRoHS],[WEEE],[CCC],[Laser]) " & _
            "Values('" & Trim(txtSN.Text) & "','" & Trim(txtType.Text) & "','" & Trim(txtModel.Text) & "','" & Trim(txtPower.Text) & "','" & Trim(txtMaterial.Text) & "','" & Trim(txtOrder.Text) & "'," & ChinaRoHS & "," & WEEE & "," & CCC & "," & Laser & ")"
             
      sql = sql & " insert tblConsenNew_log(CREATE_USER,[Part_Number],[Part_Type],[Part_Model],[Power],[Material],[SalesOrder],[ChinaRoHS],[WEEE],[CCC],[Laser],COMMENT) " & _
            "Values('" & golUSERNAME & "','" & Trim(txtSN.Text) & "','" & Trim(txtType.Text) & "','" & Trim(txtModel.Text) & "','" & Trim(txtPower.Text) & "','" & Trim(txtMaterial.Text) & "','" & Trim(txtOrder.Text) & "'," & ChinaRoHS & "," & WEEE & "," & CCC & "," & Laser & ",'Insert')"
             
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "����NEC�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "����ʧ��"
      Else
        MsgBox "����NEC�趨���ϳɹ�!", vbInformation + vbOKOnly, "�����ɹ�"
      End If
      renovate ("")
      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update tblConsenNew set Part_Type='" & Trim(txtType.Text) & "',Part_Model='" & Trim(txtModel.Text) & "',Power ='" & Trim(txtPower.Text) & "',Material='" & Trim(txtMaterial.Text) & "',SalesOrder='" & Trim(txtOrder.Text) & "',ChinaRoHS=" & ChinaRoHS & ",WEEE=" & WEEE & ",CCC=" & CCC & ",Laser=" & Laser & _
            " where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and Part_Number ='" & Trim(txtSN.Text) & "'"
      
      
      sql = sql & " insert tblConsenNew_log(CREATE_USER,[Part_Number],[Part_Type],[Part_Model],[Power],[Material],[SalesOrder],[ChinaRoHS],[WEEE],[CCC],[Laser],COMMENT) " & _
            "Values('" & golUSERNAME & "','" & Trim(txtSN.Text) & "','" & Trim(txtType.Text) & "','" & Trim(txtModel.Text) & "','" & Trim(txtPower.Text) & "','" & Trim(txtMaterial.Text) & "','" & Trim(txtOrder.Text) & "'," & ChinaRoHS & "," & WEEE & "," & CCC & "," & Laser & ",'Update')"
       
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "�޸�NEC�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "�޸�ʧ��"
      Else
         MsgBox "�޸�NEC�趨���ϳɹ�!", vbInformation + vbOKOnly, "�޸ĳɹ�"
      End If
      renovate ("")
      cmdCancel_Click
   End If
   renovate ("")
End Sub

Private Sub cmdDelete_Click()
   If mfgH3C.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫɾ������!", vbInformation + vbOKOnly, "δѡ����"
      Exit Sub
   End If
   
   sql = " insert tblConsenNew_log(CREATE_USER,[Part_Number],[Part_Type],[Part_Model],[Power],[Material],[SalesOrder],[ChinaRoHS],[WEEE],[CCC],[Laser],COMMENT) " & _
             " select '" & golUSERNAME & "',[Part_Number],[Part_Type],[Part_Model],[Power],[Material],[SalesOrder],[ChinaRoHS],[WEEE],[CCC],[Laser],'Delete' from tblConsenNew " & _
             " where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and Part_Number ='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 2) & "'"

   sql = sql & " delete from tblConsenNew where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and Part_Number ='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 2) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "ɾ�������趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "ɾ��ʧ��"
   End If
      MsgBox "ɾ�������趨���ϳɹ�!", vbInformation + vbOKOnly, "ɾ���ɹ�"
   renovate ("")
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If mfgH3C.Rows = 0 Then
      MsgBox "�����Ͽɻ��", vbExclamation + vbOKOnly, "������"
      Exit Sub
   End If
   If txtPath.Text <> "" Then
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets.Item(1)
       For i = 0 To mfgH3C.Rows - 1
         For j = 1 To mfgH3C.Cols - 1
          xlSheet.Cells(i + 1, j) = mfgH3C.TextMatrix(i, j)
       Next j
      Next i
      xlBook.SaveAs (txtPath.Text)
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "�����EXCEL���ϳɹ�!!", vbInformation + vbOKOnly, "����ɹ�"
    End If
End Sub

Private Sub cmdImport_Click()
   If txtPath.Text = "" Then
      MsgBox "����·������Ϊ��!", vbExclamation + vbOKOnly, "����·����"
      Exit Sub
   End If
   Dim action As Integer
   Dim info As Boolean
   info = True
   Set xlBook = xlApp.Workbooks.Open(txtPath.Text)
      For i = 1 To xlBook.Sheets.Count
       Set xlSheet = xlBook.Sheets.Item(i)
       For j = 2 To xlSheet.Rows.Count
        r = xlSheet.Cells(j, 1)
        If r = "" Then
           Exit For
        Else
          Dim cellValue As String
          Dim cellhvValue As String
          
          Dim isexist As Boolean
          If xlSheet.Cells(j, 19) = "" Then
             MsgBox "�������ϸ�ʽ����ȷ!", vbExclamation + vbOKOnly, "��ʽ����"
             Exit Sub
          End If
          If Not ((xlSheet.Cells(j, 18) = "N") Or (xlSheet.Cells(j, 18) = "Y")) Then
             MsgBox "�������ϸ�ʽ����ȷ!", vbExclamation + vbOKOnly, "��ʽ����"
             Exit Sub
          End If
          isexist = False
          For K = 1 To 19
          '======================================================
           If K = 3 Then
             cellValue = xlSheet.Cells(j, K)
             cellhvValue = xlSheet.Cells(j, 2)
             
             If cellValue = "" Or cellhvValue = "" Then
                MsgBox "�������ϸ�ʽ����ȷ!", vbExclamation + vbOKOnly, "��ʽ����"
                Exit Sub
             End If
             
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from tblH3C where SN='" & cellValue & "' and HV='" & cellhvValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("��Ʒ����&�汾�Ѵ���!", vbAbortRetryIgnore + vbExclamation, "�����ظ�")
                End If
                
                If action = vbAbort Then
                   MsgBox "���ϵ�������ֹ!!", vbInformation + vbOKOnly, "������ֹ"
                   rcd.Close
                   Exit Sub
                ElseIf action = vbIgnore And info = True Then
                   MsgBox "�ظ���Ʒ������ϲ��ᵼ��,���Ե�..!!", vbInformation + vbOKOnly, "�ظ����ᵼ��"
                   rcd.Close
                   info = False
                   Exit For
                ElseIf action = vbRetry And info = True Then
                   MsgBox "�ظ���Ʒ������ϻ��Զ�����,���Ե�..!!", vbInformation + vbOKOnly, "�ظ����Զ�����"
                   info = False
                End If
                isexist = True
             Else
                isexist = False
             End If
             rcd.Close
            End If
            '==================================================
            
            If K = 19 Then
               If action = vbRetry Then
                   sql = "Update tblH3C set CPN='" & xlSheet.Cells(j, 4) & "',EPN='" & xlSheet.Cells(j, 5) & "',Des='" & xlSheet.Cells(j, 6) & "',OS='" & xlSheet.Cells(j, 7) & "',GW='" & xlSheet.Cells(j, 8) & "',CE='" & xlSheet.Cells(j, 9) & "',WEEE='" & xlSheet.Cells(j, 10) & "',ChinaRoHS='" & xlSheet.Cells(j, 11) & "'," & _
                        "RoHS='" & xlSheet.Cells(j, 12) & "',TurkeyRohs='" & xlSheet.Cells(j, 13) & "',MS='" & xlSheet.Cells(j, 14) & "',MSValidFrom='" & xlSheet.Cells(j, 15) & "',NAL='" & xlSheet.Cells(j, 16) & "',ValidFrom='" & xlSheet.Cells(j, 17) & "',PrintSV='" & xlSheet.Cells(j, 18) & "',Remark='" & xlSheet.Cells(j, 19) & "'" & _
                        " where SN='" & xlSheet.Cells(j, 3) & "' and HV='" & xlSheet.Cells(j, 2) & "' "
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "�޸�H3C�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "�޸�ʧ��"
                   End If
'                   MsgBox "�޸�H3C�趨���ϳɹ�!"
               ElseIf isexist = False Then
                   sql = "Insert into tblH3C(ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark) " & _
                        " Values(" & getmaxID("tblH3C") & ",'" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "','" & xlSheet.Cells(j, 4) & "','" & xlSheet.Cells(j, 5) & "','" & xlSheet.Cells(j, 6) & "','" & xlSheet.Cells(j, 7) & "','" & xlSheet.Cells(j, 8) & "','" & xlSheet.Cells(j, 9) & "','" & xlSheet.Cells(j, 10) & "','" & xlSheet.Cells(j, 11) & "'," & _
                        "'" & xlSheet.Cells(j, 12) & "','" & xlSheet.Cells(j, 13) & "','" & xlSheet.Cells(j, 14) & "','" & xlSheet.Cells(j, 15) & "','" & xlSheet.Cells(j, 16) & "','" & xlSheet.Cells(j, 17) & "','" & xlSheet.Cells(j, 18) & "','" & xlSheet.Cells(j, 19) & "')"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "����H3C�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "�޸�ʧ��"
                   End If
'                   MsgBox "����H3C�趨���ϳɹ�!"
               End If
           End If
         Next K
         
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "H3C�趨���ϵ���ɹ�!"
      renovate ("")
End Sub

Private Sub cmdInsert_Click()
    op = "Insert"
    Reset

'    enable
'    txtSN.Text = ""
'    txtCPN.Text = ""
'    txtEPN.Text = ""
'    txtDes.Text = ""
'    txtSize.Text = ""
'    txtGW.Text = ""
'
'    chkCE.Value = 1
'    chkWEEE.Value = 1
'    chkChinaRoHS.Value = 1
'    chkRoHS.Value = 1
'    chkTurkeyRohs.Value = 1
'    chkSVPrint.Value = 1
'
'    txtNAL2.Text = "/"
'    txtNAL2Title = "/"
'    txtMS.Text = "N/A"
'
'
'    txtHV.Text = "N/A"
'    txtRemark.Text = "N/A"

End Sub

Private Sub cmdQuery_Click()
    If txtSN.Enabled = False Then
      MsgBox "�밴������ť��վͿ������ѯ����!", vbOKOnly + vbInformation, "�����ѯ����"
    End If
    If rec.State = 1 Then
        rec.Close
     End If
       sql = " SELECT [ID],[Part_Number],[Part_Type],[Part_Model],[Power],[Material],[SalesOrder] " & _
        " ,case when ChinaRoHS is null then 'N/A' when ChinaRoHS = 0 then 'No' when ChinaRoHS = 1 then 'Yes' end as 'ChinaRoHS'" & _
        " ,case when WEEE is null then 'N/A' when WEEE = 0 then 'No' when WEEE = 1 then 'Yes' end as 'WEEE' " & _
        " ,case when CCC is null then 'N/A' when CCC = 0 then 'No' when CCC = 1 then 'Yes' end as 'CCC' " & _
        " ,case when Laser is null then 'N/A' when Laser = 0 then 'No' when Laser = 1 then 'Yes' end as 'Laser' " & _
        " FROM [Print].[dbo].[tblConsenNew] where 1 = 1"
     
     If txtSN.Text <> "" Then
        sql = sql & " and Part_Number like '%" & txtSN.Text & "%'"
     End If
'     If txtCPN.Text <> "" Then
'        sql = sql & " and CPN like '%" & txtCPN.Text & "%'"
'     End If
'     If txtEPN.Text <> "" Then
'        sql = sql & " and EPN='%" & txtEPN.Text & "%'"
'     End If
'     If Me.txtProductID.Text <> "" Then
'        sql = sql & " and ProductID like '%" & Me.txtProductID.Text & "%'"
'     End If
'     If txtSize.Text <> "" Then
'        sql = sql & " and Size like '%" & txtSize.Text & "%'"
'     End If
'     If txtGW.Text <> "" Then
'        sql = sql & " and GW like '%" & txtGW.Text & "%'"
'     End If
     sql = sql & " order by Part_Number "
    renovate (sql)
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

'Private Sub cmdSelect_Click()
'   On Error Resume Next
''   cdSelect.CancelError = True
'   cdSelect.Filter = "*.xls|*.xls"
'   cdSelect.action = 1
'   If cdSelect.FileName <> "" Then txtPath.Text = cdSelect.FileName
'End Sub

Private Sub cmdUpdate_Click()
   If mfgH3C.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫ�޸ĵ���!", vbInformation + vbOKOnly, "δѡ����"
      Exit Sub
   End If
   mfgH3C_Click
   op = "Update"
   Reset
End Sub

Private Sub Form_Load()
   unable
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   
   cb5000.AddItem ("����")
   cb5000.AddItem ("������")
   cb5000.AddItem ("���漰")
   cb5000.AddItem ("����")
   
   renovate ("")
End Sub

Private Sub renovate(sql As String)
    Set mfgH3C.DataSource = Nothing
    If sql = "" Then
       sql = " SELECT [ID],[Part_Number],[Part_Type],[Part_Model],[Power],[Material],[SalesOrder] " & _
        " ,case when ChinaRoHS is null then 'N/A' when ChinaRoHS = 0 then 'No' when ChinaRoHS = 1 then 'Yes' end as 'ChinaRoHS'" & _
        " ,case when WEEE is null then 'N/A' when WEEE = 0 then 'No' when WEEE = 1 then 'Yes' end as 'WEEE' " & _
        " ,case when CCC is null then 'N/A' when CCC = 0 then 'No' when CCC = 1 then 'Yes' end as 'CCC' " & _
        " ,case when Laser is null then 'N/A' when Laser = 0 then 'No' when Laser = 1 then 'Yes' end as 'Laser' " & _
        " FROM [Print].[dbo].[tblConsenNew] where 1 = 1"
    End If
    If rec.State = 1 Then
    rec.Close
    End If
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    Set mfgH3C.DataSource = rec
    With mfgH3C
      .Cols = rec.Fields.Count + 1
      .ColWidth(0) = 400
      .ColWidth(1) = 2000
      .ColWidth(2) = 2000
      .ColWidth(3) = 2000
      .ColWidth(4) = 2000
      .ColWidth(5) = 2000
      .ColWidth(6) = 2000
      .ColWidth(7) = 2000
      .ColWidth(8) = 2000
      .ColWidth(9) = 2000
      .ColWidth(10) = 2000
'      .ColWidth(11) = 1000
'      .ColWidth(12) = 1000
'      .ColWidth(13) = 1000
'      .ColWidth(14) = 1000
'      .ColWidth(15) = 1000
'      .ColWidth(16) = 1000
'      .ColWidth(17) = 1000
'      .ColWidth(18) = 1000
'      .ColWidth(19) = 1000
'      .ColWidth(20) = 500
'      .ColWidth(21) = 1000
'      .ColWidth(22) = 1000
'      .ColWidth(23) = 1000
'      .ColWidth(24) = 1000
'      .ColWidth(25) = 1000
'      .ColWidth(26) = 1000
'      .ColWidth(27) = 1000
'      .ColWidth(28) = 1000
'      .ColWidth(29) = 1000
'      .ColWidth(30) = 1000
'      .ColWidth(31) = 1000
'      .ColWidth(32) = 1000
'      .ColWidth(33) = 1000
'      .ColWidth(34) = 1000
'      .ColWidth(35) = 1000
'      .ColWidth(36) = 1000
'      .ColWidth(37) = 800
'      .ColWidth(38) = 1000
'      ,[Part_Number]
'      ,[Part_Revision]
'      ,[Model]
'      ,[CPN]
'      ,[EPN]
'      ,[Des]
'      ,[Size]
'      ,[GW]
'      ,[MS]
'      ,[NAL1]
'      ,[NAL1_Title]
'      ,[NAL2]
'      ,[NAL2_Title]
'      ,[CE]
'      ,[WEEE]
'      ,[ChinaRoHS]
'      ,[RoHS]
'      ,[TurkeyRoHS]
'      ,[Ukraine]
      .TextMatrix(0, 1) = "ID"
      .TextMatrix(0, 2) = "��Ʒ����"
      .TextMatrix(0, 3) = "��Ʒ���"
      .TextMatrix(0, 4) = "��Ʒ�ͺ�"
      .TextMatrix(0, 5) = "��Դά��"
      .TextMatrix(0, 6) = "���ϴ���"
      .TextMatrix(0, 7) = "���۶���"
      .TextMatrix(0, 8) = "ChinaRoHS"
      .TextMatrix(0, 9) = "WEEE"
      .TextMatrix(0, 10) = "CCC"
      .TextMatrix(0, 11) = "Laser"
'      .TextMatrix(0, 10) = "�����׼"
'      .TextMatrix(0, 11) = "����1"
'      .TextMatrix(0, 12) = "��������1"
'      .TextMatrix(0, 13) = "����2"
'      .TextMatrix(0, 14) = "��������2"
'      .TextMatrix(0, 15) = "CE"
'      .TextMatrix(0, 16) = "WEEE"
'      .TextMatrix(0, 17) = "ChinaRoHS"
'      .TextMatrix(0, 18) = "H3C RoHS"
'      .TextMatrix(0, 19) = "TurkeyRoHS"
'      .TextMatrix(0, 20) = "Ukraine"
'      .TextMatrix(0, 21) = "Ukraine ID"
'      ,[ATick]
'      ,[ATick_ID]
'      ,[CTick]
'      ,[CTick_ID]
'      ,[ICT]
'      ,[ICT_ID]
'      ,[RCM]
'      ,[RCM_ID]
'      ,[Gost]
'      ,[Gost_ID]
'      ,[KC]
'      ,[KC_ID]
'      ,[Print_SV]
'      ,[Print_Power]
'      ,[Remark]

'      .TextMatrix(0, 22) = "ATick"
'      .TextMatrix(0, 23) = "ATick ID"
'      .TextMatrix(0, 24) = "CTick"
'      .TextMatrix(0, 25) = "CTick ID"
'      .TextMatrix(0, 26) = "ICT"
'      .TextMatrix(0, 27) = "ICT ID"
'      .TextMatrix(0, 28) = "RCM"
'      .TextMatrix(0, 29) = "RCM ID"
'      .TextMatrix(0, 30) = "Gost"
'      .TextMatrix(0, 31) = "Gost ID"
'      .TextMatrix(0, 32) = "KC"
'      .TextMatrix(0, 33) = "KC ID"
'      .TextMatrix(0, 34) = "��ӡ�汾"
'      .TextMatrix(0, 35) = "��ӡ��Դ"
'      .TextMatrix(0, 36) = "5000״̬"
'      .TextMatrix(0, 37) = "��ע"
'      .TextMatrix(0, 38) = "EAC"
    End With
    rec.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rec.State = 1 Then
      rec.Close
      Set conn = Nothing
   End If
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub

Private Sub mfgH3C_Click()
'    [tblConsenNew](ID,[Part_Number],[Part_Revision],[ProductID],[CPN],[EPN],[Des],[Size],[GW],[MS],
'10
'    [NAL1] , [NAL1_Title], [NAL2], [NAL2_Title], [CE], [WEEE], [ChinaRoHS], [RoHS], [TurkeyRoHS], [Ukraine],

   If mfgH3C.RowSel > 0 Then
'      txtHV.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 3)
      txtSN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 2)
      txtType.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 3)
      txtModel.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 4)
      txtPower.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 5)
      txtMaterial.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 6)
      txtOrder.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 7)
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 8))) = "YES" Then
        chkChinaRoHS.Value = 1
      ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 8)) = "NO" Then
        chkNonChinaRoHS.Value = 1
      End If
     
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 9))) = "YES" Then
        chkWEEE.Value = 1
      ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 9)) = "NO" Then
        chkNonWEEE.Value = 1
      End If
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 10))) = "YES" Then
        chkCCC.Value = 1
      ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 10)) = "NO" Then
        chkNonCCC.Value = 1
      End If
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 11))) = "YES" Then
        chkLaser.Value = 1
      ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 11)) = "NO" Then
        chkNonLaser.Value = 1
      End If
'      txtProductID.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 4)
'      txtCPN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 5)
'      txtEPN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 6)
'      txtSize.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 8)
'      txtGW.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 9)
'      txtMS.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 10)
      
'[NAL1] , [NAL1_Title], [NAL2], [NAL2_Title], [CE], [WEEE], [ChinaRoHS], [RoHS], [TurkeyRoHS], [Ukraine]
'11-20
'    Me.txtNAL1.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 11))
'    Me.txtNAL1Title.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 12))
'    Me.txtNAL2.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 13))
'    Me.txtNAL2Title.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 14))
'
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 15))) = "CE" Then
'     chkCE.Value = 1
'     chkNonCE.Value = 0
'     chkCEAddr.Value = 0
'     chkHPEAddr.Value = 0
'    ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 15) = "Non CE" Then
'     chkCE.Value = 0
'     chkNonCE.Value = 1
'     chkCEAddr.Value = 0
'     chkHPEAddr.Value = 0
'    ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 15) = "CE+CE Addr" Then
'     chkCE.Value = 1
'     chkNonCE.Value = 0
'     chkCEAddr.Value = 1
'     chkHPEAddr.Value = 0
'    ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 15) = "CE+HPE Addr" Then
'     chkCE.Value = 1
'     chkNonCE.Value = 0
'     chkCEAddr.Value = 0
'     chkHPEAddr.Value = 1
'    End If
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 16))) = "YES" Then
'     chkWEEE.Value = 1
''     //chkNonWEEE.Value = 0
'    Else
''     chkWEEE.Value = 0
'     chkNonWEEE.Value = 1
'    End If

'
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 18))) = "YES" Then
'    chkRoHS.Value = 1
''    chkNonRoHS.Value = 0
'    ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 18)) = "NO" Then
''     chkRoHS.Value = 0
'     chkNonRoHS.Value = 1
'    End If
'
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 19))) = "YES" Then
'     chkTurkeyRohs.Value = 1
''     chkNonTurkeyRohs.Value = 0
'    ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 19)) = "NO" Then
''     chkTurkeyRohs.Value = 0
'     chkNonTurkeyRohs.Value = 1
'    End If
'
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 20))) = "YES" Then
'        Me.chkUkraine.Value = 1
'        Me.txtUkraine.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 21)
'    ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 20)) = "NO" Then
'        Me.chkNonUkraine.Value = 1
'        Me.txtUkraine.Text = ""
'    End If
    

      
    '[Ukraine_ID],[ATick],[ATick_ID],[CTick],[CTick_ID],[ICT],[ICT_ID],[RCM],[RCM_ID],[Gost],[Gost_ID],[KC],[KC_ID],[Print_SV],[Print_Power],[Remark])
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 22))) = "YES" Then
'        Me.chkATick.Value = 1
'        Me.txtATick.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 23))
'    ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 22))) = "NO" Then
'        Me.chkNonATick.Value = 1
'        Me.txtATick.Text = ""
'    End If
'
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 24))) = "YES" Then
'        Me.chkCTick.Value = 1
'        Me.txtCTick.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 25))
'    ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 24))) = "NO" Then
'        Me.chkNonCTick.Value = 1
'        Me.txtCTick.Text = ""
'    End If
'
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 26))) = "YES" Then
'        Me.chkICT.Value = 1
'        Me.txtICT.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 27))
'    ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 26))) = "NO" Then
'        Me.chkNonICT.Value = 1
'        Me.txtICT.Text = ""
'    End If
'
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 28))) = "YES" Then
'        Me.chkRCM.Value = 1
'        Me.txtRCM.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 29))
'    ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 28))) = "NO" Then
'        Me.chkNonRCM.Value = 1
'        Me.txtRCM.Text = ""
'    End If
    
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 30))) = "YES" Then
'        Me.chkGost.Value = 1
'        Me.txtGost.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 31))
'    ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 30))) = "NO" Then
'        Me.chkNonGost.Value = 1
'        Me.txtGost.Text = ""
'    End If
'
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 32))) = "YES" Then
'        Me.chkKC.Value = 1
'        Me.txtKC.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 33))
'    ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 32))) = "NO" Then
'        Me.chkNonKC.Value = 1
'        Me.txtKC.Text = ""
'    End If
'
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 34))) = "YES" Then
'        Me.chkSVPrint.Value = 1
'    ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 34))) = "NO" Then
'        Me.chkNonSVPrint.Value = 1
'    End If
'
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 35))) = "YES" Then
'        Me.chkPWPrint.Value = 1
'    ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 35))) = "NO" Then
'        Me.chkNonPWPrint.Value = 1
'    End If
    
'    If Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 36)) = "Y" Then
'        cb5000.ListIndex = 0
'    ElseIf Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 36)) = "N" Then
'        cb5000.ListIndex = 1
'    ElseIf Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 36)) = "NA" Then
'        cb5000.ListIndex = 2
'    ElseIf Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 36)) = "TBD" Then
'        cb5000.ListIndex = 3
'    End If
    
'    txtRemark.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 37)
'    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 38))) = "YES" Then
'        chkEAC.Value = 1
'    ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 38)) = "NO" Then
'        chkNonEAC.Value = 1
'    End If
   End If
End Sub

Private Sub mfgH3C_SelChange()
   mfgH3C_Click
End Sub


