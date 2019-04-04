VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHUVSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HUV Setting(Test)"
   ClientHeight    =   12705
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   18915
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHUVSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12705
   ScaleWidth      =   18915
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "四合一"
      Height          =   1695
      Left            =   120
      TabIndex        =   80
      Top             =   5760
      Width           =   18735
      Begin VB.TextBox txtProduct 
         Height          =   495
         Left            =   15360
         TabIndex        =   100
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtClass4 
         Height          =   450
         Left            =   10800
         TabIndex        =   95
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtModel4 
         Height          =   450
         Left            =   10800
         TabIndex        =   94
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtRatedValue4 
         Height          =   450
         Left            =   15360
         TabIndex        =   93
         Top             =   360
         Width           =   3135
      End
      Begin VB.CheckBox chkCE4 
         Caption         =   "CE"
         Height          =   375
         Left            =   2400
         TabIndex        =   88
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkNonCE4 
         Caption         =   "无 CE"
         Height          =   375
         Left            =   3480
         TabIndex        =   87
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkWEEE4 
         Caption         =   "有"
         Height          =   375
         Left            =   2400
         TabIndex        =   86
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkNonWEEE4 
         Caption         =   "无"
         Height          =   375
         Left            =   3480
         TabIndex        =   85
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkChinaRoHS4 
         Caption         =   "有"
         Height          =   375
         Left            =   7440
         TabIndex        =   84
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkNonChinaRoHS4 
         Caption         =   "无"
         Height          =   375
         Left            =   8400
         TabIndex        =   83
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkNonCN3C4 
         Caption         =   "无"
         Height          =   375
         Left            =   8400
         TabIndex        =   82
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox chkCN3C4 
         Caption         =   "有"
         Height          =   375
         Left            =   7440
         TabIndex        =   81
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "产品型号"
         Height          =   375
         Left            =   13920
         TabIndex        =   99
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "工艺分类"
         Height          =   375
         Left            =   9480
         TabIndex        =   98
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "认证型号"
         Height          =   375
         Left            =   9480
         TabIndex        =   97
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "电源额定值"
         Height          =   375
         Left            =   13800
         TabIndex        =   96
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label lblCE4 
         Caption         =   "认证信息CE:"
         Height          =   375
         Left            =   120
         TabIndex        =   92
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblWEEE4 
         Caption         =   "认证信息WEEE:"
         Height          =   375
         Left            =   120
         TabIndex        =   91
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblChinaRoHS4 
         Caption         =   "China RoHS:"
         Height          =   375
         Left            =   5400
         TabIndex        =   90
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblCCC4 
         Caption         =   "CCC认证"
         Height          =   375
         Left            =   5400
         TabIndex        =   89
         Top             =   960
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdSelect 
      Left            =   2520
      Top             =   11040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   735
      Left            =   14280
      TabIndex        =   40
      Top             =   11880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   735
      Left            =   12840
      TabIndex        =   39
      Top             =   11880
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "确定(Confirm)"
      Height          =   735
      Left            =   11160
      TabIndex        =   38
      Top             =   11880
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(Delete)"
      Height          =   735
      Left            =   14280
      TabIndex        =   37
      Top             =   11040
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修改(Update)"
      Height          =   735
      Left            =   12840
      TabIndex        =   36
      Top             =   11040
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "新增(Insert)"
      Height          =   735
      Left            =   11160
      TabIndex        =   35
      Top             =   11040
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查询(Query)"
      Height          =   735
      Left            =   9480
      TabIndex        =   34
      Top             =   11520
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "导出(Export)"
      Height          =   735
      Left            =   3240
      TabIndex        =   33
      Top             =   11880
      Width           =   1455
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "导入(Import)"
      Height          =   735
      Left            =   3240
      TabIndex        =   32
      Top             =   11040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择(Select)"
      Height          =   495
      Left            =   1200
      TabIndex        =   31
      Top             =   12120
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   11520
      Width           =   3015
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgH3C 
      Height          =   3255
      Left            =   0
      TabIndex        =   28
      Top             =   7560
      Width           =   18645
      _ExtentX        =   32888
      _ExtentY        =   5741
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fmH3C 
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   18800
      Begin VB.TextBox txtCertification 
         Height          =   450
         Left            =   15440
         TabIndex        =   79
         Top             =   4020
         Width           =   2895
      End
      Begin VB.TextBox txtRatedValue 
         Height          =   450
         Left            =   15440
         TabIndex        =   78
         Top             =   3520
         Width           =   2895
      End
      Begin VB.TextBox txtModel 
         Height          =   450
         Left            =   15400
         TabIndex        =   77
         Top             =   3020
         Width           =   2895
      End
      Begin VB.TextBox txtClass 
         Height          =   450
         Left            =   15400
         TabIndex        =   76
         Top             =   2520
         Width           =   2895
      End
      Begin VB.ComboBox cbENCA 
         Height          =   450
         Left            =   11160
         TabIndex        =   71
         Top             =   2040
         Width           =   5535
      End
      Begin VB.ComboBox cbENCN 
         Height          =   450
         Left            =   2280
         TabIndex        =   69
         Top             =   2040
         Width           =   5535
      End
      Begin VB.ComboBox cbCNCA 
         Height          =   450
         Left            =   11160
         TabIndex        =   67
         Top             =   1440
         Width           =   5535
      End
      Begin VB.ComboBox cbCNCN 
         Height          =   450
         Left            =   2280
         TabIndex        =   65
         Top             =   1440
         Width           =   5535
      End
      Begin VB.CheckBox chkFCC 
         Caption         =   "有"
         Height          =   375
         Left            =   11520
         TabIndex        =   63
         Top             =   3840
         Width           =   615
      End
      Begin VB.CheckBox chkNonFCC 
         Caption         =   "无"
         Height          =   375
         Left            =   12360
         TabIndex        =   62
         Top             =   3840
         Width           =   735
      End
      Begin VB.ComboBox cbSalesLocations 
         DataSource      =   "中英文;纯英文;讯美"
         Height          =   450
         Left            =   11520
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CheckBox chkCN3C 
         Caption         =   "有"
         Height          =   375
         Left            =   11520
         TabIndex        =   58
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkNonCN3C 
         Caption         =   "无"
         Height          =   375
         Left            =   12360
         TabIndex        =   57
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkNonRoHS 
         Caption         =   "无"
         Height          =   495
         Left            =   3360
         TabIndex        =   55
         Top             =   3840
         Width           =   855
      End
      Begin VB.CheckBox chkRoHS 
         Caption         =   "有"
         Height          =   375
         Left            =   2400
         TabIndex        =   54
         Top             =   3840
         Width           =   975
      End
      Begin VB.CheckBox chkNoPrintSV 
         Caption         =   "否"
         Height          =   375
         Left            =   9000
         TabIndex        =   53
         Top             =   3840
         Width           =   735
      End
      Begin VB.CheckBox chkSVPrint 
         Caption         =   "是"
         Height          =   375
         Left            =   8160
         TabIndex        =   52
         Top             =   3840
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpMSValidFrom 
         Height          =   495
         Left            =   8880
         TabIndex        =   50
         Top             =   5040
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         Format          =   109772801
         CurrentDate     =   40425
      End
      Begin VB.CheckBox chkNonTurkeyRohs 
         Caption         =   "无"
         Height          =   375
         Left            =   9000
         TabIndex        =   48
         Top             =   3240
         Width           =   735
      End
      Begin VB.CheckBox chkTurkeyRohs 
         Caption         =   "有"
         Height          =   375
         Left            =   8160
         TabIndex        =   47
         Top             =   3240
         Width           =   735
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         Caption         =   "无"
         Height          =   375
         Left            =   9000
         TabIndex        =   44
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkChinaRoHS 
         Caption         =   "有"
         Height          =   375
         Left            =   8160
         TabIndex        =   43
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtRemark 
         Height          =   495
         Left            =   13920
         TabIndex        =   27
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox txtHV 
         Height          =   495
         Left            =   13920
         TabIndex        =   25
         Top             =   4500
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpValidFrom 
         Height          =   495
         Left            =   2280
         TabIndex        =   23
         Top             =   5040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   109772801
         CurrentDate     =   39757
      End
      Begin VB.TextBox txtMS 
         Height          =   450
         Left            =   8880
         TabIndex        =   22
         Top             =   4440
         Width           =   2775
      End
      Begin VB.CheckBox chkNonWEEE 
         Caption         =   "无"
         Height          =   375
         Left            =   3360
         TabIndex        =   19
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chkWEEE 
         Caption         =   "有"
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CheckBox chkNonCE 
         Caption         =   "无 CE"
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CheckBox chkCE 
         Caption         =   "CE"
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtNAL 
         Height          =   435
         Left            =   2280
         TabIndex        =   12
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox txtGW 
         Height          =   450
         Left            =   13560
         TabIndex        =   11
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtOS 
         Height          =   450
         Left            =   8160
         TabIndex        =   9
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtDes 
         Height          =   450
         Left            =   2280
         TabIndex        =   8
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtEPN 
         Height          =   450
         Left            =   13560
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtCPN 
         Height          =   450
         Left            =   8160
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtSN 
         Height          =   450
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label11 
         Caption         =   "产品通过认证信息"
         Height          =   375
         Left            =   13050
         TabIndex        =   75
         Top             =   4100
         Width           =   2400
      End
      Begin VB.Label Label10 
         Caption         =   "电源额定值"
         Height          =   375
         Left            =   13440
         TabIndex        =   74
         Top             =   3600
         Width           =   1600
      End
      Begin VB.Label Label9 
         Caption         =   "认证型号"
         Height          =   375
         Left            =   13440
         TabIndex        =   73
         Top             =   3100
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "工艺分类"
         Height          =   375
         Left            =   13440
         TabIndex        =   72
         Top             =   2600
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "英文地址:"
         Height          =   375
         Left            =   8760
         TabIndex        =   70
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "英文公司名:"
         Height          =   375
         Left            =   120
         TabIndex        =   68
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "中文地址:"
         Height          =   375
         Left            =   8760
         TabIndex        =   66
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "中文公司名:"
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "FCC:"
         Height          =   375
         Left            =   10560
         TabIndex        =   61
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "标签类型:"
         Height          =   375
         Left            =   9960
         TabIndex        =   59
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "CCC认证"
         Height          =   375
         Left            =   9960
         TabIndex        =   56
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblPrintSV 
         Caption         =   "是否打印软件版本:"
         Height          =   375
         Left            =   5520
         TabIndex        =   51
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label lblMSValidFrom 
         Caption         =   "制造标准有效期:"
         Height          =   495
         Left            =   5520
         TabIndex        =   49
         Top             =   5040
         Width           =   2775
      End
      Begin VB.Label lblChinaRoHS 
         Caption         =   "China RoHS:"
         Height          =   375
         Left            =   5760
         TabIndex        =   46
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lbTurkeyRohs 
         Caption         =   "Turkey RoHS:"
         Height          =   495
         Left            =   5760
         TabIndex        =   45
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label lblOS 
         Caption         =   "外尺寸(MM):"
         Height          =   375
         Left            =   6120
         TabIndex        =   42
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblNAL 
         Caption         =   "进网许可号:"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label lblRemark 
         Caption         =   "备注:"
         Height          =   495
         Left            =   12360
         TabIndex        =   26
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label lblHV 
         Caption         =   "硬件版本:"
         Height          =   495
         Left            =   11760
         TabIndex        =   24
         Top             =   4500
         Width           =   2055
      End
      Begin VB.Label lblMS 
         Caption         =   "制造标准:"
         Height          =   375
         Left            =   5520
         TabIndex        =   21
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label lblRoHS 
         Caption         =   "认证信息RoHS:"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label lblWEEE 
         Caption         =   "认证信息WEEE:"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblCE 
         Caption         =   "认证信息CE:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblValidFrom 
         Caption         =   "进网有效期:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   5040
         Width           =   2295
      End
      Begin VB.Label lblGW 
         Caption         =   "毛重(kg):"
         Height          =   375
         Left            =   12120
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblDes 
         Caption         =   "产品描述:"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         Caption         =   "产品名称(英文):"
         Height          =   375
         Left            =   11400
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblCPN 
         Caption         =   "产品名称(中文):"
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblSN 
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Label lblPath 
      Caption         =   "导入/导出路径:"
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   11040
      Width           =   2175
   End
End
Attribute VB_Name = "frmHUVSetting"
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

Private Sub enable()
   txtSN.Enabled = True
   txtSN.BackColor = &HFFFFFF
   txtCPN.Enabled = True
   txtCPN.BackColor = &HFFFFFF
   txtEPN.Enabled = True
   txtEPN.BackColor = &HFFFFFF
   txtDes.Enabled = True
   txtDes.BackColor = &HFFFFFF
   txtOS.Enabled = True
   txtOS.BackColor = &HFFFFFF
   txtGW.Enabled = True
   txtGW.BackColor = &HFFFFFF
   
   chkCE.Enabled = True
   chkNonCE.Enabled = True
   chkWEEE.Enabled = True
   chkNonWEEE.Enabled = True
   chkChinaRoHS.Enabled = True
   chkNonChinaRoHS.Enabled = True
   
   chkCE4.Enabled = True
   chkNonCE4.Enabled = True
   chkWEEE4.Enabled = True
   chkNonWEEE4.Enabled = True
   chkChinaRoHS4.Enabled = True
   chkNonChinaRoHS4.Enabled = True
   
   chkTurkeyRohs.Enabled = True
   chkNonTurkeyRohs.Enabled = True
   
   chkFCC.Enabled = True
   chkNonFCC.Enabled = True
   
   chkRoHS.Enabled = True
   chkNonRoHS.Enabled = True
   'optH3CRoHS.Enabled = True
   'opt3COMRoHS.Enabled = True
   'optNonRoHS.Enabled = True
   
   
   txtMS.Enabled = True
   txtMS.BackColor = &HFFFFFF
   dtpMSValidFrom.Enabled = True
   txtNAL.Enabled = True
   txtNAL.BackColor = &HFFFFFF
   dtpValidFrom.Enabled = True
   
   chkSVPrint.Enabled = True
   chkNoPrintSV.Enabled = True
   
   Me.chkCN3C.Enabled = True
   Me.chkNonCN3C.Enabled = True
   
   Me.chkCN3C4.Enabled = True
   Me.chkNonCN3C4.Enabled = True
   
   Me.cbSalesLocations.Enabled = True
   
   Me.cbCNCA.Enabled = True
   Me.cbCNCN.Enabled = True
   Me.cbENCA.Enabled = True
   Me.cbENCN.Enabled = True
   
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
   CmdCancel.Enabled = True
   
   txtClass.Enabled = True
   txtModel.Enabled = True
   txtRatedValue.Enabled = True
   
   txtClass4.Enabled = True
   txtModel4.Enabled = True
   txtProduct.Enabled = True
   txtRatedValue4.Enabled = True
   
   txtCertification.Enabled = True
   
End Sub

Private Sub unable()
   txtSN.Enabled = False
   txtSN.BackColor = &HE0E0E0
   txtCPN.Enabled = False
   txtCPN.BackColor = &HE0E0E0
   txtEPN.Enabled = False
   txtEPN.BackColor = &HE0E0E0
   txtDes.Enabled = False
   txtDes.BackColor = &HE0E0E0
   txtOS.Enabled = False
   txtOS.BackColor = &HE0E0E0
   txtGW.Enabled = False
   txtGW.BackColor = &HE0E0E0
   
   chkCE.Enabled = False
   chkCE4.Enabled = False
   chkNonCE.Enabled = False
   chkNonCE4.Enabled = False
   chkWEEE.Enabled = False
   chkWEEE4.Enabled = False
   chkNonWEEE.Enabled = False
   chkNonWEEE4.Enabled = False
   chkChinaRoHS.Enabled = False
   chkChinaRoHS4.Enabled = False
   chkNonChinaRoHS.Enabled = False
   chkNonChinaRoHS4.Enabled = False
   chkTurkeyRohs.Enabled = False
   chkNonTurkeyRohs.Enabled = False
   'optH3CRoHS.Enabled = False
   'opt3COMRoHS.Enabled = False
   'optNonRoHS.Enabled = False
   chkRoHS.Enabled = False
   chkNonRoHS.Enabled = False
   chkFCC.Enabled = False
   chkNonFCC.Enabled = False
   
   txtMS.Enabled = False
   txtMS.BackColor = &HE0E0E0
   dtpMSValidFrom.Enabled = False
   txtNAL.Enabled = False
   txtNAL.BackColor = &HE0E0E0
   dtpValidFrom.Enabled = False
   chkSVPrint.Enabled = False
   chkNoPrintSV.Enabled = False
   
   txtHV.Enabled = False
   txtHV.BackColor = &HE0E0E0
   txtRemark.Enabled = False
   txtRemark.BackColor = &HE0E0E0
   
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = True
   cmdUpdate.Enabled = True
   cmdDelete.Enabled = True
   cmdConfirm.Enabled = False
   CmdCancel.Enabled = False
   Me.chkCN3C.Enabled = False
   Me.chkCN3C4.Enabled = False
   Me.chkNonCN3C.Enabled = False
   Me.chkNonCN3C4.Enabled = False
   Me.cbSalesLocations.Enabled = False
   Me.cbCNCN.Enabled = False
   Me.cbCNCA.Enabled = False
   Me.cbENCN.Enabled = False
   Me.cbENCA.Enabled = False
   
   txtClass.Enabled = False
   txtModel.Enabled = False
   txtRatedValue.Enabled = False
   txtCertification.Enabled = False
   txtClass4.Enabled = False
   txtModel4.Enabled = False
   txtProduct.Enabled = False
   txtRatedValue4.Enabled = False

   
   
End Sub

Private Sub Check1_Click()

End Sub


Private Sub chkCE_Click()
   If chkCE.Value = 1 Then
      chkNonCE.Value = 0
   Else
      chkNonCE.Value = 1
   End If
End Sub

Private Sub chkCE4_Click()
   If chkCE4.Value = 1 Then
      chkNonCE4.Value = 0
   Else
      chkNonCE4.Value = 1
   End If
End Sub

Private Sub chkCN3C_Click()
    If Me.chkCN3C.Value = 1 Then
        Me.chkNonCN3C.Value = 0
    End If
End Sub
Private Sub chkCN3C4_Click()
    If Me.chkCN3C4.Value = 1 Then
        Me.chkNonCN3C4.Value = 0
    Else
        Me.chkNonCN3C4.Value = 1
    End If
End Sub

Private Sub chkFCC_Click()
'    If chkFCC.Value = 1 Then
'        chkNonFCC.Value = 0
'    Else
'        chkNonFCC.Value = 1
'    End If
    
End Sub

Private Sub chkNonCE_Click()
   If chkNonCE.Value = 1 Then
      chkCE.Value = 0
   Else
      chkCE.Value = 1
   End If
End Sub
Private Sub chkNonCE4_Click()
   If chkNonCE4.Value = 1 Then
      chkCE4.Value = 0
   Else
      chkCE4.Value = 1
   End If
End Sub

Private Sub chkChinaROHS_Click()
   If chkChinaRoHS.Value = 1 Then
      chkNonChinaRoHS.Value = 0
   Else
      chkNonChinaRoHS.Value = 1
   End If
End Sub
Private Sub chkChinaROHS4_Click()
   If chkChinaRoHS4.Value = 1 Then
      chkNonChinaRoHS4.Value = 0
   Else
      chkNonChinaRoHS4.Value = 1
   End If
End Sub

Private Sub chkNonChinaRoHS_Click()
   If chkNonChinaRoHS.Value = 1 Then
      chkChinaRoHS.Value = 0
   Else
      chkChinaRoHS.Value = 1
   End If
End Sub
Private Sub chkNonChinaRoHS4_Click()
   If chkNonChinaRoHS4.Value = 1 Then
      chkChinaRoHS4.Value = 0
   Else
      chkChinaRoHS4.Value = 1
   End If
End Sub

Private Sub chkNonCN3C_Click()
    If chkNonCN3C.Value = 1 Then
        Me.chkCN3C.Value = 0
    End If
    
End Sub
Private Sub chkNonCN3C4_Click()
    If chkNonCN3C4.Value = 1 Then
        Me.chkCN3C4.Value = 0
    Else
        Me.chkCN3C4.Value = 1
    End If
    
End Sub

Private Sub chkNonFCC_Click()
'    If chkNonFCC.Value = 1 Then
'        chkFCC.Value = 0
'    Else
'        chkFCC.Value = 1
'    End If
End Sub

Private Sub chkNoPrintSV_Click()
   If chkNoPrintSV.Value = 1 Then
      chkSVPrint.Value = 0
   Else
      chkSVPrint.Value = 1
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
      chkNoPrintSV.Value = 0
   Else
      chkNoPrintSV.Value = 1
   End If
End Sub

Private Sub chkWEEE_Click()
   If chkWEEE.Value = 1 Then
      chkNonWEEE.Value = 0
   Else
      chkNonWEEE.Value = 1
   End If
End Sub
Private Sub chkWEEE4_Click()
   If chkWEEE4.Value = 1 Then
      chkNonWEEE4.Value = 0
   Else
      chkNonWEEE4.Value = 1
   End If
End Sub

Private Sub chkNonWEEE_Click()
   If chkNonWEEE.Value = 1 Then
      chkWEEE.Value = 0
   Else
      chkWEEE.Value = 1
   End If
End Sub
Private Sub chkNonWEEE4_Click()
   If chkNonWEEE4.Value = 1 Then
      chkWEEE4.Value = 0
   Else
      chkWEEE4.Value = 1
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
   unable
   op = ""
End Sub

Private Sub cmdConfirm_Click()
   If Trim(txtSN.Text) = "" Then
      MsgBox "产品编码不能为空!!", vbExclamation + vbOKOnly, "产品编码空"
      txtSN.SetFocus
      Exit Sub
   End If
   If txtCPN.Text = "" Then
       MsgBox "产品名称(中文)不能为空!", vbExclamation + vbOKOnly, "产品名称(中文)空"
       txtCPN.SetFocus
       Exit Sub
   End If
   If txtEPN.Text = "" Then
      MsgBox "产品名称(英文)不能为空!", vbExclamation + vbOKOnly, "产品名称(英文)空"
      txtEPN.SetFocus
      Exit Sub
   End If
   If txtDes.Text = "" Then
      MsgBox "产品描述不能为空!", vbExclamation + vbOKOnly, "产品描述空"
      txtDes.SetFocus
      Exit Sub
   End If
   If txtOS.Text = "" Then
      MsgBox "外尺寸不能为空!", vbExclamation + vbOKOnly, "外尺寸空"
      txtOS.SetFocus
      Exit Sub
   End If
   If txtOS.Text = "/" Then
      MsgBox "无外尺寸请维护N/A!", vbExclamation + vbOKOnly, "无外尺寸"
      txtOS.SetFocus
      Exit Sub
   End If
   If txtOS.Text = "n/a" Then
      txtOS.Text = UCase(txtOS.Text)
   End If

   If txtOS.Text <> "N/A" Then
   
        txtOS.Text = LTrim(RTrim(txtOS.Text))
      
        If Right(txtOS.Text, 2) <> "mm" Then
            MsgBox "外尺寸格式错误!", vbExclamation + vbOKOnly, "外尺寸错误"
            txtOS.SetFocus
            Exit Sub
        End If
        
        If InStr(txtOS.Text, "mmm") > 0 Then
            MsgBox "外尺寸格式错误!", vbExclamation + vbOKOnly, "外尺寸错误"
            txtOS.SetFocus
            Exit Sub
        End If
   End If
   
   
   
   'If txtGW.Text = "" Then
   '   MsgBox "毛重不能为空!", vbExclamation + vbOKOnly, "毛重空"
   '   txtGW.SetFocus
   '   Exit Sub
   'End If
    If Trim(txtGW.Text) <> "" Then
        If UCase(Right(Trim(txtGW.Text), 2)) <> "KG" Then
           MsgBox "毛重必须加上单位kg!", vbExclamation + vbOKOnly, "毛重单位空"
           txtGW.SetFocus
           Exit Sub
        End If
   End If
   If txtMS.Text = "" Then
      MsgBox "制造标准不能为空!", vbExclamation + vbOKOnly, "制造标准空"
      txtMS.SetFocus
      Exit Sub
   End If
   If txtNAL.Text = "" Then
      MsgBox "进网许可号不能为空!", vbExclamation + vbOKOnly, "进网许可号空"
      txtNAL.SetFocus
      Exit Sub
   End If
   If txtHV.Text = "" Then
      MsgBox "硬件版本不能为空!", vbExclamation + vbOKOnly, "硬件版本空"
      txtHV.SetFocus
      Exit Sub
   End If
   If chkSVPrint.Value = 0 And chkNoPrintSV.Value = 0 Then
      MsgBox "是否打印软件版本不能为空!", vbExclamation + vbOKOnly, "软件件版本空"
      txtHV.SetFocus
      Exit Sub
   End If
   
   'add by shun.huang requirement 2014/04/17
   If Me.chkFCC.Value + Me.chkNonFCC = 0 Or chkFCC.Value + chkNonFCC.Value = 2 Then
      MsgBox "FCC的值不能不选或者全选!", vbExclamation + vbOKOnly, "FCC选择错误"
      Exit Sub
   End If
   
   Dim CE, WEEE, ChinaRoHS, CE4, WEEE4, ChinaRoHS4, CN3C4, RoHS, TurkeyRoHS, SVPrint, CN3C, FCC, SalesLocation, CNCN, CNCA, ENCN, ENCA As String
   If chkCE.Value = 1 Then
      CE = "CE"
   ElseIf chkNonCE.Value = 1 Then
      CE = "N/A"
   End If
   If chkWEEE.Value = 1 Then
      WEEE = "WEEE"
   ElseIf chkNonWEEE.Value = 1 Then
      WEEE = "N/A"
   End If
   If chkChinaRoHS.Value = 1 Then
      ChinaRoHS = "China RoHS"
   ElseIf chkNonChinaRoHS.Value = 1 Then
      ChinaRoHS = "N/A"
   End If
   
   'add by carson 20150623 for 四合一
   If chkCE4.Value = 1 Then
      CE4 = "CE"
   ElseIf chkNonCE4.Value = 1 Then
      CE4 = "N/A"
   End If
   If chkWEEE4.Value = 1 Then
      WEEE4 = "WEEE"
   ElseIf chkNonWEEE4.Value = 1 Then
      WEEE4 = "N/A"
   End If
   If chkChinaRoHS4.Value = 1 Then
      ChinaRoHS4 = "China RoHS"
   ElseIf chkNonChinaRoHS4.Value = 1 Then
      ChinaRoHS4 = "N/A"
   End If
   If Me.chkCN3C4.Value = 1 Then
        CN3C4 = "Yes"
   ElseIf Me.chkNonCN3C4.Value = 1 Then
        CN3C4 = "No"
   Else
        CN3C4 = ""
   End If
   
   
   If chkTurkeyRohs.Value = 1 Then
      TurkeyRoHS = "Turkey RoHS"
   ElseIf chkNonTurkeyRohs.Value = 1 Then
      TurkeyRoHS = "N/A"
   End If
   If chkSVPrint.Value = 1 Then
      SVPrint = "Y"
   ElseIf chkNoPrintSV.Value = 1 Then
      SVPrint = "N"
   End If
   
   If chkRoHS.Value = 1 Then
'      RoHS = "H3C RoHS"
      RoHS = "HUV RoHS"
   ElseIf chkNonRoHS.Value = 1 Then
      RoHS = "N/A"
   End If
   
   If Me.chkCN3C.Value = 1 Then
        CN3C = "Yes"
   ElseIf Me.chkNonCN3C.Value = 1 Then
        CN3C = "No"
   Else
        CN3C = ""
   End If
   
   If Me.cbSalesLocations.ListIndex = 0 Then
        SalesLocation = "中英文"
   ElseIf Me.cbSalesLocations.ListIndex = 1 Then
        SalesLocation = "纯英文"
   ElseIf Me.cbSalesLocations.ListIndex = 2 Then
        SalesLocation = "讯美"
   Else
        SalesLocation = ""
   End If
    
   If Me.chkFCC.Value = 1 Then
        FCC = "Yes"
   Else
        FCC = "No"
   End If


'可选项：浙江宇视科技有限公司、上海天跃科技股份有限公司、无（选无时不打印此变量）
  If Me.cbCNCN.ListIndex = 0 Then
    CNCN = "浙江宇视科技有限公司"
  ElseIf Me.cbCNCN.ListIndex = 1 Then
    CNCN = "上海天跃科技股份有限公司"
  ElseIf Me.cbCNCN.ListIndex = 2 Then
    CNCN = "天津市亚安科技股份有限公司"
  ElseIf Me.cbCNCN.ListIndex = 3 Then
    CNCN = "无"
  Else
    CNCN = ""
  End If
  
   
   '中文地址：可选项：杭州市滨江区江陵路88号、上海市杨浦区赤峰路63号9楼A座、无（选无时不打印此变量）
   
   If Me.cbCNCA.ListIndex = 0 Then
        CNCA = "杭州市滨江区江陵路88号"
   ElseIf Me.cbCNCA.ListIndex = 1 Then
        CNCA = "上海市杨浦区赤峰路63号9楼A座"
   ElseIf Me.cbCNCA.ListIndex = 2 Then
        CNCA = "天津市南开区华苑产业园区梓苑路8号"
   ElseIf Me.cbCNCA.ListIndex = 3 Then
        CNCA = "无"
   Else
        CNCA = ""
   End If
   
   
   '英文公司名：可选项：Zhejiang Uniview Technologies Co.,Ltd 、无（选无时不打印此变量）
   If Me.cbENCN.ListIndex = 0 Then
    ENCN = "Zhejiang Uniview Technologies Co.,Ltd"
   ElseIf Me.cbENCN.ListIndex = 1 Then
    ENCN = "Tianjin YAAN Technology Co.,Ltd"
    ElseIf Me.cbENCN.ListIndex = 2 Then
    ENCN = "无"
   Else
    ENCN = ""
   End If
   
   
   
   '英文地址：可选项：No.88,Jiangling Road,Hangzhou,P.R.China、无（选无时不打印此变量）
   If Me.cbENCA.ListIndex = 0 Then
    ENCA = "No.88,Jiangling Road,Hangzhou,P.R.China"
   ElseIf Me.cbENCA.ListIndex = 1 Then
    ENCA = "No.8 Ziyuan Road,Huayuan Industrial Zone Tianjin 300384,China"
    ElseIf Me.cbENCA.ListIndex = 2 Then
    ENCA = "无"
   Else
    ENCA = ""
   End If
   
    classification = txtClass.Text
    model = txtModel.Text
    ratedValue = txtRatedValue.Text
    certification = txtCertification.Text
    
    classification4 = txtClass4.Text
    model4 = txtModel4.Text
    ratedValue4 = txtRatedValue4.Text
    productModel4 = txtProduct.Text
   
   
  ' If optH3CRoHS.Value = True Then
  '    RoHS = "H3C RoHS"
  ' ElseIf opt3COMRoHS.Value = True Then
  '    RoHS = "3COM RoHS"
  ' ElseIf optNonRoHS.Value = True Then
  '    RoHS = "/"
  ' End If
  
  txtGW.Text = LCase(Trim(txtGW.Text))
   
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from tblHUV where SN='" & Trim(txtSN.Text) & "' and HV='" & Trim(txtHV.Text) & "' "
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "产品编码&版本已存在!", vbExclamation + vbOKOnly, "产品编号重复"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close

      sql = "Insert into tblHUV(ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, CCC, SalesLocation, Remark, FCC, CNCN, CNCA, ENCN, ENCA, classification, model, rated_value, certification,CE4, WEEE4, ChinaRoHS4,CCC4,classification4,model4,rated_value4,productModel4) " & _
            "Values(" & getmaxID("tblHUV") & ",'" & Trim(txtHV.Text) & "','" & Trim(txtSN.Text) & "','" & Trim(txtCPN.Text) & "','" & Trim(txtEPN.Text) & "','" & Trim(txtDes.Text) & "','" & Trim(txtOS.Text) & "','" & Trim(txtGW.Text) & "','" & CE & "','" & WEEE & "','" & ChinaRoHS & "','" & RoHS & "','" & TurkeyRoHS & "'," & _
            "'" & txtMS.Text & "','" & dtpMSValidFrom.Value & "','" & txtNAL.Text & "','" & dtpValidFrom.Value & "','" & SVPrint & "','" & CN3C & "','" & SalesLocation & "','" & txtRemark.Text & "','" & FCC & "','" & CNCN & "','" & CNCA & "','" & ENCN & "','" & ENCA & "', '" & classification & "', '" & model & "', '" & ratedValue & "', '" & certification & "','" & CE4 & "','" & WEEE4 & "','" & ChinaRoHS4 & "','" & CN3C4 & "', '" & classification4 & "', '" & model4 & "', '" & ratedValue4 & "','" & productModel4 & "' )"
      
      sql = sql & " Insert into tblHUV_log( create_user,HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, CCC, SalesLocation, Remark, FCC, CNCN, CNCA, ENCN, ENCA, classification, model,rated_value, certification,CE4, WEEE4, ChinaRoHS4,CCC4,classification4,model4,rated_value4,productModel4,comment) " & _
            " Values('" & golUSERNAME & "','" & Trim(txtHV.Text) & "','" & Trim(txtSN.Text) & "','" & Trim(txtCPN.Text) & "','" & Trim(txtEPN.Text) & "','" & Trim(txtDes.Text) & "','" & Trim(txtOS.Text) & "','" & Trim(txtGW.Text) & "','" & CE & "','" & WEEE & "','" & ChinaRoHS & "','" & RoHS & "','" & TurkeyRoHS & "'," & _
            "'" & txtMS.Text & "','" & dtpMSValidFrom.Value & "','" & txtNAL.Text & "','" & dtpValidFrom.Value & "','" & SVPrint & "','" & CN3C & "','" & SalesLocation & "','" & txtRemark.Text & "','" & FCC & "','" & CNCN & "','" & CNCA & "','" & ENCN & "','" & ENCA & "', '" & classification & "', '" & model & "', '" & ratedValue & "', '" & certification & "','" & CE4 & "','" & WEEE4 & "','" & ChinaRoHS4 & "','" & CN3C4 & "', '" & classification4 & "', '" & model4 & "', '" & ratedValue4 & "','" & productModel4 & "','Insert' )"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "新增HUV设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "新增失败"
      End If
      MsgBox "新增HUV设定资料成功!", vbInformation + vbOKOnly, "新增成功"
      renovate
      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update tblHUV set CPN='" & Trim(txtCPN.Text) & "',EPN='" & Trim(txtEPN.Text) & "',Des='" & Trim(txtDes.Text) & "',OS='" & Trim(txtOS.Text) & "',GW='" & Trim(txtGW.Text) & "',CE='" & CE & "',WEEE='" & WEEE & "',ChinaRoHS='" & ChinaRoHS & "',RoHS='" & RoHS & "',TurkeyRohs='" & TurkeyRoHS & "'," & _
            "MS='" & txtMS.Text & "',MSValidFrom='" & dtpMSValidFrom.Value & "',NAL='" & txtNAL.Text & "',ValidFrom='" & dtpValidFrom.Value & "',HV='" & txtHV.Text & "',PrintSV='" & SVPrint & "',CCC = '" & CN3C & "',SalesLocation = '" & SalesLocation & "',Remark='" & txtRemark.Text & "',FCC = '" & FCC & "'," & _
            "CNCN = '" & CNCN & "',CNCA = '" & CNCA & "',ENCN = '" & ENCN & "',ENCA = '" & ENCA & "',classification='" & classification & "', model='" & model & "', rated_value='" & ratedValue & "', certification='" & certification & "',CE4='" & CE4 & "',WEEE4='" & WEEE4 & "',ChinaRoHS4='" & ChinaRoHS4 & "',CCC4 = '" & CN3C4 & "',classification4='" & classification4 & "', model4='" & model4 & "', rated_value4='" & ratedValue4 & "', productModel4='" & productModel4 & "'" & " where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and SN='" & Trim(txtSN.Text) & "'"
      
      sql = sql & " Insert into tblHUV_log( create_user,HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, CCC, SalesLocation, Remark, FCC, CNCN, CNCA, ENCN, ENCA, classification, model,rated_value, certification,CE4, WEEE4, ChinaRoHS4,CCC4,classification4,model4,rated_value4,productModel4,comment) " & _
            " Values('" & golUSERNAME & "','" & Trim(txtHV.Text) & "','" & Trim(txtSN.Text) & "','" & Trim(txtCPN.Text) & "','" & Trim(txtEPN.Text) & "','" & Trim(txtDes.Text) & "','" & Trim(txtOS.Text) & "','" & Trim(txtGW.Text) & "','" & CE & "','" & WEEE & "','" & ChinaRoHS & "','" & RoHS & "','" & TurkeyRoHS & "'," & _
            "'" & txtMS.Text & "','" & dtpMSValidFrom.Value & "','" & txtNAL.Text & "','" & dtpValidFrom.Value & "','" & SVPrint & "','" & CN3C & "','" & SalesLocation & "','" & txtRemark.Text & "','" & FCC & "','" & CNCN & "','" & CNCA & "','" & ENCN & "','" & ENCA & "', '" & classification & "', '" & model & "', '" & ratedValue & "', '" & certification & "','" & CE4 & "','" & WEEE4 & "','" & ChinaRoHS4 & "','" & CN3C4 & "', '" & classification4 & "', '" & model4 & "', '" & ratedValue4 & "','" & productModel4 & "','Update' )"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "修改HUV设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "修改失败"
      End If
      MsgBox "修改HUV设定资料成功!", vbInformation + vbOKOnly, "修改成功"
      renovate
      cmdCancel_Click
   End If
   renovate
End Sub

Private Sub cmdDelete_Click()
   If mfgH3C.RowSel <= 0 Then
      MsgBox "请选择要删除的行!", vbInformation + vbOKOnly, "未选择行"
      Exit Sub
   End If
   'if vbOK = MsgBox "确认要删除吗？", vbYesNo, "确认要删除当前选择的数据吗？" then
   If vbNo = MsgBox("错误的删除会导致产线停线，确认要删除当前选择的数据吗?", vbYesNo + vbQuestion, "确认要删除吗?") Then
        Exit Sub
   End If

   
   
   sql = " Insert into tblHUV_log( create_user,HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, CCC, SalesLocation, Remark, FCC, CNCN, CNCA, ENCN, ENCA, classification, model,rated_value, certification,CE4, WEEE4, ChinaRoHS4,CCC4,classification4,model4,rated_value4,productModel4,comment) " & _
         " select '" & golUSERNAME & "',HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, CCC, SalesLocation, Remark, FCC, CNCN, CNCA, ENCN, ENCA, classification, model, rated_value, certification,CE4, WEEE4, ChinaRoHS4,CCC4,classification4,model4,rated_value4,productModel4,'Delete' from tblHUV  where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and SN='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 3) & "'"

   sql = sql & " delete from tblHUV where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and SN='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 3) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "删除HUV设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "删除失败"
   End If
   MsgBox "删除HUV设定资料成功!", vbInformation + vbOKOnly, "删除成功"
   renovate
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If mfgH3C.Rows = 0 Then
      MsgBox "无资料可汇出", vbExclamation + vbOKOnly, "无资料"
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
      MsgBox "汇出到EXCEL资料成功!!", vbInformation + vbOKOnly, "汇出成功"
    End If
End Sub

Private Sub cmdImport_Click()
   If txtPath.Text = "" Then
      MsgBox "导入路径不能为空!", vbExclamation + vbOKOnly, "导入路径空"
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
             MsgBox "导入资料格式不正确!", vbExclamation + vbOKOnly, "格式错误"
             Exit Sub
          End If
          If Not ((xlSheet.Cells(j, 18) = "N") Or (xlSheet.Cells(j, 18) = "Y")) Then
             MsgBox "导入资料格式不正确!", vbExclamation + vbOKOnly, "格式错误"
             Exit Sub
          End If
          isexist = False
          For K = 1 To 19
          '======================================================
           If K = 3 Then
             cellValue = xlSheet.Cells(j, K)
             cellhvValue = xlSheet.Cells(j, 2)
             
             If cellValue = "" Or cellhvValue = "" Then
                MsgBox "导入资料格式不正确!", vbExclamation + vbOKOnly, "格式错误"
                Exit Sub
             End If
             
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from tblHUV where SN='" & cellValue & "' and HV='" & cellhvValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("产品编码&版本已存在!", vbAbortRetryIgnore + vbExclamation, "资料重复")
                End If
                
                If action = vbAbort Then
                   MsgBox "资料导入已终止!!", vbInformation + vbOKOnly, "导入终止"
                   rcd.Close
                   Exit Sub
                ElseIf action = vbIgnore And info = True Then
                   MsgBox "重复产品编号资料不会导入,请稍等..!!", vbInformation + vbOKOnly, "重复不会导入"
                   rcd.Close
                   info = False
                   Exit For
                ElseIf action = vbRetry And info = True Then
                   MsgBox "重复产品编号资料会自动更新,请稍等..!!", vbInformation + vbOKOnly, "重复会自动更新"
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
                   sql = "Update tblHUV set CPN='" & xlSheet.Cells(j, 4) & "',EPN='" & xlSheet.Cells(j, 5) & "',Des='" & xlSheet.Cells(j, 6) & "',OS='" & xlSheet.Cells(j, 7) & "',GW='" & xlSheet.Cells(j, 8) & "',CE='" & xlSheet.Cells(j, 9) & "',WEEE='" & xlSheet.Cells(j, 10) & "',ChinaRoHS='" & xlSheet.Cells(j, 11) & "'," & _
                        "RoHS='" & xlSheet.Cells(j, 12) & "',TurkeyRohs='" & xlSheet.Cells(j, 13) & "',MS='" & xlSheet.Cells(j, 14) & "',MSValidFrom='" & xlSheet.Cells(j, 15) & "',NAL='" & xlSheet.Cells(j, 16) & "',ValidFrom='" & xlSheet.Cells(j, 17) & "',PrintSV='" & xlSheet.Cells(j, 18) & "',Remark='" & xlSheet.Cells(j, 19) & "'" & _
                        " where SN='" & xlSheet.Cells(j, 3) & "' and HV='" & xlSheet.Cells(j, 2) & "' "
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "修改HUV设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "修改失败"
                   End If
'                   MsgBox "修改HUV设定资料成功!"
               ElseIf isexist = False Then
                   sql = "Insert into tblHUV(ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark) " & _
                        " Values(" & getmaxID("tblHUV") & ",'" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "','" & xlSheet.Cells(j, 4) & "','" & xlSheet.Cells(j, 5) & "','" & xlSheet.Cells(j, 6) & "','" & xlSheet.Cells(j, 7) & "','" & xlSheet.Cells(j, 8) & "','" & xlSheet.Cells(j, 9) & "','" & xlSheet.Cells(j, 10) & "','" & xlSheet.Cells(j, 11) & "'," & _
                        "'" & xlSheet.Cells(j, 12) & "','" & xlSheet.Cells(j, 13) & "','" & xlSheet.Cells(j, 14) & "','" & xlSheet.Cells(j, 15) & "','" & xlSheet.Cells(j, 16) & "','" & xlSheet.Cells(j, 17) & "','" & xlSheet.Cells(j, 18) & "','" & xlSheet.Cells(j, 19) & "')"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "新增HUV设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "修改失败"
                   End If
'                   MsgBox "新增HUV设定资料成功!"
               End If
           End If
         Next K
         
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "HUV设定资料导入成功!"
      renovate
End Sub

Private Sub cmdInsert_Click()
   enable
   txtSN.Text = ""
   txtCPN.Text = ""
   txtEPN.Text = ""
   txtDes.Text = ""
   txtOS.Text = ""
   txtGW.Text = ""
   
   chkCE.Value = 1
   chkWEEE.Value = 1
   chkChinaRoHS.Value = 1
   chkRoHS.Value = 1
   chkTurkeyRohs.Value = 1
   chkSVPrint.Value = 1
   
   txtNAL.Text = "N/A"
   dtpValidFrom.Value = Date
   
   txtMS.Text = "N/A"
   dtpMSValidFrom.Value = Date
   
   txtHV.Text = "N/A"
   txtRemark.Text = "Y"
   op = "Insert"
End Sub

Private Sub cmdQuery_Click()
    If txtSN.Enabled = False Then
      MsgBox "请按新增按钮清空就可输入查询内容!", vbOKOnly + vbInformation, "输入查询内容"
    End If
    If rec.State = 1 Then
        rec.Close
     End If
     sql = "select * from tblHUV Where 1=1"
     If txtSN.Text <> "" Then
        sql = sql & " and SN like '%" & txtSN.Text & "%'"
     End If
     If txtCPN.Text <> "" Then
        sql = sql & " and CPN like '%" & txtCPN.Text & "%'"
     End If
     If txtEPN.Text <> "" Then
        sql = sql & " and EPN='%" & txtEPN.Text & "%'"
     End If
     If txtDes.Text <> "" Then
        sql = sql & " and Des like '%" & txtDes.Text & "%'"
     End If
     If txtOS.Text <> "" Then
        sql = sql & " and OS like '%" & txtOS.Text & "%'"
     End If
     If txtGW.Text <> "" Then
        sql = sql & " and GW like '%" & txtGW.Text & "%'"
     End If
'     Dim CE, WEEE, ChinaRoHS, RoHS As String
'     If chkCE.Value = 1 Then
'        CE = "CE"
'     ElseIf chkNonCE.Value = 1 Then
'        CE = "/"
'     End If
'     If chkWEEE.Value = 1 Then
'        WEEE = "WEEE"
'     ElseIf chkNonWEEE.Value = 1 Then
'        WEEE = "/"
'     End If
'     If chkChinaRoHS.Value = 1 Then
'        ChinaRoHS = "China RoHS"
'     ElseIf chkNonChinaRoHS.Value = True Then
'        ChinaRoHS = "/"
'     End If
'     If optH3CRoHS.Value = 1 Then
'        RoHS = "H3C RoHS"
'     ElseIf opt3COMRoHS.Value = 1 Then
'        RoHS = "3COM RoHS"
'     ElseIf optNonRoHS.Value = 1 Then
'        RoHS = "/"
'     End If
'     If CE <> "" Then
'        sql = sql & " and CE='" & CE & "'"
'     End If
'     If WEEE <> "" Then
'        sql = sql & " and WEEE='" & WEEE & "'"
'     End If
'     If ChinaRoHS <> "" Then
'        sql = sql & " and ChinaRoHS='" & ChinaRoHS & "'"
'     End If
'     If RoHS <> "" Then
'        sql = sql & " and RoHS='" & RoHS & "'"
'     End If
'     If txtMS.Text <> "" Then
'        sql = sql & " and MS='" & txtMS.Text & "'"
'     End If
'     If txtNAL.Text <> "" Then
'        sql = sql & " and NAL='" & txtNAL.Text & "'"
'     End If
'     If txtChangNAL.Text <> "" Then
'        sql = sql & " and ChangNAL='" & txtChangNAL.Text & "'"
'     End If
'      If txtHV.Text <> "" Then
'        sql = sql & " and HV='" & txtHV.Text & "'"
'     End If
'     If txtRemark.Text <> "" Then
'        sql = sql & " and Remark='" & txtRemark.Text & "'"
'     End If
     sql = sql & " order by ID,SN"
     rec.Open sql, conn, adOpenKeyset, adLockOptimistic
     Set mfgH3C.DataSource = rec
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdSelect_Click()
   On Error Resume Next
   cdSelect.CancelError = True
   cdSelect.Filter = "*.xls|*.xls"
   cdSelect.action = 1
   If cdSelect.Filename <> "" Then txtPath.Text = cdSelect.Filename
End Sub

Private Sub cmdUpdate_Click()
   If mfgH3C.RowSel <= 0 Then
      MsgBox "请选择要修改的行!", vbInformation + vbOKOnly, "未选择行"
      Exit Sub
   End If
   mfgH3C_Click
   enable
   txtSN.Enabled = False
   txtSN.BackColor = &HE0E0E0
   op = "Update"
End Sub


Private Sub Form_Load()
   unable
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   Me.cbSalesLocations.AddItem ("中英文")
   Me.cbSalesLocations.AddItem ("纯英文")
   Me.cbSalesLocations.AddItem ("讯美")
   
   '可选项：浙江宇视科技有限公司、上海天跃科技股份有限公司、无（选无时不打印此变量）
   Me.cbCNCN.AddItem ("浙江宇视科技有限公司")
   Me.cbCNCN.AddItem ("上海天跃科技股份有限公司")
   Me.cbCNCN.AddItem ("天津市亚安科技股份有限公司")
   Me.cbCNCN.AddItem ("无")
   
   '中文地址：可选项：杭州市滨江区江陵路88号、上海市杨浦区赤峰路63号9楼A座、无（选无时不打印此变量）
   Me.cbCNCA.AddItem ("杭州市滨江区江陵路88号")
   Me.cbCNCA.AddItem ("上海市杨浦区赤峰路63号9楼A座")
   Me.cbCNCA.AddItem ("天津市南开区华苑产业园区梓苑路8号")
   Me.cbCNCA.AddItem ("无")
   
   
   
   '英文公司名：可选项：Zhejiang Uniview Technologies Co.,Ltd 、无（选无时不打印此变量）
   Me.cbENCN.AddItem ("Zhejiang Uniview Technologies Co.,Ltd")
   Me.cbENCN.AddItem ("Tianjin YAAN Technology Co.,Ltd")
   Me.cbENCN.AddItem ("无")
   
   
   '英文地址：可选项：No.88,Jiangling Road,Hangzhou,P.R.China、无（选无时不打印此变量）
   
   Me.cbENCA.AddItem ("No.88,Jiangling Road,Hangzhou,P.R.China")
   Me.cbENCA.AddItem ("No.8 Ziyuan Road,Huayuan Industrial Zone Tianjin 300384,China")
   Me.cbENCA.AddItem ("无")
   
   
   
   renovate
End Sub

Private Sub renovate()
   sql = "select * from tblHUV order by ID,SN"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgH3C.DataSource = rec
   With mfgH3C
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 3000
        .ColWidth(4) = 2500
        .ColWidth(5) = 3000
        .ColWidth(6) = 3500
        .ColWidth(7) = 3000
        .ColWidth(8) = 1500
        .ColWidth(9) = 1000
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        .ColWidth(13) = 1500
        .ColWidth(14) = 1500
        .ColWidth(15) = 2000
        .ColWidth(16) = 2000
        .ColWidth(17) = 1500
        .ColWidth(18) = 1000
        .ColWidth(19) = 1000
        .ColWidth(20) = 1000
        .ColWidth(21) = 2000
        .ColWidth(22) = 1000
        .ColWidth(23) = 2000
        .ColWidth(24) = 2000
        .ColWidth(25) = 2000
        .ColWidth(26) = 2000
        .ColWidth(27) = 2000
        .ColWidth(28) = 2000
        .ColWidth(29) = 2000
        .ColWidth(30) = 2000
        
        .ColWidth(31) = 3000
        .ColWidth(32) = 3000
        .ColWidth(33) = 3000
        .ColWidth(34) = 2500
        .ColWidth(35) = 2500
        .ColWidth(36) = 2500
        .ColWidth(37) = 2500
        
        .TextMatrix(0, 1) = "序号(ID)"
        .TextMatrix(0, 2) = "硬件版本(Hardware Version)"
        .TextMatrix(0, 3) = "产品编码(Model Number)"
        .TextMatrix(0, 4) = "产品名称(中文)(Chinese Product Name)"
        .TextMatrix(0, 5) = "产品名称(英文)(English Product Name)"
        .TextMatrix(0, 6) = "产品描述(Description)"
        .TextMatrix(0, 7) = "外箱尺寸(Outside Size)"
        .TextMatrix(0, 8) = "毛重(Gross Weight)"
        .TextMatrix(0, 9) = "认证信息CE"
        .TextMatrix(0, 10) = "认证信息WEEE"
        .TextMatrix(0, 11) = "认证信息ChinaRoHS"
        .TextMatrix(0, 12) = "认证信息RoHS"
        .TextMatrix(0, 13) = "认证信息TurkeyRoHS"
        .TextMatrix(0, 14) = "制造标准(China MFG Standards)"
        .TextMatrix(0, 15) = "制造标准有效期(Valid From)"
        .TextMatrix(0, 16) = "进网许可号(China N.A.L.)"
        .TextMatrix(0, 17) = "进网有效期(Valid From)"
        .TextMatrix(0, 18) = "是否打印软件版本"
        .TextMatrix(0, 19) = "CCC认证"
        .TextMatrix(0, 20) = "销售区域"
        .TextMatrix(0, 21) = "备注(Remark)"
        .TextMatrix(0, 22) = "FCC认证"
        .TextMatrix(0, 23) = "中文公司名"
        .TextMatrix(0, 24) = "中文地址名"
        .TextMatrix(0, 25) = "英文公司名"
        .TextMatrix(0, 26) = "英文地址名"
        .TextMatrix(0, 27) = "工艺分类"
        .TextMatrix(0, 28) = "认证型号"
        .TextMatrix(0, 29) = "电源额定值"
        .TextMatrix(0, 30) = "产品通过认证信息"
        
        .TextMatrix(0, 31) = "认证信息CE(四合一)"
        .TextMatrix(0, 32) = "认证信息WEEE(四合一)"
        .TextMatrix(0, 33) = "认证信息ChinaRoHS(四合一)"
        .TextMatrix(0, 34) = "CCC认证(四合一)"
        .TextMatrix(0, 35) = "工艺分类(四合一)"
        .TextMatrix(0, 36) = "认证型号(四合一)"
        .TextMatrix(0, 37) = "电源额定值(四合一)"
        .TextMatrix(0, 38) = "产品型号(四合一)"
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
   If mfgH3C.RowSel > 0 Then
      txtHV.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 2)
      txtSN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 3)
      txtCPN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 4)
      txtEPN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 5)
      txtDes.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 6)

      txtOS.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 7)
      txtGW.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 8)
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 9))) = "CE" Then
         chkCE.Value = 1
         chkNonCE.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 9) = "N/A" Then
         chkCE.Value = 0
         chkNonCE.Value = 1
      End If
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 10))) = "WEEE" Then
         chkWEEE.Value = 1
         chkNonWEEE.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 10) = "N/A" Then
         chkWEEE.Value = 0
         chkNonWEEE.Value = 1
      End If
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 11))) = "CHINA ROHS" Then
         chkChinaRoHS.Value = 1
         chkNonChinaRoHS.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 11) = "N/A" Then
         chkChinaRoHS.Value = 0
         chkNonChinaRoHS.Value = 1
      End If
      
      'add by carson 20150623 for 四合一
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 31))) = "CE" Then
         chkCE4.Value = 1
         chkNonCE4.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 31) = "N/A" Then
         chkCE4.Value = 0
         chkNonCE4.Value = 1
      End If
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 32))) = "WEEE" Then
         chkWEEE4.Value = 1
         chkNonWEEE4.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 32) = "N/A" Then
         chkWEEE4.Value = 0
         chkNonWEEE4.Value = 1
      End If
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 33))) = "CHINA ROHS" Then
         chkChinaRoHS4.Value = 1
         chkNonChinaRoHS4.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 33) = "N/A" Then
         chkChinaRoHS4.Value = 0
         chkNonChinaRoHS4.Value = 1
      End If
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 34))) = "YES" Then
        Me.chkCN3C4.Value = 1
        Me.chkNonCN3C4.Value = 0
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 34))) = "NO" Then
        Me.chkCN3C4.Value = 0
        Me.chkNonCN3C4.Value = 1
      End If
      
'      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 12))) = "H3C ROHS" Then
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 12))) = "HUV ROHS" Then
        chkRoHS.Value = 1
        chkNonRoHS.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 12) = "N/A" Then
         chkRoHS.Value = 0
         chkNonRoHS.Value = 1
      End If
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 13))) = "TURKEY ROHS" Then
         chkTurkeyRohs.Value = 1
         chkNonTurkeyRohs.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 13) = "N/A" Then
         chkTurkeyRohs.Value = 0
         chkNonTurkeyRohs.Value = 1
      End If
      
      txtMS.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 14)
      dtpMSValidFrom.Value = mfgH3C.TextMatrix(mfgH3C.RowSel, 15)
      
      txtNAL.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 16)
      dtpValidFrom.Value = mfgH3C.TextMatrix(mfgH3C.RowSel, 17)
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 18))) = "Y" Then
        chkSVPrint.Value = 1
        chkNoPrintSV.Value = 0
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 18))) = "N" Then
        chkSVPrint.Value = 0
        chkNoPrintSV.Value = 1
      End If
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 19))) = "YES" Then
        Me.chkCN3C.Value = 1
        Me.chkNonCN3C.Value = 0
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 19))) = "NO" Then
        Me.chkCN3C.Value = 0
        Me.chkNonCN3C.Value = 1
      End If
      
      
      
      '中英文、纯英文、讯美
      'update by allen.yan 2014/06/07
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 20))) = "中英文" Then
        Me.cbSalesLocations.ListIndex = 0
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 20))) = "纯英文" Then
        Me.cbSalesLocations.ListIndex = 1
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 20))) = "讯美" Then
        Me.cbSalesLocations.ListIndex = 2
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 20))) = "" Then
        Me.cbSalesLocations.ListIndex = -1
      End If
            
      txtRemark.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 21)
      
      'add by allen.yan for FCC column
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 22))) = "YES" Then
        Me.chkFCC.Value = 1
        Me.chkNonFCC.Value = 0
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 22))) = "NO" Then
        Me.chkFCC.Value = 0
        Me.chkNonFCC.Value = 1
      Else
        Me.chkFCC.Value = 0
        Me.chkNonFCC.Value = 0
      End If
      
      '中文公司名，中文地址，英文公司名称，英文地址
      '可选项：浙江宇视科技有限公司、上海天跃科技股份有限公司、无（选无时不打印此变量）
      'add by allen.yan 2014/06/07
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 23))) = "浙江宇视科技有限公司" Then
        Me.cbCNCN.ListIndex = 0
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 23))) = "上海天跃科技股份有限公司" Then
        Me.cbCNCN.ListIndex = 1
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 23))) = "天津市亚安科技股份有限公司" Then
        Me.cbCNCN.ListIndex = 2
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 23))) = "无" Then
        Me.cbCNCN.ListIndex = 3
      Else
        Me.cbCNCN.ListIndex = -1
      End If
   
        
    '中文地址：可选项：杭州市滨江区江陵路88号、上海市杨浦区赤峰路63号9楼A座、无（选无时不打印此变量）
    'add by allen.yan 2014/06/07
       
         If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 24))) = "杭州市滨江区江陵路88号" Then
           Me.cbCNCA.ListIndex = 0
         ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 24))) = "上海市杨浦区赤峰路63号9楼A座" Then
           Me.cbCNCA.ListIndex = 1
        ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 24))) = "天津市南开区华苑产业园区梓苑路8号" Then
           Me.cbCNCA.ListIndex = 2
         ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 24))) = "无" Then
           Me.cbCNCA.ListIndex = 3
         Else
           Me.cbCNCA.ListIndex = -1
         End If
         
         
         
         
'       '英文公司名：可选项：Zhejiang Uniview Technologies Co.,Ltd 、无（选无时不打印此变量）
'   Me.cbENCN.AddItem ("Zhejiang Uniview Technologies Co.,Ltd")
'   Me.cbENCN.AddItem ("无")
'
'
'   '英文地址：可选项：No.88,Jiangling Road,Hangzhou,P.R.China、无（选无时不打印此变量）
'
'   Me.cbENCA.AddItem ("No.88,Jiangling Road,Hangzhou,P.R.China")
'   Me.cbENCA.AddItem ("无")
'
    '英文公司名：可选项：Zhejiang Uniview Technologies Co.,Ltd 、无（选无时不打印此变量）
    'add by allen.yan 2014/06/07
        
         If Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 25)) = "Zhejiang Uniview Technologies Co.,Ltd" Then
           Me.cbENCN.ListIndex = 0
         ElseIf Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 25)) = "Tianjin YAAN Technology Co.,Ltd" Then
           Me.cbENCN.ListIndex = 1
         ElseIf Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 25)) = "无" Then
           Me.cbENCN.ListIndex = 2
         Else
           Me.cbENCN.ListIndex = -1
         End If
        
    '英文地址：可选项：No.88,Jiangling Road,Hangzhou,P.R.China、无（选无时不打印此变量）
        
        If Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 26)) = "No.88,Jiangling Road,Hangzhou,P.R.China" Then
           Me.cbENCA.ListIndex = 0
        ElseIf Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 26)) = "No.8 Ziyuan Road,Huayuan Industrial Zone Tianjin 300384,China" Then
           Me.cbENCA.ListIndex = 1
        ElseIf Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 26)) = "无" Then
           Me.cbENCA.ListIndex = 2
        Else
           Me.cbENCA.ListIndex = -1
        End If
      
      
      txtClass.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 27)
      txtModel.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 28)
      txtRatedValue.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 29)
      txtCertification.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 30)
      
      'add by carson 20150623 for 四合一
      txtClass4.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 35)
      txtModel4.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 36)
      txtRatedValue4.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 37)
      txtProduct.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 38)
      
      
   End If
End Sub

Private Sub mfgH3C_SelChange()
   mfgH3C_Click
End Sub

