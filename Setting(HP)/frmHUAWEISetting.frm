VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmHUAWEISetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HUAWEI Setting(Test)"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHUAWEISetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cb5000 
      Height          =   465
      ItemData        =   "frmHUAWEISetting.frx":2E1A
      Left            =   8760
      List            =   "frmHUAWEISetting.frx":2E2A
      TabIndex        =   54
      Top             =   4320
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdSelect 
      Left            =   2520
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10740
      TabIndex        =   42
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9300
      TabIndex        =   41
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "确定(Confirm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7620
      TabIndex        =   40
      Top             =   9360
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(Delete)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10740
      TabIndex        =   39
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修改(Update)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9300
      TabIndex        =   38
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "新增(Insert)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7620
      TabIndex        =   37
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查询(Query)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5940
      TabIndex        =   36
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "导出(Export)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3420
      TabIndex        =   35
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "导入(Import)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3420
      TabIndex        =   34
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择(Select)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1380
      TabIndex        =   33
      Top             =   9600
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   32
      Top             =   9000
      Width           =   3015
   End
   Begin VB.Frame fmH3C 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.CheckBox chkPrintPC 
         Caption         =   "是"
         Height          =   375
         Left            =   8760
         TabIndex        =   57
         Top             =   5040
         Width           =   855
      End
      Begin VB.CheckBox chkNonPrintPC 
         Caption         =   "否"
         Height          =   375
         Left            =   10080
         TabIndex        =   56
         Top             =   5040
         Width           =   855
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         Caption         =   "无"
         Height          =   375
         Left            =   10680
         TabIndex        =   52
         Top             =   2520
         Width           =   855
      End
      Begin VB.CheckBox chkChinaRoHS 
         Caption         =   "有"
         Height          =   375
         Left            =   9120
         TabIndex        =   51
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chkNoPrintSV 
         Caption         =   "否"
         Height          =   375
         Left            =   4200
         TabIndex        =   49
         Top             =   4320
         Width           =   855
      End
      Begin VB.CheckBox chkSVPrint 
         Caption         =   "是"
         Height          =   375
         Left            =   2760
         TabIndex        =   48
         Top             =   4320
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpMSValidFrom 
         Height          =   495
         Left            =   8760
         TabIndex        =   46
         Top             =   3600
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
         Format          =   21364737
         UpDown          =   -1  'True
         CurrentDate     =   40427
      End
      Begin VB.TextBox txtSN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtCPN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9120
         TabIndex        =   16
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtEPN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2280
         TabIndex        =   15
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtDes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9120
         TabIndex        =   14
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtOS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2280
         TabIndex        =   13
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtGW 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9120
         TabIndex        =   12
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtNAL 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   11
         Top             =   3000
         Width           =   2895
      End
      Begin VB.CheckBox chkCE 
         Caption         =   "CE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox chkNonCE 
         Caption         =   "无 CE"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkWEEE 
         Caption         =   "有"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox chkNonWEEE 
         Caption         =   "无"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10680
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkRoHS 
         Caption         =   "有"
         Height          =   330
         Left            =   2280
         TabIndex        =   6
         Top             =   2520
         Width           =   855
      End
      Begin VB.CheckBox chkNonRoHS 
         Caption         =   "无"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtMS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         TabIndex        =   4
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox txtHV 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   1
         Top             =   4920
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpValidFrom 
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   3600
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
         Format          =   21364737
         UpDown          =   -1  'True
         CurrentDate     =   39757
      End
      Begin VB.Label Label2 
         Caption         =   "打印电源代码："
         Height          =   495
         Left            =   6480
         TabIndex        =   55
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "5000状态:"
         Height          =   495
         Left            =   6960
         TabIndex        =   53
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label lblChinaRoHS 
         Caption         =   "认证信息China RoHS:"
         Height          =   495
         Left            =   6240
         TabIndex        =   50
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label lblPrintSV 
         Caption         =   "打印软件版本:"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   4320
         Width           =   2535
      End
      Begin VB.Label lblMSValidFrom 
         Caption         =   "制造标准有效期:"
         Height          =   375
         Left            =   6240
         TabIndex        =   45
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label lblOS 
         Caption         =   "外尺寸(MM):"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         Caption         =   "进网许可号:"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lblSN 
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblCPN 
         Caption         =   "产品名称(中文):"
         Height          =   375
         Left            =   6840
         TabIndex        =   28
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         Caption         =   "产品名称(英文):"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblDes 
         Caption         =   "产品描述:"
         Height          =   375
         Left            =   6840
         TabIndex        =   26
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblGW 
         Caption         =   "毛重(KG):"
         Height          =   375
         Left            =   6840
         TabIndex        =   25
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblValidFrom 
         Caption         =   "进网有效期:"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label lblCE 
         Caption         =   "认证信息CE:"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblWEEE 
         Caption         =   "认证信息WEEE:"
         Height          =   375
         Left            =   6840
         TabIndex        =   22
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblRoHS 
         Caption         =   "认证信息RoHS:"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblMS 
         Caption         =   "制造标准:"
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lblHV 
         Caption         =   "硬件版本:"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label lblRemark 
         Caption         =   "备注:"
         Height          =   495
         Left            =   2880
         TabIndex        =   18
         Top             =   4920
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgHUAWEI 
      Height          =   2535
      Left            =   0
      TabIndex        =   30
      Top             =   5760
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4471
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblPath 
      Caption         =   "导入/导出路径:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   31
      Top             =   8520
      Width           =   2175
   End
End
Attribute VB_Name = "frmHUAWEISetting"
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
   chkRoHS.Enabled = True
   chkChinaRoHS.Enabled = True
   chkNonChinaRoHS.Enabled = True
   chkNonRoHS.Enabled = True
   
   txtMS.Enabled = True
   txtMS.BackColor = &HFFFFFF
   dtpMSValidFrom.Enabled = True
   
   txtNAL.Enabled = True
   txtNAL.BackColor = &HFFFFFF
   dtpValidFrom.Enabled = True

   chkSVPrint.Enabled = True
   chkNoPrintSV.Enabled = True
   
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
   chkNonCE.Enabled = False
   chkWEEE.Enabled = False
   chkNonWEEE.Enabled = False
   chkRoHS.Enabled = False
   chkNonRoHS.Enabled = False
   chkChinaRoHS.Enabled = False
   chkNonChinaRoHS.Enabled = False
   
   txtMS.Enabled = False
   txtMS.BackColor = &HE0E0E0
   Me.dtpMSValidFrom.Enabled = False
   txtNAL.Enabled = False
   txtNAL.BackColor = &HE0E0E0
   dtpValidFrom.Enabled = False
   
   txtHV.Enabled = False
   txtHV.BackColor = &HE0E0E0
   txtRemark.Enabled = False
   txtRemark.BackColor = &HE0E0E0
   chkSVPrint.Enabled = False
   chkNoPrintSV.Enabled = False
   
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = True
   cmdUpdate.Enabled = True
   cmdDelete.Enabled = True
   cmdConfirm.Enabled = False
   cmdCancel.Enabled = False
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

Private Sub chkNonPrintPC_Click()
    If Me.chkNonPrintPC.Value = 1 Then
        Me.chkPrintPC.Value = 0
    End If
End Sub

Private Sub chkPrintPC_Click()
    If Me.chkPrintPC.Value = 1 Then
        Me.chkNonPrintPC.Value = 0
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

Private Sub chkNoPrintSV_Click()
   If chkNoPrintSV.Value = 1 Then
      chkSVPrint.Value = 0
   Else
      chkSVPrint.Value = 1
   End If
End Sub

Private Sub chkSVPrint_Click()
   If chkSVPrint.Value = 1 Then
      chkNoPrintSV.Value = 0
   Else
      chkNoPrintSV.Value = 1
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
        MsgBox "是否打印软件版本不能为空!", vbExclamation + vbOKOnly, "软件版本空"
        Exit Sub
   End If
   
   Dim CE, WEEE, ChinaRoHS, RoHS, SVPrint, PCPrint, ftStatus As String
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
   If chkRoHS.Value = 1 Then
      RoHS = "Huawei RoHS"
   ElseIf chkNonRoHS.Value = 1 Then
      RoHS = "N/A"
   End If
   If chkChinaRoHS.Value = 1 Then
        ChinaRoHS = "China RoHS"
   ElseIf chkNonChinaRoHS.Value = 1 Then
        ChinaRoHS = "N/A"
   End If
   If chkSVPrint.Value = 1 Then
      SVPrint = "Y"
   ElseIf chkNoPrintSV.Value = 1 Then
      SVPrint = "N"
   End If
   
   If Me.chkPrintPC.Value = 1 Then
        PCPrint = "1"
    Else
        PCPrint = "0"
    End If
       
    If Me.cb5000.ListIndex <= -1 Then
        MsgBox "5000状态没有选择!", vbExclamation + vbOKOnly, "请选择5000状态的一个选项"
         Me.cb5000.SetFocus
         Exit Sub
    End If
    
'    Y，N，NA,TBD
    If Me.cb5000.ListIndex = 0 Then
        ftStatus = "Y"
    ElseIf Me.cb5000.ListIndex = 1 Then
        ftStatus = "N"
    ElseIf Me.cb5000.ListIndex = 2 Then
        ftStatus = "NA"
    ElseIf Me.cb5000.ListIndex = 3 Then
        ftStatus = "TBD"
    End If
    
    txtGW.Text = LCase(Trim(txtGW.Text))
     
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from tblHuaWei where SN='" & Trim(txtSN.Text) & "' and HV='" & Trim(txtHV.Text) & "'"
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "产品编码&版本已存在!"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close

      sql = "Insert into tblHuaWei(ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, MS, MSValidFrom, NAL, ValidFrom, PrintSV,[5000_Status],Print_Power,Remark) " & _
            "Values(" & getmaxID("tblHuaWei") & ",'" & Trim(txtHV.Text) & "','" & Trim(txtSN.Text) & "','" & Trim(txtCPN.Text) & "','" & Trim(txtEPN.Text) & "','" & Trim(txtDes.Text) & "','" & Trim(txtOS.Text) & "','" & Trim(txtGW.Text) & "','" & CE & "','" & WEEE & "','" & ChinaRoHS & "','" & RoHS & "'," & _
            "'" & txtMS.Text & "','" & Me.dtpMSValidFrom.Value & "','" & txtNAL.Text & "','" & dtpValidFrom.Value & "','" & SVPrint & "','" & ftStatus & "'," & PCPrint & ",'" & txtRemark.Text & "')"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "新增HUAWEI设定资料失败!" & "原因是" & status
      End If
      MsgBox "新增HUAWEI设定资料成功!"
      renovate
      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update tblHuaWei set HV='" & Trim(txtHV.Text) & "',CPN='" & Trim(txtCPN.Text) & "',EPN='" & Trim(txtEPN.Text) & "',Des='" & Trim(txtDes.Text) & "',OS='" & Trim(txtOS.Text) & "',GW='" & Trim(txtGW.Text) & "',CE='" & CE & "',WEEE='" & WEEE & "',chinarohs='" & ChinaRoHS & "',RoHS='" & RoHS & "'," & _
            "MS='" & txtMS.Text & "',MSValidFrom='" & Me.dtpMSValidFrom.Value & "',NAL='" & txtNAL.Text & "',ValidFrom='" & dtpValidFrom.Value & "',PrintSV='" & SVPrint & "',Print_Power = " & PCPrint & ",[5000_Status] = '" & ftStatus & "',Remark='" & txtRemark.Text & "'" & _
            " where ID=" & mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 1) & " and SN='" & Trim(txtSN.Text) & "'"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "修改HUAWEI设定资料失败!" & "原因是" & status
      End If
      MsgBox "修改HUAWEI设定资料成功!"
      renovate
      cmdCancel_Click
   End If
   renovate
End Sub

Private Sub cmdDelete_Click()
   If mfgHUAWEI.RowSel <= 0 Then
      MsgBox "请选择要删除的行!"
      Exit Sub
   End If
   sql = "delete from tblHuaWei where ID=" & mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 1) & " and SN='" & mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 3) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "删除HUAWEI设定资料失败!" & "原因是" & status
   End If
   MsgBox "删除HUAWEI设定资料成功!"
   renovate
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If mfgHUAWEI.Rows = 0 Then
      MsgBox "无资料可汇出"
      Exit Sub
   End If
   If txtPath.Text <> "" Then
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets.Item(1)
       For i = 0 To mfgHUAWEI.Rows - 1
         For j = 1 To mfgHUAWEI.Cols - 1
          xlSheet.Cells(i + 1, j) = mfgHUAWEI.TextMatrix(i, j)
       Next j
      Next i
      xlBook.SaveAs (txtPath.Text)
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "汇出到EXCEL资料成功!!"
    End If
End Sub

Private Sub cmdImport_Click()
   If txtPath.Text = "" Then
      MsgBox "导入路径不能为空!"
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
          If xlSheet.Cells(j, 18) = "" Then
             MsgBox "导入资料格式不正确!"
             Exit Sub
          End If
          If Not ((xlSheet.Cells(j, 17) = "N") Or (xlSheet.Cells(j, 17) = "Y")) Then
             MsgBox "导入资料格式不正确!", vbExclamation + vbOKOnly, "格式错误"
             Exit Sub
          End If
          isexist = False
          '=================================================
          For K = 1 To 18
           If K = 3 Then
             cellValue = xlSheet.Cells(j, K)
             cellhvValue = xlSheet.Cells(j, 2)
             
             If cellValue = "" Then
                MsgBox "导入资料格式不正确!"
                Exit Sub
             End If
             
             If cellhvValue = "" Then
                MsgBox "导入资料格式不正确!"
                Exit Sub
             End If
             
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from tblHuaWei where SN='" & cellValue & "' and HV='" & cellhvValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("产品编码&版本已存在!", vbAbortRetryIgnore + vbExclamation, "资料重复")
                End If
                
                If action = vbAbort Then
                   MsgBox "资料导入已终止!!"
                   rcd.Close
                   Exit Sub
                ElseIf action = vbIgnore And info = True Then
                   MsgBox "重复产品编号资料不会导入,请稍等..!!"
                   rcd.Close
                   info = False
                   Exit For
                ElseIf action = vbRetry And info = True Then
                   MsgBox "重复产品编号资料会自动更新,请稍等..!!"
                   info = False
                End If
                isexist = True
             Else
                isexist = False
             End If
             rcd.Close
            End If
            
            If K = 18 Then
               If action = vbRetry Then
                   sql = "Update tblHuaWei set CPN='" & xlSheet.Cells(j, 4) & "',EPN='" & xlSheet.Cells(j, 5) & "',Des='" & xlSheet.Cells(j, 6) & "',OS='" & xlSheet.Cells(j, 7) & "',GW='" & xlSheet.Cells(j, 8) & "',CE='" & xlSheet.Cells(j, 9) & "',WEEE='" & xlSheet.Cells(j, 10) & "',chinarohs='" & xlSheet.Cells(j, 11) & "',RoHS='" & xlSheet.Cells(j, 12) & "'," & _
                        "MS='" & xlSheet.Cells(j, 13) & "',MSValidFrom='" & xlSheet.Cells(j, 14) & "',NAL='" & xlSheet.Cells(j, 15) & "',ValidFrom='" & xlSheet.Cells(j, 16) & "',PrintSV='" & xlSheet.Cells(j, 17) & "',Remark='" & xlSheet.Cells(j, 18) & "'" & _
                        " where SN='" & xlSheet.Cells(j, 3) & "' and HV='" & xlSheet.Cells(j, 2) & "'"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "修改HUAWEI设定资料失败!" & "原因是" & status
                   End If
                   
               ElseIf isexist = False Then
                   sql = "Insert into tblHuaWei(ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, MS, MSValidFrom, NAL, ValidFrom, PrintSV,5000_status Remark) " & _
                   " Values(" & getmaxID("tblHuaWei") & ",'" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "','" & xlSheet.Cells(j, 4) & "','" & xlSheet.Cells(j, 5) & "','" & xlSheet.Cells(j, 6) & "','" & xlSheet.Cells(j, 7) & "','" & xlSheet.Cells(j, 8) & "','" & xlSheet.Cells(j, 9) & "','" & xlSheet.Cells(j, 10) & "','" & xlSheet.Cells(j, 11) & "','" & xlSheet.Cells(j, 12) & "','" & xlSheet.Cells(j, 13) & "','" & xlSheet.Cells(j, 14) & "','" & xlSheet.Cells(j, 15) & "','" & xlSheet.Cells(j, 16) & "','" & xlSheet.Cells(j, 17) & "','" & xlSheet.Cells(j, 18) & "')"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "新增HUAWEI设定资料失败!" & "原因是" & status
                   End If
                   
               End If
           End If
         Next K
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "HUAWEI设定资料导入成功!"
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
   txtMS.Text = "N/A"
   Me.dtpMSValidFrom.Value = Date
   txtNAL.Text = "N/A"
   dtpValidFrom.Value = Date
   
   chkCE.Value = 1
   chkWEEE.Value = 1
   chkRoHS.Value = 1
   chkChinaRoHS.Value = 1
   
   Me.chkSVPrint.Value = 1
   txtHV.Text = "N/A"
   txtRemark.Text = "N/A"
   op = "Insert"
End Sub

Private Sub cmdQuery_Click()
   MsgBox "请按新增按钮清空就可输入查询内容!", vbOKOnly + vbInformation, "输入查询内容"
   If rec.State = 1 Then
      rec.Close
   End If
   sql = "select * from tblHuaWei Where 1=1"
   If txtSN.Text <> "" Then
      sql = sql & " and SN like '%" & txtSN.Text & "%'"
   End If
   If txtCPN.Text <> "" Then
      sql = sql & " and CPN like '%" & txtCPN.Text & "%'"
   End If
   If txtEPN.Text <> "" Then
      sql = sql & " and EPN like '%" & txtEPN.Text & "%'"
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
'   Dim CE, WEEE, RoHS As String
'   If chkCE.Value = 1 Then
'      CE = "CE"
'   ElseIf chkNonCE.Value = 1 Then
'      CE = "/"
'   End If
'   If chkWEEE.Value = 1 Then
'      WEEE = "WEEE"
'   ElseIf chkNonWEEE.Value = 1 Then
'      WEEE = "/"
'   End If
'   If chkRoHS.Value = 1 Then
'      RoHS = "RoHS"
'   ElseIf chkNonRoHS.Value = 1 Then
'      RoHS = "/"
'   End If
'   If CE <> "" Then
'      sql = sql & " and CE='" & CE & "'"
'   End If
'   If WEEE <> "" Then
'      sql = sql & " and WEEE='" & WEEE & "'"
'   End If
'   If RoHS <> "" Then
'      sql = sql & " and RoHS='" & RoHS & "'"
'   End If
'   If txtMS.Text <> "" Then
'      sql = sql & " and MS='" & txtMS.Text & "'"
'   End If
'   If txtNAL.Text <> "" Then
'      sql = sql & " and NAL='" & txtNAL.Text & "'"
'   End If
'   If txtChangNAL.Text <> "" Then
'      sql = sql & " and ChangNAL='" & txtChangNAL.Text & "'"
'   End If
'    If txtHV.Text <> "" Then
'      sql = sql & " and HV='" & txtHV.Text & "'"
'   End If
'   If txtRemark.Text <> "" Then
'      sql = sql & " and Remark='" & txtRemark.Text & "'"
'   End If
   sql = sql & " order by ID,SN"
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgHUAWEI.DataSource = rec
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdSelect_Click()
   On Error Resume Next
   cdSelect.CancelError = True
   cdSelect.Filter = "*.xls|*.xls"
   cdSelect.action = 1
   If cdSelect.FileName <> "" Then txtPath.Text = cdSelect.FileName
End Sub

Private Sub cmdUpdate_Click()
   If mfgHUAWEI.RowSel <= 0 Then
      MsgBox "请选择要修改的行!"
      Exit Sub
   End If
   mfgHUAWEI_Click
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
   renovate
End Sub

Private Sub renovate()
   sql = "select * from tblHuaWei order by ID,SN"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockReadOnly
   Set mfgHUAWEI.DataSource = rec
   With mfgHUAWEI
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 3500
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .ColWidth(5) = 3000
        .ColWidth(6) = 3500
        .ColWidth(7) = 1500
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1000
        .ColWidth(11) = 1500
        .ColWidth(12) = 1500
        .ColWidth(13) = 1500
        .ColWidth(14) = 1500
        .ColWidth(15) = 1500
        .ColWidth(16) = 1000
        .ColWidth(17) = 2000
        .ColWidth(18) = 1000
        .ColWidth(19) = 1000
        .ColWidth(20) = 500
        
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
        .TextMatrix(0, 11) = "ChinaRoHS"
        .TextMatrix(0, 12) = "认证信息RoHS"
        .TextMatrix(0, 13) = "制造标准(China MFG Standards)"
        .TextMatrix(0, 14) = "制造标准有效期"
        .TextMatrix(0, 15) = "进网许可号(China N.A.L.)"
        .TextMatrix(0, 16) = "进网有效期(Valid From)"
        .TextMatrix(0, 17) = "是否打印软件版本"
        .TextMatrix(0, 18) = "5000米状态"
        .TextMatrix(0, 19) = "电源代码打印"
        .TextMatrix(0, 20) = "备注(Remark)"
   End With
   rec.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rec.State = 1 Then
      rec.Close
      Set rec = Nothing
   End If
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub

Private Sub mfgHUAWEI_Click()
   If mfgHUAWEI.RowSel > 0 Then
      txtSN.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 3)
      txtCPN.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 4)
      txtEPN.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 5)
      txtDes.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 6)
      txtOS.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 7)
      txtGW.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 8)
      
      If UCase(Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 9))) = "CE" Then
         chkCE.Value = 1
         chkNonCE.Value = 0
      ElseIf mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 9) = "N/A" Then
         chkCE.Value = 0
         chkNonCE.Value = 1
      End If
      If UCase(Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 10))) = "WEEE" Then
         chkWEEE.Value = 1
         chkNonWEEE.Value = 0
      ElseIf mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 10) = "N/A" Then
         chkWEEE.Value = 0
         chkNonWEEE.Value = 1
      End If
      If UCase(Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 11))) = "CHINA ROHS" Then
         chkChinaRoHS.Value = 1
         Else
         chkChinaRoHS.Value = 0
         End If
      If UCase(Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 12))) = "HUAWEI ROHS" Then
         chkRoHS.Value = 1
         chkNonRoHS.Value = 0
      ElseIf mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 12) = "N/A" Then
         chkRoHS.Value = 0
         chkNonRoHS.Value = 1
      End If
      
      txtMS.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 13)
      Me.dtpMSValidFrom.Value = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 14)
      txtNAL.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 15)
      dtpValidFrom.Value = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 16)
      
      If UCase(Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 17))) = "Y" Then
         chkSVPrint.Value = 1
         chkNoPrintSV.Value = 0
      Else
         chkSVPrint.Value = 0
         chkNoPrintSV.Value = 1
      End If
      
      txtHV.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 2)
            
    If UCase(Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 19))) = "TRUE" Then
        Me.chkPrintPC.Value = 1
    ElseIf UCase(Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 19))) = "FALSE" Then
        Me.chkNonPrintPC.Value = 1
    Else
        Me.chkPrintPC.Value = 0
        Me.chkNonPrintPC.Value = 0
    End If
    
    If Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 18)) = "Y" Then
        cb5000.ListIndex = 0
    ElseIf Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 18)) = "N" Then
        cb5000.ListIndex = 1
    ElseIf Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 18)) = "NA" Then
        cb5000.ListIndex = 2
    ElseIf Trim(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 18)) = "TBD" Then
        cb5000.ListIndex = 3
    Else
        cb5000.ListIndex = -1
    End If

      txtRemark.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 20)
   End If
End Sub

Private Sub mfgHUAWEI_SelChange()
   mfgHUAWEI_Click
End Sub


