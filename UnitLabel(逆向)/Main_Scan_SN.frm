VERSION 5.00
Begin VB.Form Main_Scan_SN 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "单板出货标签打印-逆向"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   Icon            =   "Main_Scan_SN.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   11445
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdDeleteBoxID 
      Height          =   495
      Left            =   7920
      Picture         =   "Main_Scan_SN.frx":073E
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7800
      Width           =   2775
   End
   Begin VB.CommandButton cmdPrintBefore 
      Height          =   495
      Left            =   7920
      Picture         =   "Main_Scan_SN.frx":102F
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7080
      Width           =   2775
   End
   Begin VB.CommandButton cmdRePrint 
      Height          =   495
      Left            =   7920
      Picture         =   "Main_Scan_SN.frx":1907
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdDeleteAll 
      Height          =   495
      Left            =   7920
      Picture         =   "Main_Scan_SN.frx":21F0
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5640
      Width           =   2775
   End
   Begin VB.CommandButton cmdDeleteOne 
      Height          =   495
      Left            =   7920
      Picture         =   "Main_Scan_SN.frx":2ADB
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4920
      Width           =   2775
   End
   Begin VB.ListBox lstScanTmp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      ItemData        =   "Main_Scan_SN.frx":339A
      Left            =   7800
      List            =   "Main_Scan_SN.frx":339C
      TabIndex        =   31
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   6735
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   7575
      Begin VB.CheckBox CheckN4 
         BackColor       =   &H80000004&
         Caption         =   "N4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   47
         Top             =   6120
         Width           =   735
      End
      Begin VB.CheckBox CheckN3 
         BackColor       =   &H80000004&
         Caption         =   "N3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   46
         Top             =   6120
         Width           =   735
      End
      Begin VB.CheckBox CheckN2 
         BackColor       =   &H80000004&
         Caption         =   "N2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   45
         Top             =   6120
         Width           =   735
      End
      Begin VB.CheckBox CheckN1 
         BackColor       =   &H80000004&
         Caption         =   "N1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   44
         Top             =   6120
         Width           =   735
      End
      Begin VB.CheckBox CheckY1 
         BackColor       =   &H80000004&
         Caption         =   "Y1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   43
         Top             =   5640
         Width           =   735
      End
      Begin VB.CheckBox CheckY2 
         BackColor       =   &H80000004&
         Caption         =   "Y2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   42
         Top             =   5640
         Width           =   735
      End
      Begin VB.CheckBox CheckY3 
         BackColor       =   &H80000004&
         Caption         =   "Y3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   41
         Top             =   5640
         Width           =   735
      End
      Begin VB.CheckBox CheckY4 
         BackColor       =   &H80000004&
         Caption         =   "Y4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   40
         Top             =   5640
         Width           =   735
      End
      Begin VB.CheckBox CheckYx 
         BackColor       =   &H80000004&
         Caption         =   "Y*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   39
         Top             =   5640
         Width           =   735
      End
      Begin VB.CheckBox CheckNx 
         BackColor       =   &H80000004&
         Caption         =   "N*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   38
         Top             =   5640
         Width           =   735
      End
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H80000001&
         Caption         =   $"Main_Scan_SN.frx":339E
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   29
         Top             =   5160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkRoHS 
         BackColor       =   &H80000001&
         Caption         =   "RoHS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   5160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtQuality 
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
         Left            =   3240
         TabIndex        =   27
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtRev 
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
         Left            =   3240
         TabIndex        =   26
         Top             =   4440
         Width           =   3735
      End
      Begin VB.TextBox txtBoxID 
         BackColor       =   &H80000000&
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
         Left            =   3240
         TabIndex        =   25
         Top             =   3840
         Width           =   3735
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H80000000&
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
         Left            =   3240
         TabIndex        =   24
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtModel 
         BackColor       =   &H80000000&
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
         Left            =   3240
         TabIndex        =   23
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H80000000&
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
         Left            =   3240
         TabIndex        =   22
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox txtContpactNo 
         BackColor       =   &H80000000&
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
         Left            =   3240
         TabIndex        =   21
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtVenderName 
         BackColor       =   &H80000000&
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
         Left            =   3240
         TabIndex        =   20
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtVendorCode 
         BackColor       =   &H80000000&
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
         Left            =   3240
         TabIndex        =   19
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         Caption         =   "PCS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   37
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label lblRohs 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 环保属性(Rohs) :"
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
         Index           =   10
         Left            =   120
         TabIndex        =   18
         Top             =   5640
         Width           =   3015
      End
      Begin VB.Label lblQuality 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 内含数量(Quality) :"
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
         Index           =   9
         Left            =   120
         TabIndex        =   17
         Top             =   5040
         Width           =   3015
      End
      Begin VB.Label lblRev 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 条码版本(Rev) :"
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
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   4440
         Width           =   3015
      End
      Begin VB.Label lblBoxid 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 箱号(Box ID) :"
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
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 描述(Description) :"
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
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label lblModel 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 机种编码(Model) :"
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
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 日期(Date) :"
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
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label lblContpact 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 合同号(Contpact No) :"
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
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label lblVN 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 供应商名称(Vendor Name) :"
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
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblVC 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 供应商代码(Vendor Code) :"
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
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7575
      Begin VB.CommandButton cmdPrint 
         Height          =   855
         Left            =   6240
         Picture         =   "Main_Scan_SN.frx":33AC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtArea 
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
         Left            =   2760
         TabIndex        =   6
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtCodeNo 
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
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "PCS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblArea 
         BackColor       =   &H00FFFFFF&
         Caption         =   "扫描区域(Scan Area)"
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
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblCodeNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "条码个数(Bar Code No)"
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Label lblScaned 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 已扫描条码 :"
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
      Index           =   0
      Left            =   7800
      TabIndex        =   30
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   $"Main_Scan_SN.frx":3B23
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "Main_Scan_SN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim myApp As New LabelManager2.Application
Dim mydoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects

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

Private Sub cmdDeleteAll_Click()

    lstScanTmp.Clear
    
     If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
        
    sql = " Delete from tblUNIT_Tmp where FLAG=0 AND userid='" & golUSERNAME & "'"
    conn.Execute sql
    
End Sub

Private Sub cmdDeleteBoxID_Click()
    frmUnLock.Show
End Sub

Private Sub cmdDeleteOne_Click()
    
    If lstScanTmp.ListCount > 0 Then
    
        If conn.State = 0 Then
            conn.ConnectionString = Connect.getConnectionstring
            conn.Open
        End If
        
        sql = " Delete from tblUNIT_Tmp where sn='" & lstScanTmp.List(lstScanTmp.ListIndex) & "' and FLAG=0 AND userid='" & golUSERNAME & "' "
        conn.Execute sql
    
        lstScanTmp.RemoveItem lstScanTmp.ListIndex
    
    End If
    
End Sub

Private Sub cmdPrint_Click()

On Error GoTo errhandle

    If CheckY1.Value + CheckY2.Value + CheckY3.Value + CheckY4.Value + CheckYx.Value + CheckNx.Value + CheckN1.Value + CheckN2.Value + CheckN3.Value + CheckN4.Value <> 1 Then
        MsgBox "请选择一项环保属性!"
        Exit Sub
    End If
    
    If Trim(txtRev.Text) = "" Then
        MsgBox "条码版本不可为空！"
        txtRev.SetFocus
        Exit Sub
    End If
    
    If Trim(txtQuality.Text) = "" Then
        MsgBox "内含数量不可为空！"
        txtQuality.SetFocus
        Exit Sub
    End If
    
    'If chkRoHS.Value = 0 And chkNonRoHS.Value = 0 Then
    If CheckY1.Value = 0 And CheckY2.Value = 0 And CheckY3.Value = 0 And CheckY4.Value = 0 And CheckYx.Value = 0 And CheckNx.Value = 0 And CheckN1.Value = 0 And CheckN2.Value = 0 And CheckN3.Value = 0 And CheckN4.Value = 0 Then
        MsgBox "环保属性不可为空！"
        'chkRoHS.SetFocus
        Exit Sub
    End If
    
    If Trim(txtCodeNo.Text) = "" Then
        MsgBox "条码个数不可为空！"
        txtArea.SetFocus
        Exit Sub
    End If
    
    If lstScanTmp.ListCount < 10 And CInt(Me.txtCodeNo.Text) = 10 Then
        MsgBox "流水号不够10个!"
        txtArea.SetFocus
        Exit Sub
    End If
    
    If lstScanTmp.ListCount <> CInt(Me.txtCodeNo.Text) Then
        MsgBox "扫描数量与条码个数设定不相符!"
        txtArea.SetFocus
        Exit Sub
    End If
    
    
    OpenLppx
    
    Dim i As Integer
    Dim pb As String
    
    myVars.Item("Vendor Code").Value = txtVendorCode.Text
    myVars.Item("Vendor Name").Value = txtVenderName.Text
    myVars.Item("Contract No").Value = txtContpactNo.Text
    myVars.Item("Date").Value = txtDate.Text
    
    myVars.Item("Machine Kind").Value = txtModel.Text
    myVars.Item("Description").Value = txtDescription.Text
    myVars.Item("ver").Value = IIf(Trim(txtRev.Text) = "/", "", Trim(txtRev.Text))
    
    myVars.Item("Quantity").Value = Trim(txtQuality.Text)
    myVars.Item("Box ID").Value = txtBOXID.Text
    'myVars.Item("RoHS").Value = IIf(chkRoHS.Value = 1, "RoHS", "Non-RoHS")
    
    'Add by mike for PB information 2015.5.20
    If CheckY1.Value = 1 Then pb = "Y1"
    If CheckY2.Value = 1 Then pb = "Y2"
    If CheckY3.Value = 1 Then pb = "Y3"
    If CheckY4.Value = 1 Then pb = "Y4"
    If CheckYx.Value = 1 Then pb = "Y*"
    If CheckNx.Value = 1 Then pb = "N*"
    If CheckN1.Value = 1 Then pb = "N1"
    If CheckN2.Value = 1 Then pb = "N2"
    If CheckN3.Value = 1 Then pb = "N3"
    If CheckN4.Value = 1 Then pb = "N4"

    myVars.Item("RoHS").Value = pb

        Select Case lstScanTmp.ListCount
        
        Case 1
            myVars.Item("SN1").Value = lstScanTmp.List(0)
            
            myObjs("Barcode3").Top = 100000
            myObjs("Barcode11").Top = 100000
            myObjs("Barcode15").Top = 100000
            myObjs("Barcode16").Top = 100000
            myObjs("Barcode22").Top = 100000
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
            
        Case 2
            myVars.Item("SN1").Value = lstScanTmp.List(0)
            myVars.Item("SN2").Value = lstScanTmp.List(1)
            
            myObjs("Barcode11").Top = 100000
            myObjs("Barcode15").Top = 100000
            myObjs("Barcode16").Top = 100000
            myObjs("Barcode22").Top = 100000
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        
        Case 3
            myVars.Item("SN1").Value = lstScanTmp.List(0)
            myVars.Item("SN2").Value = lstScanTmp.List(1)
            myVars.Item("SN3").Value = lstScanTmp.List(2)
            
            myObjs("Barcode15").Top = 100000
            myObjs("Barcode16").Top = 100000
            myObjs("Barcode22").Top = 100000
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 4
            myVars.Item("SN1").Value = lstScanTmp.List(0)
            myVars.Item("SN2").Value = lstScanTmp.List(1)
            myVars.Item("SN3").Value = lstScanTmp.List(2)
            myVars.Item("SN4").Value = lstScanTmp.List(3)
            
            myObjs("Barcode16").Top = 100000
            myObjs("Barcode22").Top = 100000
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 5
            myVars.Item("SN1").Value = lstScanTmp.List(0)
            myVars.Item("SN2").Value = lstScanTmp.List(1)
            myVars.Item("SN3").Value = lstScanTmp.List(2)
            myVars.Item("SN4").Value = lstScanTmp.List(3)
            myVars.Item("SN5").Value = lstScanTmp.List(4)
            
            myObjs("Barcode22").Top = 100000
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 6
            myVars.Item("SN1").Value = lstScanTmp.List(0)
            myVars.Item("SN2").Value = lstScanTmp.List(1)
            myVars.Item("SN3").Value = lstScanTmp.List(2)
            myVars.Item("SN4").Value = lstScanTmp.List(3)
            myVars.Item("SN5").Value = lstScanTmp.List(4)
            myVars.Item("SN6").Value = lstScanTmp.List(5)
            
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 7
            myVars.Item("SN1").Value = lstScanTmp.List(0)
            myVars.Item("SN2").Value = lstScanTmp.List(1)
            myVars.Item("SN3").Value = lstScanTmp.List(2)
            myVars.Item("SN4").Value = lstScanTmp.List(3)
            myVars.Item("SN5").Value = lstScanTmp.List(4)
            myVars.Item("SN6").Value = lstScanTmp.List(5)
            myVars.Item("SN7").Value = lstScanTmp.List(6)
            
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 8
            myVars.Item("SN1").Value = lstScanTmp.List(0)
            myVars.Item("SN2").Value = lstScanTmp.List(1)
            myVars.Item("SN3").Value = lstScanTmp.List(2)
            myVars.Item("SN4").Value = lstScanTmp.List(3)
            myVars.Item("SN5").Value = lstScanTmp.List(4)
            myVars.Item("SN6").Value = lstScanTmp.List(5)
            myVars.Item("SN7").Value = lstScanTmp.List(6)
            myVars.Item("SN8").Value = lstScanTmp.List(7)
            
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 9
            myVars.Item("SN1").Value = lstScanTmp.List(0)
            myVars.Item("SN2").Value = lstScanTmp.List(1)
            myVars.Item("SN3").Value = lstScanTmp.List(2)
            myVars.Item("SN4").Value = lstScanTmp.List(3)
            myVars.Item("SN5").Value = lstScanTmp.List(4)
            myVars.Item("SN6").Value = lstScanTmp.List(5)
            myVars.Item("SN7").Value = lstScanTmp.List(6)
            myVars.Item("SN8").Value = lstScanTmp.List(7)
            myVars.Item("SN9").Value = lstScanTmp.List(8)
            
            myObjs("Barcode36").Top = 100000
        Case 10
            myVars.Item("SN1").Value = lstScanTmp.List(0)
            myVars.Item("SN2").Value = lstScanTmp.List(1)
            myVars.Item("SN3").Value = lstScanTmp.List(2)
            myVars.Item("SN4").Value = lstScanTmp.List(3)
            myVars.Item("SN5").Value = lstScanTmp.List(4)
            myVars.Item("SN6").Value = lstScanTmp.List(5)
            myVars.Item("SN7").Value = lstScanTmp.List(6)
            myVars.Item("SN8").Value = lstScanTmp.List(7)
            myVars.Item("SN9").Value = lstScanTmp.List(8)
            myVars.Item("SN10").Value = lstScanTmp.List(9)
        End Select
        
        If conn.State = 0 Then
            conn.ConnectionString = Connect.getConnectionstring
            conn.Open
        End If
        
    
    sql = "Delete from tblUNIT_Tmp WHERE boxid<>'" & txtBOXID.Text & "' and  userid='" & golUSERNAME & "'"
    conn.Execute sql
    
    'sql = "Update tblUNIT_Tmp set flag=1,Rev='" & IIf(txtRev.Text = "/", "", Trim(txtRev.Text)) & "',Quality=" & Trim(txtQuality.Text) & ",RoHS='" & IIf(chkRoHS.Value = 1, 1, 0) & "' where  boxid='" & txtBoxID.Text & "' AND userid='" & golUSERNAME & "'"
    sql = "Update tblUNIT_Tmp set flag=1,Rev='" & IIf(txtRev.Text = "/", "", Trim(txtRev.Text)) & "',Quality=" & Trim(txtQuality.Text) & ",RoHS='" & pb & "' where  boxid='" & txtBOXID.Text & "' AND userid='" & golUSERNAME & "'"
    conn.Execute sql
    
    sql = ""
    sql = "Insert tblUnit(SN, BarCodeNum, VendorCode, VendorName, ContpactNo, PDate, Model, Description, BoxID, Rev, Quality, RoHS, UserID, PrintTime) "
    sql = sql & " select  SN, BarCodeNum, VendorCode, VendorName, ContpactNo, PDate, Model, Description, BoxID, Rev, Quality, RoHS, UserID,getdate() from tblUNIT_Tmp where boxid='" & txtBOXID.Text & "' and userid='" & golUSERNAME & "'"
    status = Connect.excuteUpdate(sql)
    
    mydoc.PrintLabel 1
    mydoc.FormFeed
    
    UnloadLppx
   
    
    clearForm
    
    Exit Sub
errhandle:

    MsgBox Err.Description

End Sub

Private Sub cmdPrintBefore_Click()
On Error GoTo errhandle

    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    
    
    Dim j As Integer
    j = 0
    Dim recTmp2 As New ADODB.Recordset
    sql = "select count(*) from tblUNIT_Tmp  where flag=1 and userid='" & golUSERNAME & "'"
    recTmp2.Open sql, conn, adOpenKeyset, adLockOptimistic
    If recTmp2.EOF = False Then
        j = CInt(recTmp2.Fields(0))
    Else
        MsgBox "系统资料读取错误!"
        txtArea.SetFocus
        Exit Sub
    End If
    recTmp2.Close
    
    
    If j = 0 Then
        MsgBox "前一箱号资料已经被删除!"
        txtArea.SetFocus
        Exit Sub
    End If
    
    OpenLppx
    
    Dim iCount As Integer
    Dim recTmp As New ADODB.Recordset
    sql = "select distinct BarCodeNum, VendorCode, VendorName, ContpactNo, PDate, Model, Description, BoxID, Rev, Quality, RoHS from tblUNIT_Tmp where flag=1 and userid='" & golUSERNAME & "'"
    recTmp.Open sql, conn, adOpenKeyset, adLockOptimistic
    If recTmp.EOF = False Then
        iCount = recTmp.Fields(0)
        myVars.Item("Vendor Code").Value = recTmp.Fields(1)
        myVars.Item("Vendor Name").Value = recTmp.Fields(2)
        myVars.Item("Contract No").Value = recTmp.Fields(3)
        myVars.Item("Date").Value = recTmp.Fields(4)
        myVars.Item("Machine Kind").Value = recTmp.Fields(5)
        myVars.Item("Description").Value = recTmp.Fields(6)
        myVars.Item("ver").Value = IIf(recTmp.Fields(8) = "/", "", recTmp.Fields(8))
        myVars.Item("Quantity").Value = recTmp.Fields(9)
        myVars.Item("Box ID").Value = recTmp.Fields(7)
        'myVars.Item("RoHS").Value = IIf(CBool(recTmp.Fields(10)) = True, "RoHS", "Non-RoHS")
        myVars.Item("RoHS").Value = recTmp.Fields(10)
    Else
        MsgBox "前一箱号资料已经被删除!"
        txtArea.SetFocus
        Exit Sub
    End If
    recTmp.Close
    
    Dim arrSN(10) As String
    Dim k As Integer
    k = 0
    
    Dim recTmp3 As New ADODB.Recordset
    sql = "select sn from tblUNIT_Tmp where flag=1 and userid='" & golUSERNAME & "'"
    recTmp3.Open sql, conn, adOpenKeyset, adLockOptimistic
    Do While Not recTmp3.EOF
        arrSN(k) = recTmp3.Fields(0)
        k = k + 1
        recTmp3.MoveNext
    Loop
    recTmp3.Close
    
 
    Select Case j
        
        Case 1
            myVars.Item("SN1").Value = arrSN(0)
            
            myObjs("Barcode3").Top = 100000
            myObjs("Barcode11").Top = 100000
            myObjs("Barcode15").Top = 100000
            myObjs("Barcode16").Top = 100000
            myObjs("Barcode22").Top = 100000
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
            
        Case 2
            myVars.Item("SN1").Value = arrSN(0)
            myVars.Item("SN2").Value = arrSN(1)
            
            myObjs("Barcode11").Top = 100000
            myObjs("Barcode15").Top = 100000
            myObjs("Barcode16").Top = 100000
            myObjs("Barcode22").Top = 100000
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        
        Case 3
            myVars.Item("SN1").Value = arrSN(0)
            myVars.Item("SN2").Value = arrSN(1)
            myVars.Item("SN3").Value = arrSN(2)
            
            myObjs("Barcode15").Top = 100000
            myObjs("Barcode16").Top = 100000
            myObjs("Barcode22").Top = 100000
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 4
            myVars.Item("SN1").Value = arrSN(0)
            myVars.Item("SN2").Value = arrSN(1)
            myVars.Item("SN3").Value = arrSN(2)
            myVars.Item("SN4").Value = arrSN(3)
            
            myObjs("Barcode16").Top = 100000
            myObjs("Barcode22").Top = 100000
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 5
            myVars.Item("SN1").Value = arrSN(0)
            myVars.Item("SN2").Value = arrSN(1)
            myVars.Item("SN3").Value = arrSN(2)
            myVars.Item("SN4").Value = arrSN(3)
            myVars.Item("SN5").Value = arrSN(4)
            
            myObjs("Barcode22").Top = 100000
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 6
            myVars.Item("SN1").Value = arrSN(0)
            myVars.Item("SN2").Value = arrSN(1)
            myVars.Item("SN3").Value = arrSN(2)
            myVars.Item("SN4").Value = arrSN(3)
            myVars.Item("SN5").Value = arrSN(4)
            myVars.Item("SN6").Value = arrSN(5)
            
            myObjs("Barcode32").Top = 100000
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 7
            myVars.Item("SN1").Value = arrSN(0)
            myVars.Item("SN2").Value = arrSN(1)
            myVars.Item("SN3").Value = arrSN(2)
            myVars.Item("SN4").Value = arrSN(3)
            myVars.Item("SN5").Value = arrSN(4)
            myVars.Item("SN6").Value = arrSN(5)
            myVars.Item("SN7").Value = arrSN(6)
            
            myObjs("Barcode34").Top = 100000
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 8
            myVars.Item("SN1").Value = arrSN(0)
            myVars.Item("SN2").Value = arrSN(1)
            myVars.Item("SN3").Value = arrSN(2)
            myVars.Item("SN4").Value = arrSN(3)
            myVars.Item("SN5").Value = arrSN(4)
            myVars.Item("SN6").Value = arrSN(5)
            myVars.Item("SN7").Value = arrSN(6)
            myVars.Item("SN8").Value = arrSN(7)
            
            myObjs("Barcode35").Top = 100000
            myObjs("Barcode36").Top = 100000
        Case 9
            myVars.Item("SN1").Value = arrSN(0)
            myVars.Item("SN2").Value = arrSN(1)
            myVars.Item("SN3").Value = arrSN(2)
            myVars.Item("SN4").Value = arrSN(3)
            myVars.Item("SN5").Value = arrSN(4)
            myVars.Item("SN6").Value = arrSN(5)
            myVars.Item("SN7").Value = arrSN(6)
            myVars.Item("SN8").Value = arrSN(7)
            myVars.Item("SN9").Value = arrSN(8)
            
            myObjs("Barcode36").Top = 100000
        Case 10
            myVars.Item("SN1").Value = arrSN(0)
            myVars.Item("SN2").Value = arrSN(1)
            myVars.Item("SN3").Value = arrSN(2)
            myVars.Item("SN4").Value = arrSN(3)
            myVars.Item("SN5").Value = arrSN(4)
            myVars.Item("SN6").Value = arrSN(5)
            myVars.Item("SN7").Value = arrSN(6)
            myVars.Item("SN8").Value = arrSN(7)
            myVars.Item("SN9").Value = arrSN(8)
            myVars.Item("SN10").Value = arrSN(9)
        End Select
    
    mydoc.PrintLabel 1
    mydoc.FormFeed
    
    UnloadLppx
    
    Exit Sub
errhandle:

    MsgBox Err.Description

End Sub

Private Sub cmdRePrint_Click()

    txtRev.Text = ""
    txtQuality.Text = ""
    chkRoHS.Value = 1
    txtCodeNo.Text = "10"
    
    clearForm
    
End Sub

Private Sub Form_Load()
    Me.Show
    
    txtCodeNo.Text = "1"
    chkRoHS.Value = 1
    If UCase(golUSERNAME) = "ADMIN" Or UCase(golUSERNAME) = "SZ6897" Then
        cmdDeleteBoxID.Enabled = True
    Else
        cmdDeleteBoxID.Enabled = False
    End If
    
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
        
    txtArea.SetFocus
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    sql = "Delete from tblUNIT_Tmp where userid='" & golUSERNAME & "' "
    conn.Execute sql
    
    If conn.State = 1 Then
        conn.Close
        Set conn = Nothing
    End If
    
End Sub

Private Sub txtArea_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Dim pb As String
        If Len(txtArea.Text) < 10 Then
            MsgBox "产品序号长度不能小于10!"
            txtArea.Text = ""
            txtArea.SetFocus
            Exit Sub
        End If
        
        If Trim(Me.txtCodeNo.Text) = "" Then
            MsgBox "请先设定条码个数!"
            txtCodeNo.SetFocus
            Exit Sub
        End If
        
        If Trim(txtRev.Text) = "" Then
            MsgBox "扫描前请先设定条码版本！"
            txtArea.Text = ""
            txtRev.SetFocus
            Exit Sub
        End If
    
        If Trim(txtQuality.Text) = "" Then
            MsgBox "扫描前请先设定内含数量！"
            txtArea.Text = ""
            txtQuality.SetFocus
            Exit Sub
        End If
        
        If lstScanTmp.ListCount >= CInt(Me.txtCodeNo.Text) Then
            MsgBox "扫描数量超过条码个数设定!"
            txtArea.SetFocus
            Exit Sub
        End If
        
        txtVendorCode.Text = "141078"
        txtVenderName.Text = "飞旭电子(苏州)有限公司"
        txtContpactNo.Text = "0"
        txtDate.Text = Year(DateTime.Date) & "年" & Month(DateTime.Date) & "月" & Day(DateTime.Date) & "日"
            
            
        Dim modelno As String
        modelno = txtModel.Text
        
        Dim str As String
        str = Trim(txtArea.Text)
        If Left(str, 2) = "21" Then
            txtModel.Text = "HWFC" & UCase(Mid(str, 3, 8))
            'txtModel.Text = UCase(Mid(str, 3, 8))
        Else
            txtModel.Text = "HWFC03" & UCase(Mid(str, 1, 6))
            'txtModel.Text = "03" & UCase(Mid(str, 1, 6))
        End If
        
        If modelno <> "" And modelno <> Trim(txtModel.Text) Then
            txtModel.Text = ""
            MsgBox "必须扫描同一机种条码!"
            txtArea.SetFocus
            Exit Sub
        End If
        
        If conn.State = 0 Then
            conn.ConnectionString = Connect.getConnectionstring
            conn.Open
        End If
        
        Dim rsUnit As New ADODB.Recordset
        sql = "select * from tblUNIT where sn='" & Trim(txtArea.Text) & "'"
        rsUnit.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rsUnit.EOF = False Then
            MsgBox "此条码已经打印过,不可重复!"
            txtArea.Text = ""
            txtVendorCode.Text = ""
            txtVenderName.Text = ""
            txtContpactNo.Text = ""
            txtDate.Text = ""
            txtModel.Text = ""
            txtDescription.Text = ""
            txtArea.SetFocus
            Exit Sub
        End If
        rsUnit.Close
        
        Dim rctmp As New ADODB.Recordset
        sql = "Select * from tblUNIT_Tmp WHERE sn='" & Trim(txtArea.Text) & "' "
        rctmp.Open sql, conn, adOpenKeyset
        If rctmp.EOF = False Then
            MsgBox "不可重复扫描!"
            txtArea.Text = ""
            txtVendorCode.Text = ""
            txtVenderName.Text = ""
            txtContpactNo.Text = ""
            txtDate.Text = ""
            txtModel.Text = ""
            txtDescription.Text = ""
            txtArea.SetFocus
            Exit Sub
        End If
        rctmp.Close
        
        sql = ""
        sql = "select type from SingleUnit where sn='" & Mid(txtModel.Text, InStr(txtModel.Text, "0"), 8) & "'"
        rec.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rec.EOF = False Then
             txtDescription.Text = rec.Fields(0)
        Else
            MsgBox "此机种描述没有维护资料!"
            txtArea.Text = ""
            txtVendorCode.Text = ""
            txtVenderName.Text = ""
            txtContpactNo.Text = ""
            txtDate.Text = ""
            txtModel.Text = ""
            txtDescription.Text = ""
            txtArea.SetFocus
            Exit Sub
        End If
        rec.Close
        
        Dim strYY As String
        Dim strMM As String
        Dim strDD As String
        strYY = Right(CStr(Year(DateTime.Date)), 2)
        strMM = "0" & CStr(Month(DateTime.Date))
        strMM = Right(strMM, 2)
        strDD = "0" & CStr(Day(DateTime.Date))
        strDD = Right(strDD, 2)
        
        
        If txtBOXID.Text = "" Then
        
            Dim rcSet As New ADODB.Recordset
            sql = "select GetNum from tblUnitGetNum where DateDiff(Month,ActDate,GETDATE())=0 "
            rcSet.Open sql, conn, adOpenKeyset, adLockOptimistic
            If rcSet.EOF = False Then
            
                Dim i As Integer
                Dim sTmp As String
                i = CInt(rcSet.Fields(0))
                i = i + 1
                sTmp = "000" & CStr(i)
                sTmp = Right(sTmp, 4)
                txtBOXID.Text = "B" & strYY & strMM & sTmp
                
                sql = "update tblUnitGetNum set GetNum=isnull(GetNum,0)+1 where DateDiff(Month,ActDate,convert(datetime,getdate()))=0 "
                status = Connect.excuteUpdate(sql)
                
            Else
                
                Dim sTmp2 As String
                sTmp2 = "0001"
                txtBOXID.Text = "B" & strYY & strMM & sTmp2
                sql = "Insert into tblUnitGetNum(ActDate,GetNum) values (getdate(),1) "
                conn.Execute sql
                
            End If
            
            rcSet.Close
            
        End If
       
        If CheckY1.Value = 1 Then pb = "Y1"
        If CheckY2.Value = 1 Then pb = "Y2"
        If CheckY3.Value = 1 Then pb = "Y3"
        If CheckY4.Value = 1 Then pb = "Y4"
        If CheckYx.Value = 1 Then pb = "Y*"
        If CheckNx.Value = 1 Then pb = "N*"
        If CheckN1.Value = 1 Then pb = "N1"
        If CheckN2.Value = 1 Then pb = "N2"
        If CheckN3.Value = 1 Then pb = "N3"
        If CheckN4.Value = 1 Then pb = "N4"

        
        sql = "Insert into tblUNIT_Tmp values ('" & UCase(Trim(txtArea.Text)) & "','" & CInt(txtCodeNo.Text) & "','" & txtVendorCode.Text & "','" & txtVenderName.Text & "',"
        sql = sql & " '" & txtContpactNo.Text & "','" & txtDate.Text & "','" & UCase(txtModel.Text) & "','" & txtDescription.Text & "',"
        sql = sql & " '" & txtBOXID.Text & "','" & UCase(Trim(txtRev.Text)) & "','" & CInt(txtQuality.Text) & "',"
        'sql = sql & "'" & IIf(chkRoHS.Value = 1, 1, 0) & "','" & golUSERNAME & "',0)"
        sql = sql & "'" & pb & "','" & golUSERNAME & "',0)"
        conn.Execute sql
        
        lstScanTmp.AddItem UCase(Trim(Me.txtArea.Text))
        Me.txtArea.Text = ""
            
        If lstScanTmp.ListCount = 10 And CInt(Me.txtCodeNo.Text) = 10 Then
        
            cmdPrint_Click
            
            clearForm
        End If
        
        Me.txtArea.SetFocus
    End If
    
End Sub

Private Sub clearForm()
    Me.txtArea.Text = ""
    txtVendorCode.Text = ""
    txtVenderName.Text = ""
    txtContpactNo.Text = ""
    txtDate.Text = ""
    txtModel.Text = ""
    txtDescription.Text = ""
    txtBOXID.Text = ""
    
    lstScanTmp.Clear
End Sub

Private Sub txtCodeNo_KeyPress(KeyAscii As Integer)
    'If KeyAscii <> 8 Then
    '    If (KeyAscii > 57) Or (KeyAscii < 48) Then
    '        MsgBox "请输入小于10的数字!"
    '        txtCodeNo.SetFocus
    '        Exit Sub
    '    End If
    'End If
    
    If (KeyAscii = 13) Then
        txtArea.SetFocus
    End If
End Sub

Private Sub txtCodeNo_LostFocus()
    If Len(Trim(txtCodeNo.Text)) = 1 Then
        If txtCodeNo.Text = "0" Then
            txtCodeNo.Text = ""
            txtCodeNo.SetFocus
            Exit Sub
        End If
        
        e = Val(txtCodeNo.Text)
        If e = 0 Then
            k = MsgBox("此栏只能输入数字，不包含其他字符！ ", vbExclamation)
            txtCodeNo.Text = ""
            txtCodeNo.SetFocus
            Exit Sub
        End If
    ElseIf Len(Trim(txtCodeNo.Text)) = 2 Then
        Dim str As String
        str = txtCodeNo.Text
        
        If Left(str, 1) = "0" Then
            txtCodeNo.Text = Right(str, 1)
            
            If txtCodeNo.Text = "0" Then
                txtCodeNo.Text = ""
                txtCodeNo.SetFocus
                Exit Sub
            End If
        
            e = Val(txtCodeNo.Text)
            If e = 0 Then
                k = MsgBox("此栏只能输入数字，不包含其他字符！ ", vbExclamation)
                txtCodeNo.Text = ""
                txtCodeNo.SetFocus
                Exit Sub
            End If
            
        Else
            
            If Trim(txtCodeNo.Text) <> "10" Then
                k = MsgBox("此栏只能输入数字，并且不能大于10！ ", vbExclamation)
                txtCodeNo.Text = ""
                txtCodeNo.SetFocus
                Exit Sub
            End If
        End If
        
    End If
End Sub

Private Sub txtQuality_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        chkRoHS.SetFocus
    End If
    
    If KeyAscii <> 8 And KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRev_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        txtQuality.SetFocus
    End If
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set mydoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "单板出货标签.lab")
   Me.MousePointer = vbDefault
   Set myVars = mydoc.Variables
   Set myObjs = mydoc.DocObjects
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub txtRev_LostFocus()
    Me.txtRev.Text = UCase(Trim(txtRev.Text))
End Sub
