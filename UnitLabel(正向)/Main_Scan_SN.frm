VERSION 5.00
Begin VB.Form Main_Scan_SN 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "单板出货标签打印-正向"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "Main_Scan_SN.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11715
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      Height          =   6375
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   7695
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
         Top             =   5640
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
         TabIndex        =   38
         Top             =   5640
         Width           =   735
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
         TabIndex        =   26
         Top             =   240
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
         TabIndex        =   25
         Top             =   840
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
         TabIndex        =   24
         Top             =   1440
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
         TabIndex        =   23
         Top             =   2040
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
         TabIndex        =   22
         Top             =   2640
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
         TabIndex        =   21
         Top             =   3240
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
         TabIndex        =   20
         Top             =   3840
         Width           =   3735
      End
      Begin VB.TextBox txtRev 
         BackColor       =   &H80000003&
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
         Top             =   4440
         Width           =   3735
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
         TabIndex        =   18
         Top             =   5040
         Width           =   975
      End
      Begin VB.CheckBox chkRoHS 
         BackColor       =   &H80000004&
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
         TabIndex        =   17
         Top             =   4680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H80000004&
         Caption         =   $"Main_Scan_SN.frx":073E
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
         TabIndex        =   16
         Top             =   4680
         Visible         =   0   'False
         Width           =   1815
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
         TabIndex        =   37
         Top             =   240
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
         TabIndex        =   36
         Top             =   840
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
         TabIndex        =   35
         Top             =   1440
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
         TabIndex        =   34
         Top             =   2040
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
         TabIndex        =   33
         Top             =   2640
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
         TabIndex        =   32
         Top             =   3240
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
         TabIndex        =   31
         Top             =   3840
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
         TabIndex        =   30
         Top             =   4440
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
         TabIndex        =   29
         Top             =   5040
         Width           =   3015
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
         TabIndex        =   28
         Top             =   5640
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000004&
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
         TabIndex        =   27
         Top             =   5160
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdDeleteBoxID 
      Height          =   495
      Left            =   8160
      Picture         =   "Main_Scan_SN.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7800
      Width           =   2775
   End
   Begin VB.CommandButton cmdPrintBefore 
      Height          =   495
      Left            =   8160
      Picture         =   "Main_Scan_SN.frx":103D
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Width           =   2775
   End
   Begin VB.CommandButton cmdRePrint 
      Height          =   495
      Left            =   8160
      Picture         =   "Main_Scan_SN.frx":1915
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdDeleteAll 
      Height          =   495
      Left            =   8160
      Picture         =   "Main_Scan_SN.frx":21FE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   2775
   End
   Begin VB.CommandButton cmdDeleteOne 
      Height          =   495
      Left            =   8160
      Picture         =   "Main_Scan_SN.frx":2AE9
      Style           =   1  'Graphical
      TabIndex        =   10
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
      ItemData        =   "Main_Scan_SN.frx":33A8
      Left            =   8040
      List            =   "Main_Scan_SN.frx":33AA
      TabIndex        =   9
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7695
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
         MaxLength       =   3
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
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
      Left            =   8040
      TabIndex        =   8
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
Dim rsDS As New ADODB.Recordset
Dim rsFTPC As New ADODB.Recordset

Dim sql As String
Dim myApp As New LabelManager2.Application
Dim mydoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim s1 As String
Dim s2 As String

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
    
    setPB ("NONE")
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
        
        If lstScanTmp.ListCount = 0 Then
            setPB ("NONE")
        End If
    End If
    
End Sub

Private Sub cmdPrint_Click()

On Error GoTo errhandle

    If CheckY1.Value + CheckY2.Value + CheckY3.Value + CheckY4.Value + CheckYx.Value + CheckNx.Value <> 1 Then
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
    If CheckY1.Value = 0 And CheckY2.Value = 0 And CheckY3.Value = 0 And CheckY4.Value = 0 And CheckYx.Value = 0 And CheckNx.Value = 0 Then
        MsgBox "环保属性不可为空！"
        'chkRoHS.SetFocus
        Exit Sub
    End If

    If Trim(txtCodeNo.Text) = "" Then
        MsgBox "条码个数不可为空！"
        txtArea.SetFocus
        Exit Sub
    End If

   ' If lstScanTmp.ListCount < 72 And CInt(Me.txtCodeNo.Text) = 72 Then
   '     MsgBox "流水号不够72个!"
   '     txtArea.SetFocus
   '     Exit Sub
   ' End If

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
    
    'Add by mike for PB information 2015.5.19
    If CheckY1.Value = 1 Then pb = "Y1"
    If CheckY2.Value = 1 Then pb = "Y2"
    If CheckY3.Value = 1 Then pb = "Y3"
    If CheckY4.Value = 1 Then pb = "Y4"
    If CheckYx.Value = 1 Then pb = "Y*"
    If CheckNx.Value = 1 Then pb = "N*"
    myVars.Item("RoHS").Value = pb
    
    
        For m = 0 To lstScanTmp.ListCount - 1
         myVars.Item("SN" & CStr(m + 1)).Value = lstScanTmp.List(m)
        Next m
        
        For n = lstScanTmp.ListCount + 1 To 108
          myObjs("Barcode" & CStr(n)).Top = 100000
        Next n

        If conn.State = 0 Then
            conn.ConnectionString = Connect.getConnectionstring
            conn.Open
        End If
        
     
    sql = "Delete from tblUNIT_Tmp WHERE boxid<>'" & txtBOXID.Text & "' and  userid='" & golUSERNAME & "'"
    conn.Execute sql
    
    'sql = "Update tblUNIT_Tmp set flag=1,Rev='" & IIf(txtRev.Text = "/", "", Trim(txtRev.Text)) & "',Quality=" & Trim(txtQuality.Text) & ",RoHS='" & IIf(chkRoHS.Value = 1, 1, 0) & "' where  boxid='" & txtBOXID.Text & "' AND userid='" & golUSERNAME & "'"
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
    UnloadLppx
    MsgBox "打印异常，请重新打印，若多次发生请重开电脑"

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
    
 
   
        For m = 0 To j - 1
         myVars.Item("SN" & CStr(m + 1)).Value = arrSN(m)
        Next m
        
        For n = j + 1 To 108
          myObjs("Barcode" & CStr(n)).Top = 100000
        Next n

    
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
    txtCodeNo.Text = "108"
    
    clearForm
    
End Sub

Private Sub Form_Load()
    Me.Show
    
    txtCodeNo.Text = "108"
    chkRoHS.Value = 1
    If UCase(golUSERNAME) = "ADMIN" Or UCase(golUSERNAME) = "SZ6897" Or UCase(golUSERNAME) = "SZ10510" Then
        cmdDeleteBoxID.Enabled = True
    Else
        cmdDeleteBoxID.Enabled = False
    End If
    
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    
'   If conn1.State = 0 Then
'      conn1.ConnectionString = "Provider=SQLOLEDB;User ID=datasweep;PWD=datasweep;Initial Catalog=dsActive;Data Source=DS-DB"
'      conn1.Open
'   End If
    setPB ("NONE")
    txtArea.SetFocus
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    sql = "Delete from tblUNIT_Tmp where userid='" & golUSERNAME & "' "
    conn.Execute sql
    
    If conn.State = 1 Then
        conn.Close
        Set conn = Nothing
    End If
    
     If conn1.State = 1 Then
      conn1.Close
      Set conn1 = Nothing
   End If
   
   Call UnloadLppx
   
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
        
'        If Trim(txtRev.Text) = "" Then
'            MsgBox "扫描前请先设定条码版本！"
'            txtArea.Text = ""
'            txtRev.SetFocus
'            Exit Sub
'        End If
    
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
        Dim stringtmpmodel As String
        
        Dim str As String
        str = Trim(txtArea.Text)
'        If Trim(txtModel.Text) = "" Then
'            If Left(str, 2) = "21" Then
'                txtModel.Text = "HWFC" & UCase(Mid(str, 3, 8))
'                'txtModel.Text = UCase(Mid(str, 3, 8))
'                stringtmpmodel = UCase(Mid(str, 3, 8))
'            Else
'                txtModel.Text = "HWFC03" & UCase(Mid(str, 1, 6))
'                'txtModel.Text = "03" & UCase(Mid(str, 1, 6))
'                stringtmpmodel = "03" & UCase(Mid(str, 1, 6))
'            End If
'        Else
'            If Left(str, 2) = "21" Then
'                modelno = "HWFC" & UCase(Mid(str, 3, 8))
'            Else
'                modelno = "HWFC03" & UCase(Mid(str, 1, 6))
'            End If
'        End If
'
'
'        If modelno <> "" And modelno <> Trim(txtModel.Text) Then
'            txtModel.Text = ""
'            MsgBox "必须扫描同一机种条码!"
'            txtArea.SetFocus
'            Exit Sub
'        End If
        'comment by allen.yan for the accurate part number is not OK
        Dim conFTPC As ADODB.Connection
        Set conFTPC = New ADODB.Connection
        conFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
        conFTPC.Open

        sql = "select top 1 part_number , part_revision from unit with (nolock) where serial_number like '" & str & "%' order by creation_time desc"


        rsFTPC.Open sql, conFTPC, adOpenKeyset
        If rsFTPC.EOF = True Then
            rsFTPC.Close
            conFTPC.Close
            MsgBox "MES FTPC 中缺少此条码信息!"
            txtArea.Text = ""
            txtVendorCode.Text = ""
            txtVenderName.Text = ""
            txtContpactNo.Text = ""
            txtDate.Text = ""
            txtModel.Text = ""
            txtDescription.Text = ""
            txtArea.SetFocus
            Exit Sub
        Else
            'comment for temporarily not check
'            If Left(rsFTPC.Fields(0), 4) <> "HWFC" Then
'                rsFTPC.Close
'                conFTPC.Close
'                MsgBox "系统中改条码对应的机种不是单板阶HWFC开头，请确认！"
'                txtArea.Text = ""
'                txtVendorCode.Text = ""
'                txtVenderName.Text = ""
'                txtContpactNo.Text = ""
'                txtDate.Text = ""
'                txtModel.Text = ""
'                txtDescription.Text = ""
'                txtArea.SetFocus
'                Exit Sub
'            End If

            txtModel.Text = rsFTPC.Fields(0)
            
            If Left(rsFTPC.Fields(0), 4) <> "HWFC" And Left(rsFTPC.Fields(0), 3) = "HWF" Then
                txtModel.Text = Replace(rsFTPC.Fields(0).Value, "HWF", "HWFC")

            End If

            If Left(rsFTPC.Fields(0), 4) <> "HUVC" And Left(rsFTPC.Fields(0), 3) = "HUV" Then
                txtModel.Text = Replace(rsFTPC.Fields(0).Value, "HUV", "HUVC")

            End If
            
            
            If Trim(txtRev.Text) = "" Then
                txtRev.Text = rsFTPC.Fields(1)
            ElseIf Trim(txtRev.Text) <> rsFTPC.Fields(1) Then
                MsgBox "第一个单元版本" & txtRev.Text & "与当前版本" & rsFTPC.Fields(1) & "不一致"
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

        End If
        rsFTPC.Close
        
        
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
            txtRev.Text = ""
            txtArea.SetFocus
            Exit Sub
        End If
        rsUnit.Close
        
        Dim rctmp As New ADODB.Recordset
     sql = "Select * from tblUNIT_Tmp WHERE sn='" & Trim(txtArea.Text) & "'"
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
            txtRev.Text = ""
            txtArea.SetFocus
            Exit Sub
        End If
        rctmp.Close
        sql = "select top 1 a.op_name,b.status from tracked_object_history a with(NOLOCK) left join tracked_object_status b with(NOLOCK) on a.tobj_key=b.tobj_key where a.tobj_key in(select unit_key from unit with(NOLOCK) where serial_number = '" & Trim(txtArea.Text) & "') order by a.complete_time desc"

       
        rsFTPC.Open sql, conFTPC, adOpenKeyset
        If rsFTPC.EOF = True Then
            rsFTPC.Close
            conFTPC.Close
            MsgBox "DataSweep资料读取失败!"
            txtArea.Text = ""
            txtVendorCode.Text = ""
            txtVenderName.Text = ""
            txtContpactNo.Text = ""
            txtDate.Text = ""
            txtModel.Text = ""
            txtDescription.Text = ""
            txtRev.Text = ""
            txtArea.SetFocus
            Exit Sub
        Else
            s1 = rsFTPC.Fields(0)
            s2 = rsFTPC.Fields(1)
            If InStr(s1, "(FGI)") > 0 Then
                If s2 <> "Finished" And s2 <> "Closed" Then
                    MsgBox "FGI未Closed 或者 Finished!"
                    txtArea.Text = ""
                    txtVendorCode.Text = ""
                    txtVenderName.Text = ""
                    txtContpactNo.Text = ""
                    txtDate.Text = ""
                    txtModel.Text = ""
                    txtDescription.Text = ""
                    txtArea.SetFocus
                    txtRev.Text = ""
                    rsFTPC.Close
                    Exit Sub
                End If
            Else
                MsgBox "未扫FGI!"
                txtArea.Text = ""
                txtVendorCode.Text = ""
                txtVenderName.Text = ""
                txtContpactNo.Text = ""
                txtDate.Text = ""
                txtModel.Text = ""
                txtDescription.Text = ""
                txtArea.SetFocus
                txtRev.Text = ""
                rsFTPC.Close
                Exit Sub
            End If
        End If
        
        rsFTPC.Close
        
        '=========Add by mike get PB information 2015.5.19 start==========
        sql = "select pb from H3C_PB with(NOLOCK) WHERE serial_number ='" & Trim(txtArea.Text) & "' "
        rsFTPC.Open sql, conFTPC, adOpenKeyset
        If Not rsFTPC.EOF Then
            If CheckY1.Value = 0 And CheckY2.Value = 0 And CheckY3.Value = 0 And CheckY4.Value = 0 And CheckYx.Value = 0 And CheckNx.Value = 0 Then
                setPB (rsFTPC(0))
            ElseIf (CheckY1.Value = 1 And rsFTPC(0) <> "Y1") Or (CheckY2.Value = 1 And rsFTPC(0) <> "Y2") Or (CheckY3.Value = 1 And rsFTPC(0) <> "Y3") Or (CheckY4.Value = 1 And rsFTPC(0) <> "Y4") Or (CheckYx.Value = 1 And rsFTPC(0) <> "Y*") Or (CheckNx.Value = 1 And rsFTPC(0) = "N*") Then
                MsgBox "条码[" & Trim(txtArea.Text) & "]属性是[" & rsFTPC(0) & "],与当前选择不符!"
                rsFTPC.Close
                Exit Sub
            End If
        ElseIf Mid(Trim(txtArea.Text), Len(Trim(txtArea.Text)) - 6, 1) = "9" Then
            If CheckY1.Value = 0 And CheckY2.Value = 0 And CheckY3.Value = 0 And CheckY4.Value = 0 And CheckYx.Value = 0 And CheckNx.Value = 0 Then
                setPB ("Y2")
            ElseIf CheckY2.Value <> 1 Then
                MsgBox "条码[" & Trim(txtArea.Text) & "]属性是[Y2],与当前选择不符!"
                rsFTPC.Close
                Exit Sub
            End If
        Else
            If CheckY1.Value = 0 And CheckY2.Value = 0 And CheckY3.Value = 0 And CheckY4.Value = 0 And CheckYx.Value = 0 And CheckNx.Value = 0 Then
                MsgBox "无PB信息,请先手动选择!"
                setPB ("NONE")
                rsFTPC.Close
                Exit Sub
            End If
        End If
        rsFTPC.Close
        setPB ("LOCK")
        '=========Add by mike get PB information 2015.5.19 end  ==========
        
        sql = ""
        'modified by allen.yan for the model which is not started by HWFC
        '2014/10/15
        If rec.State = 1 Then rec.Close
        sql = "select type from SingleUnit where sn='" & Mid(txtModel.Text, InStr(txtModel.Text, "0")) & "'"
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
        
        sql = "Insert into tblUNIT_Tmp values ('" & UCase(Trim(txtArea.Text)) & "','" & CInt(txtCodeNo.Text) & "','" & txtVendorCode.Text & "','" & txtVenderName.Text & "',"
        sql = sql & " '" & txtContpactNo.Text & "','" & txtDate.Text & "','" & UCase(txtModel.Text) & "','" & txtDescription.Text & "',"
        sql = sql & " '" & txtBOXID.Text & "','" & UCase(Trim(txtRev.Text)) & "','" & CInt(txtQuality.Text) & "',"
        'sql = sql & "'" & IIf(chkRoHS.Value = 1, 1, 0) & "','" & golUSERNAME & "',0)"
        sql = sql & " '" & pb & "','" & golUSERNAME & "',0)"
        conn.Execute sql

        lstScanTmp.AddItem UCase(Trim(Me.txtArea.Text))
        Me.txtArea.Text = ""
            
        If lstScanTmp.ListCount = CInt(Me.txtCodeNo.Text) Then
        
            cmdPrint_Click
            
            clearForm
            setPB ("NONE")
            
        End If
        
        Me.txtArea.SetFocus
    End If
    
End Sub

Private Sub setPB(pb As String)
    If UCase(pb) <> "LOCK" Then
        CheckY1.Value = 0
        CheckY2.Value = 0
        CheckY3.Value = 0
        CheckY4.Value = 0
        CheckYx.Value = 0
        CheckNx.Value = 0
    End If
    CheckY1.Enabled = False
    CheckY2.Enabled = False
    CheckY3.Enabled = False
    CheckY4.Enabled = False
    CheckNx.Enabled = False
    CheckYx.Enabled = False
    
    If UCase(pb) = "LOCK" Then
        Exit Sub
    End If
    
    If UCase(pb) = "Y1" Then
        CheckY1.Value = 1
    ElseIf UCase(pb) = "Y2" Then
        CheckY2.Value = 1
    ElseIf UCase(pb) = "Y3" Then
        CheckY3.Value = 1
    ElseIf UCase(pb) = "Y4" Then
        CheckY4.Value = 1
    ElseIf UCase(pb) = "Y*" Then
        CheckYx.Value = 1
    ElseIf UCase(pb) = "N*" Then
        CheckNx.Value = 1
    ElseIf UCase(pb) = "NONE" Then
        CheckY1.Enabled = True
        'CheckY2.Enabled = True
        CheckY3.Enabled = True
        'CheckY4.Enabled = True
        CheckNx.Enabled = True
        CheckYx.Enabled = True
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
    txtRev.Text = ""
    
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
            
'        Else
'
'            If Trim(txtCodeNo.Text) <> "72" Then
'                k = MsgBox("此栏只能输入数字，并且不能大于72！ ", vbExclamation)
'                txtCodeNo.Text = ""
'                txtCodeNo.SetFocus
'                Exit Sub
'            End If
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
   Set mydoc = myApp.Documents.Open("\\10.11.1.25\Public\Manufacture\标签模板\" & "新108单板出货标签.Lab")
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
