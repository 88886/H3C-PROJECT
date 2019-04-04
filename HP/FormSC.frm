VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormSC 
   BackColor       =   &H00FF8080&
   Caption         =   "生产"
   ClientHeight    =   7065
   ClientLeft      =   5745
   ClientTop       =   2340
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   9510
   Begin VB.CommandButton Command7 
      Caption         =   "大华SN标签"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   15
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "海康H3C二维标签"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   14
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "海康SN二维标签"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   13
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "海康SN标签"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   12
      Top             =   4440
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "纯HP产品"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   8775
      Begin VB.CommandButton Command8 
         Caption         =   "HP SN MAC地址合并"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdSingle3 
         Caption         =   "HP SN标签14.6*7.7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6000
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdSingle1 
         Caption         =   "HP SN标签50*20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdSingle2 
         Caption         =   "HP SN标签48*6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "HP双标签产品"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   6015
      Begin VB.CommandButton cmdDouble2 
         Caption         =   "HP SN标签48*6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   6
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdDouble1 
         Caption         =   "HP SN标签50*20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "HP发货标签(竖)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HP发货标签"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "导入序列号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "标签打印选择(Label Printed Select)"
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
      TabIndex        =   2
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "FormSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDouble1_Click()
    FormHPSN.Show
End Sub

Private Sub cmdDouble2_Click()
    FormHPMSN.Show
End Sub

Private Sub Command1_Click()
    frmConsen7046.Show
End Sub

Private Sub cmdSingle1_Click()
    frmHP5020.Show
End Sub

Private Sub cmdSingle2_Click()
    FrmHP4806.Show
End Sub

Private Sub cmdSingle3_Click()
    frmHP14677.Show
End Sub

Private Sub Command2_Click()
    FormHPFahuo.Show
End Sub

Private Sub Command3_Click()
    frmHK2D14677.Show
End Sub

Private Sub Command4_Click()
Dim xlConn As New ADODB.Connection
Dim xlRs As New ADODB.Recordset
Dim sITEM_CODE As String
Dim sBARCODE As String
Dim strConn As String
Dim xlCnt As Integer
Dim strWO As String

   If conn1.State = 0 Then
      conn1.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
      conn1.Open
   End If
   
'If conn.State = 0 Then
'        conn.ConnectionString = Connect.getConnectionstring
'        conn.Open
'End If
sql = "delete from tblHP_Print"
conn1.Execute sql
conn1.Close
CommonDialog1.Filter = "Excel File (*.xls)|*.xls"
CommonDialog1.DefaultExt = "xls"
CommonDialog1.DialogTitle = "请选择你需要你导入的文件"
CommonDialog1.ShowOpen
strConn = CommonDialog1.Filename
If strConn = "" Then
    MsgBox ("请选择文件!")
    Exit Sub
End If

strWO = InputBox("请输入要导入的工单号", "导入提示", "")
If strWO = "" Then
    MsgBox ("没有输入工单!")
    Exit Sub
End If

Dim message As String, intResponse As Integer, intQuantity As Integer

strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strConn & ";Extended Properties='Excel 8.0;HDR=yes;IMEX=1'"
xlConn.Open strConn
'xlRs.Open "select ITEM_CODE,BARCODE from [sheet1$]", xlConn, adOpenStatic, adLockReadOnly
'WORKORDER   ITEM_CODE  QUANTITY
xlRs.Open "select ITEM_CODE,QUANTITY from [log$] where [WORKORDER] = '" & strWO & "'", xlConn, adOpenStatic, adLockReadOnly
xlCnt = xlRs.RecordCount
If xlCnt <= 0 Then
    MsgBox ("import.xls文件中无该工单的数据！")
    xlConn.Close
    Exit Sub
Else
    intQuantity = CInt(xlRs.Fields("QUANTITY"))
    message = "请确认当前工单:" + strWO + "所对应的ITEM是否为:" + xlRs.Fields("ITEM_CODE") + ",数量是否为:" + xlRs.Fields("QUANTITY")
    intResponse = MsgBox(message, vbYesNo + vbQuestion, "请确认工单信息是否正确")
    If intResponse = vbNo Then
       xlConn.Close
       Exit Sub
    End If
End If
If xlRs.State = 1 Then
    xlRs.Close
End If

xlRs.Open "select ITEM_CODE,BARCODE from [serials$] where MO_NO = '" & strWO & "'", xlConn, adOpenStatic, adLockReadOnly

xlCnt = xlRs.RecordCount
If xlCnt <> intQuantity Then
    MsgBox ("工单数量和条码数量不相等,请确认文件内容是否正确")
    Exit Sub
End If


For i = 1 To xlCnt
If IsNull(xlRs("ITEM_CODE")) <> True Then
    sITEM_CODE = xlRs("ITEM_CODE")
    sBARCODE = xlRs("BARCODE")
        'If conn.State = 0 Then
        '    conn.ConnectionString = Connect.getConnectionstring
        '    conn.Open
        'End If
       If conn1.State = 0 Then
          conn1.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
          conn1.Open
       End If
       
    sql = "insert tblHP_Print(ITEM_CODE,BARCODE) select '" & sITEM_CODE & "','" & sBARCODE & "' where not exists(select 1 from tblHP_Print where ITEM_CODE='" & sITEM_CODE & "' and BARCODE='" & sBARCODE & "')"
    conn1.Execute sql
    conn1.Close
    End If
    xlRs.MoveNext
Next
xlConn.Close
MsgBox ("序列号导入成功！")
End Sub

Private Sub Command5_Click()
    FormHPFahuoS.Show
End Sub

Private Sub Command6_Click()
    frmH3CHK2DPrint.Show
End Sub

Private Sub Command7_Click()
    frmDahuaZX.Show
End Sub

Private Sub Command8_Click()
    frmHPSNAndMac.Show
End Sub

'Private Sub Command9_Click()
'
' If Connect.AccessCheck(golUSERNAME, "offline") = False Then
'        MsgBox "您没有该界面的权限!"
'        Exit Sub
' Else
'         frmHPSNAndMacReprint.Show
' End If
'
'End Sub

