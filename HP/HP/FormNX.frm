VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormNX 
   BackColor       =   &H00FF8080&
   Caption         =   "逆向"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   10920
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "HPE发货标签(竖)"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   2640
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   3480
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
      Height          =   2175
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   6015
      Begin VB.CommandButton Btn_HP5020RMA 
         Caption         =   "HPE(5020)RMA"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   7
         Top             =   1440
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
         TabIndex        =   5
         Top             =   480
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
         TabIndex        =   4
         Top             =   480
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
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
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
      Left            =   840
      TabIndex        =   1
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HP发货标签"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "FormNX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_HP5020RMA_Click()
frmHP5020_RMA.Show
End Sub

Private Sub cmdSingle1_Click()
    frmHP5020NX.Show
End Sub

Private Sub cmdSingle2_Click()
    FrmHP4806NX.Show
End Sub

Private Sub cmdSingle3_Click()
    frmHP14677NX.Show
End Sub

Private Sub Command1_Click()
    FormHPFahuoNX.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command4_Click()
Dim xlConn As New ADODB.Connection
Dim xlRs As New ADODB.Recordset
Dim sITEM_CODE As String
Dim sBARCODE As String
Dim strConn As String
Dim xlCnt As Integer

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

strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strConn & ";Extended Properties='Excel 8.0;HDR=yes;IMEX=1'"
xlConn.Open strConn
xlRs.Open "select ITEM_CODE,BARCODE from [sheet1$]", xlConn, adOpenStatic, adLockReadOnly
xlCnt = xlRs.RecordCount
If xlCnt <= 0 Then
    MsgBox ("import.xls文件中无数据！")
    xlConn.Close
    Exit Sub
Else
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
End If
   
End Sub

Private Sub Command5_Click()
    FormHPFahuoSNX.Show
End Sub
