VERSION 5.00
Begin VB.Form FormSC 
   BackColor       =   &H00FFFFFF&
   Caption         =   "生产"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   5700
   StartUpPosition =   2  '屏幕中心
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
      Left            =   1680
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HP产品标签"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   720
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HP模块序列号标签"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   3240
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
      Left            =   1680
      TabIndex        =   0
      Top             =   4080
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
      TabIndex        =   4
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "FormSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    FormHPSN.Show
End Sub

Private Sub Command2_Click()
    FormHPFahuo.Show
End Sub

Private Sub Command3_Click()
    FormHPMSN.Show
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
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\import.xls;Extended Properties='Excel 8.0;HDR=yes;IMEX=1'"
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
FormHPFahuoS.Show
End Sub
