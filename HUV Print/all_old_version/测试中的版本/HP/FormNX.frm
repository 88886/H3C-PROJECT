VERSION 5.00
Begin VB.Form FormNX 
   BackColor       =   &H00FFFFFF&
   Caption         =   "逆向"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
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
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "FormNX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormHPFahuoNX.Show
End Sub

Private Sub Command4_Click()
Dim xlConn As New ADODB.Connection
Dim xlRs As New ADODB.Recordset
Dim sn As String
Dim strConn As String
Dim xlCnt As Integer
If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
sql = "delete from hp_print"
conn.Execute sql
conn.Close
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\import.xls;Extended Properties='Excel 8.0;HDR=yes;IMEX=1'"
xlConn.Open strConn
xlRs.Open "select SN from [sheet1$]", xlConn, adOpenStatic, adLockReadOnly
xlCnt = xlRs.RecordCount
If xlCnt <= 1 Then
MsgBox ("import.xls文件中无数据！")
xlConn.Close
Exit Sub
Else
For i = 1 To xlCnt
If IsNull(xlRs("sn")) <> True Then
sn = xlRs("sn")

    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
sql = "insert hp_print select '" & sn & "' where not exists(select 1 from hp_print where sn='" & sn & "')"
conn.Execute sql
conn.Close
End If
xlRs.MoveNext
Next
xlConn.Close
MsgBox ("序列号导入成功！")
End If
End Sub
