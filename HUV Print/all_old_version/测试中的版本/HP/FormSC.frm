VERSION 5.00
Begin VB.Form FormSC 
   BackColor       =   &H00FFFFFF&
   Caption         =   "����"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "HP������ǩ(��)"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "HP��Ʒ��ǩ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "HP������ǩ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "HPģ�����кű�ǩ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�������к�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ǩ��ӡѡ��(Label Printed Select)"
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
MsgBox ("import.xls�ļ��������ݣ�")
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
MsgBox ("���кŵ���ɹ���")
End If
End Sub

Private Sub Command5_Click()
FormHPFahuoS.Show
End Sub
