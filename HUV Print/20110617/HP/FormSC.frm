VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
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
   StartUpPosition =   2  '��Ļ����
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
Dim sITEM_CODE As String
Dim sBARCODE As String
Dim strConn As String
Dim xlCnt As Integer
Dim fileName As String
Dim workOrder As String



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

CommonDialog1.Filter = "Excel (*.xls)|*.xls|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen
fileName = CommonDialog1.fileName
If fileName = "" Then
    MsgBox "û��ѡ��Excel�ļ�"
    Exit Sub
Else
'    MsgBox fileName
End If
workOrder = InputBox("�����빤����", "��������")
If workOrder = "" Then
    MsgBox ("��������Ϊ��")
    Exit Sub
End If



Dim excelString As String
'queryString = "select ITEM_CODE,BARCODE from [sheet1$] where MO_NO = '"& workOrder&"'"
excelString = "select ITEM_CODE,BARCODE from [serials$] where MO_NO = '" & workOrder & "'"


strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fileName & ";Extended Properties='Excel 8.0;HDR=yes;IMEX=1'"
xlConn.Open strConn
xlRs.Open excelString, xlConn, adOpenStatic, adLockReadOnly
xlCnt = xlRs.RecordCount
If xlCnt <= 0 Then
    MsgBox ("import.xls�ļ��������ݣ�")
    xlConn.Close
    Exit Sub
End If
answer = MsgBox("��ȷ��Item�Ƿ�Ϊ" + xlRs("ITEM_CODE") + "?" + " ���������Ƿ�Ϊ" & Str(xlCnt) & "?", vbExclamation + vbYesNo, "ȷ�Ϲ�������")
If answer = vbNo Then
    MsgBox ("�ֹ�ȷ�ϲ���ȷ�˳�")
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
MsgBox ("���кŵ���ɹ���")
End Sub

Private Sub Command5_Click()
FormHPFahuoS.Show
End Sub
