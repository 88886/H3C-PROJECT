VERSION 5.00
Begin VB.Form FormHPMSN 
   BackColor       =   &H80000009&
   Caption         =   "HPģ�����кű�ǩ"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7125
   ScaleHeight     =   4305
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "������ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "�� ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "ȡ ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdReturn_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtSN 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtProduct 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2160
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      Picture         =   "FormHPMSN.frx":0000
      ScaleHeight     =   945
      ScaleWidth      =   7065
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "��Ʒ���кţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "��Ʒ��ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "FormHPMSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects

Private Sub cmdCancel_HPSN_Click()
txtSN.Text = ""
txtProduct.Text = ""
txtSN.SetFocus

End Sub

Private Sub cmdMPrint_Click()
cmdReturn_HPSN.Enabled = False
cmdPrint_HPSN.Enabled = False
cmdCancel_HPSN.Enabled = False
sql = "select sn from hp_print where isnull(sn,'')<>'' order by sn"
rs.Open sql, conn, adOpenStatic, adLockReadOnly
If rs.EOF = True Then
    MsgBox ("���к�δ���룡")
    rs.Close
    cmdReturn_HPSN.Enabled = True
    cmdPrint_HPSN.Enabled = True
    cmdCancel_HPSN.Enabled = True
    Exit Sub
Else
    For i = 1 To rs.RecordCount
    txtSN = rs("sn")
    'begin
        If Len(txtSN.Text) < 10 Then
            MsgBox "��Ʒ��ų��Ȳ���С��10!"
            txtSN.SetFocus
            Exit Sub
        End If
sql = "select * from hp where hp_sn_iii=substring('" & Trim(txtSN.Text) & "',5,3)"
        'MsgBox (sql)
        rec.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rec.EOF = True Then
            MsgBox "�����к�δά����Ϣ!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
        Else
            If IsNull(rec.Fields("hp_product")) Then
                MsgBox ("�����к�δά����Ʒ����!")
                rec.Close
                Exit Sub
            Else
                txtProduct = rec.Fields("hp_product")
            End If
            rec.Close
      
       cmdPrint_HPSN_Click
       
        End If
      'end
       rs.MoveNext
    Next
    rs.Close
End If
del_excel
del_sql
cmdReturn_HPSN.Enabled = True
cmdPrint_HPSN.Enabled = True
cmdCancel_HPSN.Enabled = True
End Sub
Sub del_sql()
Dim delsql As String
delsql = "delete from hp_print"
conn.Execute delsql
End Sub
Sub del_excel()
          Dim tempxlApp     As New Excel.Application
          Dim tempxlWorkbook     As New Excel.Workbook
          Dim tempxlSheet     As New Excel.Worksheet
          Set tempxlWorkbook = tempxlApp.Workbooks.Open(App.Path & "\import.xls")
          'tempxlApp.DisplayAlerts = False
          Set tempxlSheet = tempxlWorkbook.Worksheets("Sheet1")
          tempxlSheet.Select
          tempxlSheet.Cells.Select
          Selection.Delete Shift:=xlUp
          Cells(1, 1) = "SN"
          Set tempxlSheet = Nothing
          Set tempxlWorkbook = Nothing
          tempxlApp.Quit
          Set tempxlApp = Nothing

End Sub

Private Sub cmdPrint_HPSN_Click()
    If txtSN.Text = "" Then
        MsgBox ("���к�δ���룬���ܴ�ӡ��")
        txtSN.SetFocus
        Exit Sub
    End If
    If txtProduct.Text = "" Then
        MsgBox ("��Ʒ����δ���������ܴ�ӡ��")
        Exit Sub
    End If
    OpenLppx
    myVars.Item("SN").Value = UCase(txtSN.Text)
    myVars.Item("PN").Value = UCase(txtProduct.Text)
    myDoc.PrintLabel 1
    myDoc.FormFeed
    UnloadLppx
    cmdCancel_HPSN_Click
    
End Sub



Private Sub cmdReturn_HPSN_Click()
Unload Me
End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub





Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If Len(txtSN.Text) < 10 Then
            MsgBox "��Ʒ��ų��Ȳ���С��10!"
            txtSN.SetFocus
            Exit Sub
        End If
sql = "select * from hp where hp_sn_iii=substring('" & Trim(txtSN.Text) & "',5,3)"
        'MsgBox (sql)
        rec.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rec.EOF = True Then
            MsgBox "�����к�δά����Ϣ!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
        Else
            If IsNull(rec.Fields("hp_product")) Then
                MsgBox ("�����к�δά����Ʒ����!")
                rec.Close
                Exit Sub
            Else
                txtProduct = rec.Fields("hp_product")
            End If
      
       
        End If
    End If
End Sub
Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub
Private Sub OpenLppx()
    Me.MousePointer = vbHourglass
    'Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HUAWEI-����.lab")
    Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HPģ�����кű�ǩ.lab")
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub

