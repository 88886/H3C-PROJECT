VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "װ���嵥��ӡ"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   Icon            =   "Packlist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10335
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtModel 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7080
      TabIndex        =   4
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtVersion 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7080
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txtSN 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7080
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   8535
      Left            =   0
      Picture         =   "Packlist.frx":0ECA
      ScaleHeight     =   8475
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�汾:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SN:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim con As ADODB.Connection
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim str As String
Dim com As ADODB.Command
Dim status As String


'Private Sub cmdDataUpdate_Click()
'    frmDataupdate.Show 1
'End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
'   If conn1.State = 0 Then
'        conn1.CursorLocation = adUseClient
'        conn1.ConnectionString = "Provider=SQLOLEDB;User ID=datasweep;PWD=datasweep;Initial Catalog=dsActive;Data Source=DS-DB"
'        conn1.ConnectionTimeout = 100
'        conn1.Open
'   End If
   If conn2.State = 0 Then
        conn2.CursorLocation = adUseClient
        conn2.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
        conn2.ConnectionTimeout = 100
        conn2.Open
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   If conn1.State = 1 Then
      conn1.Close
      Set conn1 = Nothing
   End If
   If conn2.State = 2 Then
    conn2.Close
    Set conn2 = Nothing
   End If
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If Len(txtSN.Text) < 10 Then
         MsgBox "��Ʒ��ų��Ȳ���С��10!"
         txtSN.SetFocus
         Exit Sub
        End If
      If Len(Replace(Trim(txtSN.Text), Chr(13) & Chr(10), "")) = 10 Then
            '==================
            Dim str As String
            str = " select top 1 part_number,part_revision,creation_time,order_number from (" & _
            "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
            "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
            "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
            Set rs4 = New ADODB.Recordset
            Set rs4.ActiveConnection = conn2
            rs4.Open str
            If rs4.EOF = True Then
                MsgBox "MESϵͳû�ж�Ӧ���ְ汾��Ϣ", vbOKOnly + vbExclamation, "����"
                txtSN.Text = ""
                txtSN.SetFocus
                rs3.Close
                rs4.Close
                Exit Sub
            Else
                Me.txtModel.Text = Mid(Trim(rs4.Fields(0)), 4, 8)
                txtVersion.Text = rs4.Fields(1)
            End If
            rs4.Close
      End If
      
      If Len(Replace(Trim(txtSN.Text), Chr(13) & Chr(10), "")) = 20 Then
        
        str = " select top 1 part_number,part_revision,creation_time,order_number from (" & _
        "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
        "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
        "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
        Set rs4 = New ADODB.Recordset
        Set rs4.ActiveConnection = conn2
        rs4.Open str
        If rs4.EOF = True Then
            MsgBox "MESû�ж�Ӧ���ְ汾��Ϣ", vbOKOnly + vbExclamation, "����"
            txtSN.Text = ""
            txtSN.SetFocus
            rs.Close
            rs4.Close
            Exit Sub
        Else
            Me.txtModel.Text = Mid(Trim(rs4.Fields(0)), 4, 8)
            txtVersion.Text = rs4.Fields(1)
        End If
        rs4.Close
        
        txtModel.Text = Mid(Trim(txtSN.Text), 3, 8)
            
            'Set con = New ADODB.Connection
            'con.CursorLocation = adUseClient
            'con.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
            'con.ConnectionTimeout = 100
            'If con.State = 1 Then
            '    con.Close
            'End If
   
            'con.Open
        
            'Dim rcSet As New ADODB.Recordset
            'rcSet.ActiveConnection = con
            'sql = "select top 1 * from tblPackList where model='" & Mid(Trim(txtSN.Text), 3, 8) & "' "
            'rcSet.Open sql, con, adOpenKeyset, adLockOptimistic
            'If rcSet.EOF Then
            '    rcSet.Close
            'Else
            '    txtVersion.Text = rcSet.Fields(1)
            'End If
            'If rcSet.State = 1 Then
            '    rcSet.Close
            'End If
            
      End If
      
        'Dim con7 As ADODB.Connection
        'Dim rs7 As ADODB.Recordset

        'Set con7 = New ADODB.Connection
        'Set rs7 = New ADODB.Recordset
        
        'con7.ConnectionString = "Provider=SQLOLEDB;User ID=sa;PWD=Itadmin1;Initial Catalog=Print;Data Source=sz-sql01"
        'con7.ConnectionTimeout = 50
        'con7.Open
        'Dim str7 As String
        'Set rs7.ActiveConnection = con7
        'rs7.CursorType = adOpenDynamic
        '    str7 = "select * from tblPackList where UseFlag='Y' and model='" & Trim(txtModel.Text) & "' and Version='" & Trim(txtVersion.Text) & "'"
        '    rs7.Open str7
        'If rs7.EOF = False Then
        '    If rs7.Fields("UseFlag") = "Y" Then
        '         MsgBox "װ���嵥�Ѿ�����", vbOKOnly + vbExclamation, "����"
        '        txtSN.Text = ""
        '        txtVersion.Text = ""
        '        txtModel.Text = ""
        '        txtSN.SetFocus
        '        rs7.Close
        '        Exit Sub
        '    End If
  
        'End If
        
        'If rs7.State = 1 Then
        '    rs7.Close
        'End If
      
      Dim spath As String
      spath = "\\10.11.1.25\Public\Manufacture\��ǩģ��\װ���嵥\" + Trim(txtModel.Text) + Trim(txtVersion.Text) + ".doc"
       'spath = "\\10.11.1.25\Public\Manufacture\��ǩģ��\װ���嵥����\" + Trim(txtModel.Text) + "_" + Trim(txtVersion.Text) + "����һ����.docx"

        Set fs = CreateObject("Scripting.FileSystemObject")
            If Not fs.FileExists(spath) Then
                MsgBox "û�ж�Ӧ���ִ�ӡģ��", vbOKOnly + vbExclamation, "����"
                txtSN.Text = ""
                txtVersion.Text = ""
                txtModel.Text = ""
                txtSN.SetFocus
                Exit Sub
            End If
            
        Dim wrdObject As Word.Application
        Dim wrdDoc As Word.Document
        Set wrdObject = CreateObject("Word.Application")
        Set wrdDoc = wrdObject.Documents.Open(spath)
        
        wrdObject.Visible = True
        'wrdObject.Selection.TypeText "This is some text."     'д�����ֵ�word�� 210231a86ph095000083
        wrdDoc.PrintOut
        'wrdDoc.ClosePrintPreview  '�رմ�ӡԤ��
        wrdDoc.Close          '�ر�word�ĵ�
        wrdObject.Quit   'wordӦ���˳�
        Set wrdDoc = Nothing '�ͷ��ڴ�
        Set wrdObject = Nothing '�ͷ��ڴ�
        txtSN.Text = ""
        txtVersion.Text = ""
        txtModel.Text = ""
        txtSN.SetFocus
    End If
End Sub
