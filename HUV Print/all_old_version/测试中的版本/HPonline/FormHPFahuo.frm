VERSION 5.00
Begin VB.Form FormHPFahuo 
   BackColor       =   &H80000009&
   Caption         =   "HP������ǩ"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   6945
      TabIndex        =   14
      Top             =   120
      Width           =   6975
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   0
         Picture         =   "FormHPFahuo.frx":0000
         ScaleHeight     =   2985
         ScaleWidth      =   6945
         TabIndex        =   15
         Top             =   0
         Width           =   6975
      End
   End
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
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
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
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      Top             =   6600
      Visible         =   0   'False
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
      TabIndex        =   11
      Top             =   6600
      Visible         =   0   'False
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
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   10
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtUPC 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5640
      Width           =   2895
   End
   Begin VB.TextBox txtProduct 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox txtPN 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4440
      Width           =   2895
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
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Caption         =   "��Ʒ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "��ƷUPC��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "��Ʒ��ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "��Ʒ���֣�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "��Ʒ���кţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
End
Attribute VB_Name = "FormHPFahuo"
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

Public Sub cmdMPrint_Click()



    'begin
            If Len(txtSN.Text) < 10 Then
            MsgBox "��Ʒ��ų��Ȳ���С��10!"
            txtSN.SetFocus
            Exit Sub
        End If
        sql = "select * from hp where charindex(hp_sn_iii,'" & Trim(txtSN.Text) & "')=5 and pack_label='Y'"
        'MsgBox (sql)
        rec.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rec.EOF = True Then
            MsgBox "�����к�δά����Ϣ!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
        Else
        
           If IsNull(rec.Fields("hp_pn")) Then                     'modified by Jimmy Sun 2010.06.17
           txtPN = ""
            '    MsgBox ("�����к�δά������!")
             '   rec.Close
              '  Exit Sub
            Else
                txtPN = rec.Fields("hp_pn")
            End If
             
            If IsNull(rec.Fields("hp_gtin_number")) Then
                MsgBox ("�����к�δά��UPC!")
                rec.Close
                Exit Sub
            Else
                txtUPC = rec.Fields("hp_gtin_number")
            End If
        
            If IsNull(rec.Fields("hp_product")) Then
                MsgBox ("�����к�δά����Ʒ����!")
                rec.Close
                Exit Sub
            Else
                txtProduct = rec.Fields("hp_product")
            End If
      
            If IsNull(rec.Fields("hp_desc1")) Then
                MsgBox ("�����к�δά��������Ϣ!")
                rec.Close
                Exit Sub
            Else
                txtDesc = rec.Fields("hp_desc1")
            End If
            
            If Not IsNull(rec.Fields("hp_desc2")) Then
                txtDesc = txtDesc & " " & rec.Fields("hp_desc2")
            End If
            rec.Close
            cmdPrint_HPSN_Click
        End If
        'FormHPFahuo.Hide
        frmH3CPrint.Show
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
Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub
Private Sub OpenLppx()
    Me.MousePointer = vbHourglass
    'Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HUAWEI-����.lab")
    'If Trim(txtPN.Text) <> "" Then
    '    Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HP������ǩ.lab")
    'Else
    '    Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HP������ǩ_NO_PN.lab")
    'End If
    Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HP������ǩNEW.lab")
    
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub

Private Sub cmdCancel_HPSN_Click()
txtSN.Text = ""
txtProduct.Text = ""
txtDesc.Text = ""
txtUPC.Text = ""
txtPN.Text = ""
txtSN.SetFocus

End Sub

Private Sub cmdPrint_HPSN_Click()

    If txtSN.Text = "" Then
        MsgBox ("���к�δ���룬���ܴ�ӡ��")
        txtSN.SetFocus
        Exit Sub
    End If
'    If txtPN.Text = "" Then
'        MsgBox ("����δ���������ܴ�ӡ��")
'        txtSN.SetFocus
'        Exit Sub
'    End If
    If txtProduct.Text = "" Then
        MsgBox ("��Ʒ����δ���������ܴ�ӡ��")
        Exit Sub
    End If
    If txtDesc.Text = "" Then
        MsgBox ("��Ʒ����δ���������ܴ�ӡ��")
        Exit Sub
    End If
    If txtUPC.Text = "" Then
        MsgBox ("��ƷUPCδ���������ܴ�ӡ��")
        Exit Sub
    End If
    OpenLppx
    'myVars.Item("ID").Value = UCase(txtDesc.Text)
     myVars.Item("ID").Value = txtDesc.Text
    'myVars.Item("SN1").Value = UCase(txtSN.Text)
    'myVars.Item("SN2").Value = "S" & UCase(txtSN.Text)
    myVars.Item("SN2").Value = UCase(txtSN.Text)
    'myVars.Item("PN1").Value = UCase(txtPN.Text)
    If Trim(txtPN.Text) <> "" Then
    myVars.Item("PN2").Value = UCase(txtPN.Text)
    End If
    'myVars.Item("Product1").Value = UCase(txtProduct.Text)
    myVars.Item("Product2").Value = UCase(txtProduct.Text)
    myVars.Item("UPC").Value = Left(Trim(txtUPC.Text), 11)
    

    myDoc.PrintLabel 1
    myDoc.FormFeed
    UnloadLppx
    cmdCancel_HPSN_Click
    
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If Len(txtSN.Text) < 10 Then
            MsgBox "��Ʒ��ų��Ȳ���С��10!"
            txtSN.SetFocus
            Exit Sub
        End If
        sql = "select * from hp where charindex(hp_sn_iii,'" & Trim(txtSN.Text) & "')<>0"
        'MsgBox (sql)
        rec.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rec.EOF = True Then
            MsgBox "�����к�δά����Ϣ!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
        Else
        
            If IsNull(rec.Fields("hp_pn")) Then
                MsgBox ("�����к�δά������!")
                rec.Close
                Exit Sub
            Else
                txtPN = rec.Fields("hp_pn")
            End If
             
            If IsNull(rec.Fields("hp_gtin_number")) Then
                MsgBox ("�����к�δά��UPC!")
                rec.Close
                Exit Sub
            Else
                txtUPC = rec.Fields("hp_gtin_number")
            End If
        
            If IsNull(rec.Fields("hp_product")) Then
                MsgBox ("�����к�δά����Ʒ����!")
                rec.Close
                Exit Sub
            Else
                txtProduct = rec.Fields("hp_product")
            End If
      
            If IsNull(rec.Fields("hp_desc1")) Then
                MsgBox ("�����к�δά��������Ϣ!")
                rec.Close
                Exit Sub
            Else
                txtDesc = rec.Fields("hp_desc1")
            End If
            
            If Not IsNull(rec.Fields("hp_desc2")) Then
                txtDesc = txtDesc & " " & rec.Fields("hp_desc2")
            End If
        
        End If
    End If
End Sub

Private Sub cmdReturn_HPSN_Click()
Unload Me
End Sub
