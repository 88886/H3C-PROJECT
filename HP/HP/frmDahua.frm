VERSION 5.00
Begin VB.Form frmDahuaZX 
   Caption         =   "��SN��ǩ"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11415
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox tbFirst 
      Height          =   270
      Left            =   9360
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtCustomerCode 
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
      Height          =   525
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   3600
      Width           =   3735
   End
   Begin VB.TextBox txtACSign 
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
      Height          =   525
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   4200
      Width           =   3735
   End
   Begin VB.TextBox txtDCSign 
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
      Height          =   525
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   4800
      Width           =   3735
   End
   Begin VB.TextBox txtRev 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   27
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkY2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Y2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   6360
      Width           =   735
   End
   Begin VB.CheckBox chkY 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Y*"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   24
      Top             =   6360
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   360
      Picture         =   "frmDahua.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   23
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtPart 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   21
      Top             =   600
      Width           =   3735
   End
   Begin VB.PictureBox Picture3 
      Height          =   615
      Left            =   360
      Picture         =   "frmDahua.frx":09F8
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   6360
      Width           =   615
   End
   Begin VB.CheckBox chkNonChinaRoHS 
      BackColor       =   &H0000C000&
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   5640
      Width           =   735
   End
   Begin VB.CheckBox chkNonWEEE 
      BackColor       =   &H0000C000&
      Caption         =   "��"
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
      Left            =   2400
      TabIndex        =   18
      Top             =   6360
      Width           =   735
   End
   Begin VB.CheckBox chkWEEE 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   6360
      Width           =   735
   End
   Begin VB.CheckBox chkChinaRoHS 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   5640
      Width           =   615
   End
   Begin VB.CheckBox chkCCC 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   5640
      Width           =   735
   End
   Begin VB.CheckBox chkNonCCC 
      BackColor       =   &H0000C000&
      Caption         =   "��"
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
      Left            =   5160
      TabIndex        =   14
      Top             =   5640
      Width           =   735
   End
   Begin VB.PictureBox Picture15 
      Height          =   495
      Left            =   3360
      Picture         =   "frmDahua.frx":238A
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox txtWorkOrder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtSN 
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
      Left            =   1440
      TabIndex        =   7
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox txtDesc 
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
      Left            =   1440
      TabIndex        =   6
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox txtModel 
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
      Left            =   1440
      TabIndex        =   5
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton cmdPrint_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "�� ӡ"
      Enabled         =   0   'False
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
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   6960
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
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   6960
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
      TabIndex        =   2
      Top             =   6960
      Width           =   1095
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
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   5220
      Left            =   5880
      Picture         =   "frmDahua.frx":2A90
      ScaleHeight     =   5160
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000004&
      Caption         =   "�ͻ����룺"
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
      Left            =   120
      TabIndex        =   34
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000004&
      Caption         =   "ֱ����Դ��"
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
      Left            =   120
      TabIndex        =   32
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000004&
      Caption         =   "������Դ��"
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
      Left            =   120
      TabIndex        =   31
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000004&
      Caption         =   "�汾:"
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
      Left            =   240
      TabIndex        =   28
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000004&
      Caption         =   "��������:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   26
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "����:"
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
      Left            =   240
      TabIndex        =   22
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "����:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "SN:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ���:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ�ͺ�:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
End
Attribute VB_Name = "frmDahuaZX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New Recordset
Dim bom_code As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim rs As New Recordset
Dim newLableFlag As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Sub chkY_Click()
   If chkY.Value = 1 Then
      chkY2.Value = 0
   Else
      chkY2.Value = 1
   End If
End Sub

Private Sub chkY2_Click()
   If chkY2.Value = 1 Then
      chkY.Value = 0
   Else
      chkY.Value = 1
   End If
End Sub

Private Sub cmdCancel_HPSN_Click()
    Me.txtWorkOrder.Text = ""
    Me.txtPart.Text = ""
    Me.txtRev.Text = ""
    Me.txtDesc.Text = ""
    Me.txtModel.Text = ""
    Me.txtCustomerCode.Text = ""
    Me.txtACSign.Text = ""
    Me.txtDCSign.Text = ""
    Me.chkChinaRoHS.Value = 0
    Me.chkNonChinaRoHS.Value = 0
    Me.chkCCC.Value = 0
    Me.chkNonCCC.Value = 0
    Me.chkWEEE.Value = 0
    Me.chkNonWEEE.Value = 0
    Me.txtSN.Text = ""
'    Me.txtOrder.Text = ""

    
    
End Sub

Private Sub cmdMPrint_Click()
    Dim model As String
    If Me.txtPart.Text = "" And txtModel.Text = "" Then
        MsgBox "�û�����Ϣ���ܴ�ӡ!"
        Exit Sub
    ElseIf Me.txtPart.Text <> "" Then
        model = Trim(txtPart.Text)
'    ElseIf txtModel.Text <> "" Then
'        model = Trim(txtModel.Text)
    End If
    chkY.Enabled = False
    chkY2.Enabled = False

sql = "select ITEM_CODE,BARCODE from tblHP_Print where isnull(BARCODE,'')<>'' and isnull(ITEM_CODE,'')<>'' order by BARCODE"
If conn1.State = 0 Then
    conn1.Open
End If
rs.Open sql, conn1, adOpenStatic, adLockReadOnly
If rs.EOF = True Then
    MsgBox ("���к�δ���룡")
    rs.Close
    cmdReturn_HPSN.Enabled = True
    'cmdPrint_HPSN.Enabled = True
    cmdCancel_HPSN.Enabled = True
    Exit Sub
Else
    For i = 1 To rs.RecordCount
    
        txtSN.Text = rs("BARCODE")
        'txtModel.Text = rs("ITEM_CODE")
        'begin
'        If Len(txtSN.Text) <> 9 Then
'            MsgBox "��Ʒ��ų��Ȳ�����9!"
'            Exit Sub
'        End If
        If Mid(Trim(txtPart.Text), 4, 8) <> Trim(rs("ITEM_CODE")) Then
            MsgBox ("�ù����Ϻź������Ӧ���ϺŲ�һ�£���ȷ�����빤���Ƿ���ȷ!")
            rs.Close
            Exit Sub
        End If
        
        cmdPrint_HPSN_Click
        rs.MoveNext
        If i Mod 100 = 0 Then
            Sleep (1000 * 10)
       End If
    Next
    UnloadLppx
    cmdCancel_HPSN_Click
    rs.Close
End If
'del_excel
del_sql
cmdReturn_HPSN.Enabled = True
'cmdPrint_HPSN.Enabled = True
cmdCancel_HPSN.Enabled = True
'MsgBox ("������ӡ�ɹ���")
End Sub

Private Sub cmdPrint_HPSN_Click()
    Dim Pb As String
    
    If txtSN.Text = "" Then
        MsgBox ("���к�δ���룬���ܴ�ӡ��")
'        txtSN.SetFocus
        Exit Sub
    End If

    If txtDesc.Text = "" Then
        MsgBox ("��Ʒ���δ���������ܴ�ӡ��")
        Exit Sub
    End If
    
    If txtModel.Text = "" Then
        MsgBox ("��Ʒ�ͺ�δ���������ܴ�ӡ��")
        Exit Sub
    End If
    
    If txtCustomerCode.Text = "" Then
        MsgBox ("�ͻ�����δ���������ܴ�ӡ��")
        Exit Sub
    End If
    
   
   If chkCCC.Value = 0 And chkNonCCC.Value = 0 Then
        MsgBox ("CCCδ���������ܴ�ӡ��")
        Exit Sub
   Else
        If chkCCC.Value = 1 Then
            CCC = "1"
        Else
            CCC = "0"
        End If
   End If
   
   If chkY2.Value = 0 And chkY.Value = 0 Then
        MsgBox ("SN���кŵ�Pbδ���������ܴ�ӡ��")
        Exit Sub
   Else
        If chkY2.Value = 1 Then
            Pb = "Y2"
        Else
            Pb = "Y*"
        End If
   End If
   
   If chkWEEE.Value = 0 And chkNonWEEE.Value = 0 Then
        MsgBox ("SN���кŵ�WEEEδ���������ܴ�ӡ��")
        Exit Sub
   Else
        If chkWEEE.Value = 1 Then
            WEEE = "1"
        Else
            WEEE = "0"
        End If
   End If
    
    OpenLppx
    myVars.Item("SN").Value = UCase(Trim(txtSN.Text)) + " " + UCase(Trim(txtRev.Text)) + " " + Pb
    myVars.Item("Product Name1").Value = UCase(Trim(txtDesc.Text))
    myVars.Item("Product ID").Value = UCase(Trim(txtModel.Text))
    myVars.Item("P/N").Value = UCase(Trim(txtCustomerCode.Text))
    
    If Trim(txtACSign.Text) <> "" Then
        myVars.Item("jiaoliu").Value = Trim(txtACSign.Text)
    Else
        myObjs("jiaoliu").Top = 10000
        myObjs("jiaoliu(1)").Top = 10000
    End If
    
    If Trim(txtDCSign.Text) <> "" Then
        myVars.Item("zhiliu").Value = Trim(txtDCSign.Text)
    Else
        myObjs("zhiliu").Top = 10000
        myObjs("zhiliu(1)").Top = 10000
    End If
    
    If Trim(txtACSign.Text) = "" And Trim(txtDCSign.Text) = "" Then
        myObjs("Text1(9)").Top = 10000
    End If

    If chkNonChinaRoHS.Value = 1 Then
      myObjs("China RoHS").Top = 10000
    End If
    
    If chkNonCCC.Value = 1 Then
      myObjs("CCC").Top = 10000
    End If
    
    If chkNonWEEE.Value = 1 Then
      myObjs("WEEE").Top = 10000
    End If
        
    If UploadH3CInfo(Pb, Trim(UCase(txtSN.Text)), Trim(UCase(txtRev.Text)), "NA", "N/A", "CHINA", "frmDahuZX") = False Then
        MsgBox "���ϱ���ʧ�ܲ��ܴ�ӡ!", vbInformation + vbOKOnly, "���ϱ���ʧ��"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
    If UploadH3C_PB(Pb, Trim(UCase(txtSN.Text)), Trim(UCase(txtRev.Text)), "NA", "N/A", "CHINA", "frmDahuZX") = False Then
        MsgBox "PB���ϱ���ʧ�ܲ��ܴ�ӡ!", vbInformation + vbOKOnly, "���ϱ���ʧ��"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If

'    If UploadConsen_Rec(Trim(UCase(txtSN.Text)), Trim(UCase(txtWorkOrder.Text)), Trim(UCase(txtPart.Text)), Trim(UCase(txtDesc.Text)), Trim(UCase(txtType.Text)), Trim(UCase(txtMaterial.Text)), Trim(UCase(txtPower.Text)), Trim(UCase(txtOrder.Text)), IIf(chkCCC.Value, "1", "0"), IIf(chkChinaRoHS.Value, "1", "0"), IIf(chkWEEE.Value, "1", "0"), IIf(chkLaser.Value, "1", "0"), "frmConsen7046") = False Then
'        MsgBox "Consen���ϱ���ʧ�ܲ��ܴ�ӡ!", vbInformation + vbOKOnly, "���ϱ���ʧ��"
''        txtSN.SetFocus
'        UnloadLppx
'        Exit Sub
'    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
    

    'OpenLppx
    myDoc.PrintLabel 1
    myDoc.FormFeed
End Sub

Private Sub OverridePb(ByRef Pb As String)
    Dim labelHistory As New Label_History
    Dim sn As String
    sn = txtSN.Text
    If labelHistory.Init(sn) Then
        Pb = labelHistory.Pb
    End If
End Sub

Private Sub cmdReturn_HPSN_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    If conn1.State = 0 Then
      conn1.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
      conn1.Open
   End If
   If connFTPC.State = 0 Then
        connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
        connFTPC.Open
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
   
   If connFTPC.State = 1 Then
        connFTPC.Close
        Set connFTPC = Nothing
   End If
End Sub



Private Function getDaHuaInformation() As Boolean

      sql = "select * from tblDaHuaNew where Part_Number='" & Mid(txtPart.Text, 4, 8) & "'"
      If rec.State = 1 Then
            rec.Close
      End If
      
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      
      If rec.EOF = True Then
          MsgBox "�����к�δά����Ϣ!"
          txtSN.Text = ""
          rec.Close
          getDaHuaInformation = False
          Exit Function
      Else
          If IsNull(rec.Fields("Part_Number")) Then
              MsgBox ("�����к�δά����Ʒ����!")
              rec.Close
              getDaHuaInformation = False
              Exit Function

          End If
    
          If IsNull(rec.Fields("Part_Desc")) Then
              MsgBox ("�����к�δά����Ʒ���!")
              rec.Close
              getDaHuaInformation = False
              Exit Function
          Else
              txtDesc.Text = rec.Fields("Part_Desc")
          End If
          
          If IsNull(rec.Fields("Part_Model")) Then
              MsgBox ("�����к�δά����Ʒ�ͺ�!")
              rec.Close
              getDaHuaInformation = False
              Exit Function
          Else
              txtModel.Text = rec.Fields("Part_Model")
          End If
    

          If IsNull(rec.Fields("CustomerCode")) Then
              MsgBox ("�����к�δά���ͻ�����!")
              rec.Close
              getDaHuaInformation = False
              Exit Function
          Else
              txtCustomerCode.Text = rec.Fields("CustomerCode")
          End If
          
          If IsNull(rec.Fields("ACSign")) Then
              txtACSign.Text = ""
          Else
              txtACSign.Text = rec.Fields("ACSign")
          End If
        
          If IsNull(rec.Fields("DCSign")) Then
              txtDCSign.Text = ""
          Else
              txtDCSign.Text = rec.Fields("DCSign")
          End If
          

            If UCase(Trim(rec.Fields("RoHS"))) = True Then
               chkChinaRoHS.Value = 1
               chkNonChinaRoHS.Value = 0
            Else
               chkChinaRoHS.Value = 0
               chkNonChinaRoHS.Value = 1
            End If
            
            If UCase(Trim(rec.Fields("WEEE"))) = True Then
               chkWEEE.Value = 1
               chkNonWEEE.Value = 0
            Else
               chkWEEE.Value = 0
               chkNonWEEE.Value = 1
            End If
            
            If UCase(Trim(rec.Fields("CCC"))) = True Then
               chkCCC.Value = 1
               chkNonCCC.Value = 0
            Else
               chkCCC.Value = 0
               chkNonCCC.Value = 1
            End If
      End If
      getDaHuaInformation = True
End Function

Private Sub txtWorkOrder_KeyPress(KeyAscii As Integer)
    Dim first As String
    Me.txtPart.Text = ""
    Me.txtRev.Text = ""
    Me.txtDesc.Text = ""
    Me.txtModel.Text = ""
    Me.txtCustomerCode.Text = ""
    Me.txtACSign.Text = ""
    Me.txtDCSign.Text = ""
    Me.chkChinaRoHS.Value = 0
    Me.chkNonChinaRoHS.Value = 0
    Me.chkCCC.Value = 0
    Me.chkNonCCC.Value = 0
    Me.chkWEEE.Value = 0
    Me.chkNonWEEE.Value = 0
    Me.txtSN.Text = ""
    chkY.Enabled = False
    chkY2.Enabled = False
'    Me.txtOrder.Text = ""
    
    If KeyAscii = 13 Then
        If Trim(Me.txtWorkOrder.Text) <> "" Then
            If rec.State = 1 Then
                rec.Close
            End If
           
            If connFTPC.State = 0 Then
                connFTPC.Open
            End If
            Dim tempWO As String
         
            tempWO = Trim(Me.txtWorkOrder.Text)
            If tempWO = "" Or tempWO = Null Then Return
            If UCase(tempWO) = "TASK" Then
                chkY.Enabled = True
                chkY2.Enabled = True
                Exit Sub
            End If
            first = ""
            tbFirst.Text = first
            sql = "select b.part_number,b.part_revision,c.order_type_S from WORK_ORDER a,WORK_ORDER_ITEMS b,UDA_Order c where a.order_key = b.order_key and c.object_key=a.order_key and a.order_number = '" & tempWO & "'"
            rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                MsgBox ("�ù��������ڣ���ȷ�����빤���Ƿ���ȷ!")
                rec.Close
                Exit Sub
            Else
                Me.txtPart.Text = rec.Fields("part_number")
                Me.txtRev.Text = rec.Fields("part_revision")
                If rec.Fields("order_type_S") = "PP05" Then
                    chkY.Enabled = True
                    chkY2.Enabled = True
                    rec.Close
                    'Exit Sub
                Else
                    Pb = Connect.getPbByPartList(tempWO, first)
                    Me.tbFirst.Text = first
                    Me.tbFirst.Text = "0"
                    If Pb = "Y2" Then
                        Me.chkY2.Enabled = False
                        Me.chkY2.Value = 1
                        Me.chkY.Enabled = False
    
                    ElseIf Pb = "Y*" Then
                        Me.chkY.Enabled = False
                        Me.chkY.Value = 1
                        Me.chkY2.Enabled = False
                    Else
                        MsgBox "Ǧ���Բ���Y2����Y*,���ܴ�ӡ����ȷ�ϸù������½׵�Ǧ�����Ƿ��趨"
                        Me.tbFirst.Text = ""
                        rec.Close
                        Exit Sub
                    End If
                End If
                OverridePb (Pb)
                
                If getDaHuaInformation = False Then
                    If connFTPC.State = 1 Then
                        connFTPC.Close
                        Exit Sub
                    End If
                    If rec.State = 1 Then
                        rec.Close
                        Exit Sub
                    End If
                End If

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
    Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\��ģ��\" & "�󻪱����ǩ.lab")
    
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub

Sub del_sql()
    Dim delsql As String
    delsql = "delete from tblHP_Print"
    conn1.Execute delsql
End Sub

