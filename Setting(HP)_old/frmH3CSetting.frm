VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmH3CSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H3C Setting(Test)"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmH3CSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   12120
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fmH3C 
      Height          =   5775
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12015
      Begin VB.TextBox txtSN 
         Height          =   450
         Left            =   2280
         TabIndex        =   38
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtCPN 
         Height          =   450
         Left            =   8880
         TabIndex        =   37
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtEPN 
         Height          =   450
         Left            =   2280
         TabIndex        =   36
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtDes 
         Height          =   450
         Left            =   8880
         TabIndex        =   35
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtOS 
         Height          =   450
         Left            =   2280
         TabIndex        =   34
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtGW 
         Height          =   450
         Left            =   8880
         TabIndex        =   33
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtNAL 
         Height          =   435
         Left            =   2280
         TabIndex        =   32
         Top             =   3840
         Width           =   2895
      End
      Begin VB.CheckBox chkCE 
         Caption         =   "CE"
         Height          =   375
         Left            =   2280
         TabIndex        =   31
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox chkNonCE 
         Caption         =   "�� CE"
         Height          =   375
         Left            =   3360
         TabIndex        =   30
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox chkWEEE 
         Caption         =   "��"
         Height          =   375
         Left            =   2400
         TabIndex        =   29
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CheckBox chkNonWEEE 
         Caption         =   "��"
         Height          =   375
         Left            =   3960
         TabIndex        =   28
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtMS 
         Height          =   450
         Left            =   8880
         TabIndex        =   27
         Top             =   3840
         Width           =   2775
      End
      Begin VB.TextBox txtHV 
         Height          =   495
         Left            =   2280
         TabIndex        =   25
         Top             =   5040
         Width           =   2895
      End
      Begin VB.TextBox txtRemark 
         Height          =   495
         Left            =   7440
         TabIndex        =   24
         Top             =   5040
         Width           =   4335
      End
      Begin VB.CheckBox chkChinaRoHS 
         Caption         =   "��"
         Height          =   375
         Left            =   8880
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         Caption         =   "��"
         Height          =   375
         Left            =   10080
         TabIndex        =   22
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox chkTurkeyRohs 
         Caption         =   "��"
         Height          =   375
         Left            =   8880
         TabIndex        =   21
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkNonTurkeyRohs 
         Caption         =   "��"
         Height          =   375
         Left            =   10080
         TabIndex        =   20
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CheckBox chkSVPrint 
         Caption         =   "��"
         Height          =   495
         Left            =   8880
         TabIndex        =   18
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chkNoPrintSV 
         Caption         =   "��"
         Height          =   495
         Left            =   10080
         TabIndex        =   17
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CheckBox chkRoHS 
         Caption         =   "��"
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   3240
         Width           =   975
      End
      Begin VB.CheckBox chkNonRoHS 
         Caption         =   "��"
         Height          =   495
         Left            =   3960
         TabIndex        =   15
         Top             =   3240
         Width           =   855
      End
      Begin VB.CheckBox chkCEAddr 
         Caption         =   "CE Addr"
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   2040
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpMSValidFrom 
         Height          =   495
         Left            =   8880
         TabIndex        =   19
         Top             =   4440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         Format          =   89784321
         CurrentDate     =   40425
      End
      Begin MSComCtl2.DTPicker dtpValidFrom 
         Height          =   495
         Left            =   2280
         TabIndex        =   26
         Top             =   4440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   89784321
         CurrentDate     =   39757
      End
      Begin VB.Label lblSN 
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblCPN 
         Caption         =   "��Ʒ����(����):"
         Height          =   375
         Left            =   6480
         TabIndex        =   55
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         Caption         =   "��Ʒ����(Ӣ��):"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblDes 
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   6480
         TabIndex        =   53
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblGW 
         Caption         =   "ë��(KG):"
         Height          =   375
         Left            =   6480
         TabIndex        =   52
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblValidFrom 
         Caption         =   "������Ч��:"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label lblCE 
         Caption         =   "��֤��ϢCE:"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblWEEE 
         Caption         =   "��֤��ϢWEEE:"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblRoHS 
         Caption         =   "��֤��ϢRoHS:"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblMS 
         Caption         =   "�����׼:"
         Height          =   375
         Left            =   6240
         TabIndex        =   47
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label lblHV 
         Caption         =   "Ӳ���汾:"
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label lblRemark 
         Caption         =   "��ע:"
         Height          =   495
         Left            =   6240
         TabIndex        =   45
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label lblNAL 
         Caption         =   "������ɺ�:"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label lblOS 
         Caption         =   "��ߴ�(MM):"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lbTurkeyRohs 
         Caption         =   "Turkey RoHS:"
         Height          =   495
         Left            =   6480
         TabIndex        =   42
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblChinaRoHS 
         Caption         =   "China RoHS:"
         Height          =   375
         Left            =   6480
         TabIndex        =   41
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblMSValidFrom 
         Caption         =   "�����׼��Ч��:"
         Height          =   495
         Left            =   6240
         TabIndex        =   40
         Top             =   4440
         Width           =   2775
      End
      Begin VB.Label lblPrintSV 
         Caption         =   "�Ƿ��ӡ����汾:"
         Height          =   375
         Left            =   6240
         TabIndex        =   39
         Top             =   3240
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog cdSelect 
      Left            =   2760
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   735
      Left            =   10800
      TabIndex        =   12
      Top             =   9720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   735
      Left            =   9360
      TabIndex        =   11
      Top             =   9720
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "ȷ��(Confirm)"
      Height          =   735
      Left            =   7680
      TabIndex        =   10
      Top             =   9720
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(Delete)"
      Height          =   735
      Left            =   10800
      TabIndex        =   9
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�޸�(Update)"
      Height          =   735
      Left            =   9360
      TabIndex        =   8
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "����(Insert)"
      Height          =   735
      Left            =   7680
      TabIndex        =   7
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "��ѯ(Query)"
      Height          =   735
      Left            =   6000
      TabIndex        =   6
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "����(Export)"
      Height          =   735
      Left            =   3480
      TabIndex        =   5
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "����(Import)"
      Height          =   735
      Left            =   3480
      TabIndex        =   4
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ѡ��(Select)"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   9960
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   9360
      Width           =   3015
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgH3C 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4895
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblPath 
      Caption         =   "����/����·��:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   8880
      Width           =   2175
   End
End
Attribute VB_Name = "frmH3CSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim op As String
Dim xlApp As New Excel.Application
Dim xlBook As New Excel.Workbook
Dim xlSheet As New Excel.Worksheet
Dim query As Boolean

Private Sub enable()
   txtSN.Enabled = True
   txtSN.BackColor = &HFFFFFF
   txtCPN.Enabled = True
   txtCPN.BackColor = &HFFFFFF
   txtEPN.Enabled = True
   txtEPN.BackColor = &HFFFFFF
   txtDes.Enabled = True
   txtDes.BackColor = &HFFFFFF
   txtOS.Enabled = True
   txtOS.BackColor = &HFFFFFF
   txtGW.Enabled = True
   txtGW.BackColor = &HFFFFFF
   
   chkCE.Enabled = True
   chkCEAddr.Enabled = True
   chkNonCE.Enabled = True
   chkWEEE.Enabled = True
   chkNonWEEE.Enabled = True
   chkChinaROHS.Enabled = True
   chkNonChinaROHS.Enabled = True
   chkTurkeyROHS.Enabled = True
   chkNonTurkeyROHS.Enabled = True
   
   chkRoHS.Enabled = True
   chkNonRoHS.Enabled = True
   'optH3CRoHS.Enabled = True
   'opt3COMRoHS.Enabled = True
   'optNonRoHS.Enabled = True
   
   
   txtMS.Enabled = True
   txtMS.BackColor = &HFFFFFF
   dtpMSValidFrom.Enabled = True
   txtNAL.Enabled = True
   txtNAL.BackColor = &HFFFFFF
   dtpValidFrom.Enabled = True
   
   chkSVPrint.Enabled = True
   chkNoPrintSV.Enabled = True
   
   txtHV.Enabled = True
   txtHV.BackColor = &HFFFFFF
   txtRemark.Enabled = True
   txtRemark.BackColor = &HFFFFFF
   
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = False
   cmdUpdate.Enabled = False
   cmdDelete.Enabled = False
   cmdConfirm.Enabled = True
   cmdCancel.Enabled = True
End Sub

Private Sub unable()
   txtSN.Enabled = False
   txtSN.BackColor = &HE0E0E0
   txtCPN.Enabled = False
   txtCPN.BackColor = &HE0E0E0
   txtEPN.Enabled = False
   txtEPN.BackColor = &HE0E0E0
   txtDes.Enabled = False
   txtDes.BackColor = &HE0E0E0
   txtOS.Enabled = False
   txtOS.BackColor = &HE0E0E0
   txtGW.Enabled = False
   txtGW.BackColor = &HE0E0E0
   
   chkCE.Enabled = False
   chkCEAddr.Enabled = False
   chkNonCE.Enabled = False
   chkWEEE.Enabled = False
   chkNonWEEE.Enabled = False
   chkChinaROHS.Enabled = False
   chkNonChinaROHS.Enabled = False
   chkTurkeyROHS.Enabled = False
   chkNonTurkeyROHS.Enabled = False
   'optH3CRoHS.Enabled = False
   'opt3COMRoHS.Enabled = False
   'optNonRoHS.Enabled = False
   chkRoHS.Enabled = False
   chkNonRoHS.Enabled = False
   
   txtMS.Enabled = False
   txtMS.BackColor = &HE0E0E0
   dtpMSValidFrom.Enabled = False
   txtNAL.Enabled = False
   txtNAL.BackColor = &HE0E0E0
   dtpValidFrom.Enabled = False
   chkSVPrint.Enabled = False
   chkNoPrintSV.Enabled = False
   
   txtHV.Enabled = False
   txtHV.BackColor = &HE0E0E0
   txtRemark.Enabled = False
   txtRemark.BackColor = &HE0E0E0
   
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = True
   cmdUpdate.Enabled = True
   cmdDelete.Enabled = True
   cmdConfirm.Enabled = False
   cmdCancel.Enabled = False
End Sub

Private Sub chkCE_Click()
   If chkCE.Value = 1 Then
      chkNonCE.Value = 0
   Else
      chkNonCE.Value = 1
   End If
End Sub

Private Sub chkCEAddr_Click()
   If chkCEAddr.Value = 1 Then
      chkCE.Value = 1
      chkNonCE.Value = 0
   End If
End Sub

Private Sub chkNonCE_Click()
   If chkNonCE.Value = 1 Then
      chkCE.Value = 0
      chkCEAddr.Value = 0
   Else
      chkCE.Value = 1
   End If
End Sub

Private Sub chkChinaROHS_Click()
   If chkChinaROHS.Value = 1 Then
      chkNonChinaROHS.Value = 0
   Else
      chkNonChinaROHS.Value = 1
   End If
End Sub

Private Sub chkNonChinaRoHS_Click()
   If chkNonChinaROHS.Value = 1 Then
      chkChinaROHS.Value = 0
   Else
      chkChinaROHS.Value = 1
   End If
End Sub

Private Sub chkNoPrintSV_Click()
   If chkNoPrintSV.Value = 1 Then
      chkSVPrint.Value = 0
   Else
      chkSVPrint.Value = 1
   End If
End Sub

Private Sub chkRoHS_Click()
    If chkRoHS.Value = 1 Then
        chkNonRoHS.Value = 0
    Else
        chkNonRoHS.Value = 1
    End If
End Sub

Private Sub chkNonRoHS_Click()
    If chkNonRoHS.Value = 1 Then
        chkRoHS.Value = 0
    Else
        chkRoHS.Value = 1
    End If
End Sub

Private Sub chkSVPrint_Click()
   If chkSVPrint.Value = 1 Then
      chkNoPrintSV.Value = 0
   Else
      chkNoPrintSV.Value = 1
   End If
End Sub

Private Sub chkWEEE_Click()
   If chkWEEE.Value = 1 Then
      chkNonWEEE.Value = 0
   Else
      chkNonWEEE.Value = 1
   End If
End Sub

Private Sub chkNonWEEE_Click()
   If chkNonWEEE.Value = 1 Then
      chkWEEE.Value = 0
   Else
      chkWEEE.Value = 1
   End If
End Sub

Private Sub chkTurkeyROHS_Click()
    If chkTurkeyROHS.Value = 1 Then
        chkNonTurkeyROHS.Value = 0
    Else
        chkNonTurkeyROHS.Value = 1
    End If
End Sub

Private Sub chkNonTurkeyRohs_Click()
    If chkNonTurkeyROHS.Value = 1 Then
        chkTurkeyROHS.Value = 0
    Else
        chkTurkeyROHS.Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
   unable
   op = ""
End Sub

Private Sub cmdConfirm_Click()
   If Trim(txtSN.Text) = "" Then
      MsgBox "��Ʒ���벻��Ϊ��!!", vbExclamation + vbOKOnly, "��Ʒ�����"
      txtSN.SetFocus
      Exit Sub
   End If
   If txtCPN.Text = "" Then
       MsgBox "��Ʒ����(����)����Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ����(����)��"
       txtCPN.SetFocus
       Exit Sub
   End If
   If txtEPN.Text = "" Then
      MsgBox "��Ʒ����(Ӣ��)����Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ����(Ӣ��)��"
      txtEPN.SetFocus
      Exit Sub
   End If
   If txtDes.Text = "" Then
      MsgBox "��Ʒ��������Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ������"
      txtDes.SetFocus
      Exit Sub
   End If
   If txtOS.Text = "" Then
      MsgBox "��ߴ粻��Ϊ��!", vbExclamation + vbOKOnly, "��ߴ��"
      txtOS.SetFocus
      Exit Sub
   End If
   If txtOS.Text = "/" Then
      MsgBox "����ߴ���ά��N/A!", vbExclamation + vbOKOnly, "����ߴ�"
      txtOS.SetFocus
      Exit Sub
   End If
   If txtOS.Text = "n/a" Then
      txtOS.Text = UCase(txtOS.Text)
   End If

   If txtOS.Text <> "N/A" Then
   
        txtOS.Text = LTrim(RTrim(txtOS.Text))
      
        If Right(txtOS.Text, 2) <> "mm" Then
            MsgBox "��ߴ��ʽ����!", vbExclamation + vbOKOnly, "��ߴ����"
            txtOS.SetFocus
            Exit Sub
        End If
        
        If InStr(txtOS.Text, "mmm") > 0 Then
            MsgBox "��ߴ��ʽ����!", vbExclamation + vbOKOnly, "��ߴ����"
            txtOS.SetFocus
            Exit Sub
        End If
   End If
   
   
   
   'If txtGW.Text = "" Then
   '   MsgBox "ë�ز���Ϊ��!", vbExclamation + vbOKOnly, "ë�ؿ�"
   '   txtGW.SetFocus
   '   Exit Sub
   'End If
    If Trim(txtGW.Text) <> "" Then
        If UCase(Right(Trim(txtGW.Text), 2)) <> "KG" Then
           MsgBox "ë�ر�����ϵ�λkg!", vbExclamation + vbOKOnly, "ë�ص�λ��"
           txtGW.SetFocus
           Exit Sub
        End If
   End If
   If txtMS.Text = "" Then
      MsgBox "�����׼����Ϊ��!", vbExclamation + vbOKOnly, "�����׼��"
      txtMS.SetFocus
      Exit Sub
   End If
   If txtNAL.Text = "" Then
      MsgBox "������ɺŲ���Ϊ��!", vbExclamation + vbOKOnly, "������ɺſ�"
      txtNAL.SetFocus
      Exit Sub
   End If
   If txtHV.Text = "" Then
      MsgBox "Ӳ���汾����Ϊ��!", vbExclamation + vbOKOnly, "Ӳ���汾��"
      txtHV.SetFocus
      Exit Sub
   End If
   If chkSVPrint.Value = 0 And chkNoPrintSV.Value = 0 Then
      MsgBox "�Ƿ��ӡ����汾����Ϊ��!", vbExclamation + vbOKOnly, "������汾��"
      txtHV.SetFocus
      Exit Sub
   End If
   
   Dim CE, WEEE, ChinaRoHS, RoHS, TurkeyRoHS, SVPrint As String
   If chkCE.Value = 1 Then
'      CE = "CE"
      If chkCEAddr.Value = 1 Then
        CE = "CE+CE Addr"
      Else
        CE = "CE"
      End If
   ElseIf chkNonCE.Value = 1 Then
      CE = "N/A"
   End If
   If chkWEEE.Value = 1 Then
      WEEE = "WEEE"
   ElseIf chkNonWEEE.Value = 1 Then
      WEEE = "N/A"
   End If
   If chkChinaROHS.Value = 1 Then
      ChinaRoHS = "China RoHS"
   ElseIf chkNonChinaROHS.Value = 1 Then
      ChinaRoHS = "N/A"
   End If
   If chkTurkeyROHS.Value = 1 Then
      TurkeyRoHS = "Turkey RoHS"
   ElseIf chkNonTurkeyROHS.Value = 1 Then
      TurkeyRoHS = "N/A"
   End If
   If chkSVPrint.Value = 1 Then
      SVPrint = "Y"
   ElseIf chkNoPrintSV.Value = 1 Then
      SVPrint = "N"
   End If
   
   If chkRoHS.Value = 1 Then
      RoHS = "H3C RoHS"
   ElseIf chkNonRoHS.Value = 1 Then
      RoHS = "N/A"
   End If
   
  ' If optH3CRoHS.Value = True Then
  '    RoHS = "H3C RoHS"
  ' ElseIf opt3COMRoHS.Value = True Then
  '    RoHS = "3COM RoHS"
  ' ElseIf optNonRoHS.Value = True Then
  '    RoHS = "/"
  ' End If
  
  txtGW.Text = LCase(Trim(txtGW.Text))
   
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from tblH3C where SN='" & Trim(txtSN.Text) & "' and HV='" & Trim(txtHV.Text) & "' "
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "��Ʒ����&�汾�Ѵ���!", vbExclamation + vbOKOnly, "��Ʒ����ظ�"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close

      sql = "Insert into tblH3C(ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark) " & _
            "Values(" & getmaxID("tblH3C") & ",'" & Trim(txtHV.Text) & "','" & Trim(txtSN.Text) & "','" & Trim(txtCPN.Text) & "','" & Trim(txtEPN.Text) & "','" & Trim(txtDes.Text) & "','" & Trim(txtOS.Text) & "','" & Trim(txtGW.Text) & "','" & CE & "','" & WEEE & "','" & ChinaRoHS & "','" & RoHS & "','" & TurkeyRoHS & "'," & _
            "'" & txtMS.Text & "','" & dtpMSValidFrom.Value & "','" & txtNAL.Text & "','" & dtpValidFrom.Value & "','" & SVPrint & "','" & txtRemark.Text & "')"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "����H3C�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "����ʧ��"
      End If
      MsgBox "����H3C�趨���ϳɹ�!", vbInformation + vbOKOnly, "�����ɹ�"
      renovate
      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update tblH3C set CPN='" & Trim(txtCPN.Text) & "',EPN='" & Trim(txtEPN.Text) & "',Des='" & Trim(txtDes.Text) & "',OS='" & Trim(txtOS.Text) & "',GW='" & Trim(txtGW.Text) & "',CE='" & CE & "',WEEE='" & WEEE & "',ChinaRoHS='" & ChinaRoHS & "',RoHS='" & RoHS & "',TurkeyRohs='" & TurkeyRoHS & "'," & _
            "MS='" & txtMS.Text & "',MSValidFrom='" & dtpMSValidFrom.Value & "',NAL='" & txtNAL.Text & "',ValidFrom='" & dtpValidFrom.Value & "',HV='" & txtHV.Text & "',PrintSV='" & SVPrint & "',Remark='" & txtRemark.Text & "'" & _
            " where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and SN='" & Trim(txtSN.Text) & "'"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "�޸�H3C�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "�޸�ʧ��"
      End If
      MsgBox "�޸�H3C�趨���ϳɹ�!", vbInformation + vbOKOnly, "�޸ĳɹ�"
      renovate
      cmdCancel_Click
   End If
   renovate
End Sub

Private Sub cmdDelete_Click()
   If mfgH3C.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫɾ������!", vbInformation + vbOKOnly, "δѡ����"
      Exit Sub
   End If
   sql = "delete from tblH3C where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and SN='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 3) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "ɾ��H3C�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "ɾ��ʧ��"
   End If
   MsgBox "ɾ��H3C�趨���ϳɹ�!", vbInformation + vbOKOnly, "ɾ���ɹ�"
   renovate
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If mfgH3C.Rows = 0 Then
      MsgBox "�����Ͽɻ��", vbExclamation + vbOKOnly, "������"
      Exit Sub
   End If
   If txtPath.Text <> "" Then
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets.Item(1)
       For i = 0 To mfgH3C.Rows - 1
         For j = 1 To mfgH3C.Cols - 1
          xlSheet.Cells(i + 1, j) = mfgH3C.TextMatrix(i, j)
       Next j
      Next i
      xlBook.SaveAs (txtPath.Text)
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "�����EXCEL���ϳɹ�!!", vbInformation + vbOKOnly, "����ɹ�"
    End If
End Sub

Private Sub cmdImport_Click()
   If txtPath.Text = "" Then
      MsgBox "����·������Ϊ��!", vbExclamation + vbOKOnly, "����·����"
      Exit Sub
   End If
   Dim action As Integer
   Dim info As Boolean
   info = True
   Set xlBook = xlApp.Workbooks.Open(txtPath.Text)
      For i = 1 To xlBook.Sheets.Count
       Set xlSheet = xlBook.Sheets.Item(i)
       For j = 2 To xlSheet.Rows.Count
        r = xlSheet.Cells(j, 1)
        If r = "" Then
           Exit For
        Else
          Dim cellValue As String
          Dim cellhvValue As String
          
          Dim isexist As Boolean
          If xlSheet.Cells(j, 19) = "" Then
             MsgBox "�������ϸ�ʽ����ȷ!", vbExclamation + vbOKOnly, "��ʽ����"
             Exit Sub
          End If
          If Not ((xlSheet.Cells(j, 18) = "N") Or (xlSheet.Cells(j, 18) = "Y")) Then
             MsgBox "�������ϸ�ʽ����ȷ!", vbExclamation + vbOKOnly, "��ʽ����"
             Exit Sub
          End If
          isexist = False
          For K = 1 To 19
          '======================================================
           If K = 3 Then
             cellValue = xlSheet.Cells(j, K)
             cellhvValue = xlSheet.Cells(j, 2)
             
             If cellValue = "" Or cellhvValue = "" Then
                MsgBox "�������ϸ�ʽ����ȷ!", vbExclamation + vbOKOnly, "��ʽ����"
                Exit Sub
             End If
             
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from tblH3C where SN='" & cellValue & "' and HV='" & cellhvValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("��Ʒ����&�汾�Ѵ���!", vbAbortRetryIgnore + vbExclamation, "�����ظ�")
                End If
                
                If action = vbAbort Then
                   MsgBox "���ϵ�������ֹ!!", vbInformation + vbOKOnly, "������ֹ"
                   rcd.Close
                   Exit Sub
                ElseIf action = vbIgnore And info = True Then
                   MsgBox "�ظ���Ʒ������ϲ��ᵼ��,���Ե�..!!", vbInformation + vbOKOnly, "�ظ����ᵼ��"
                   rcd.Close
                   info = False
                   Exit For
                ElseIf action = vbRetry And info = True Then
                   MsgBox "�ظ���Ʒ������ϻ��Զ�����,���Ե�..!!", vbInformation + vbOKOnly, "�ظ����Զ�����"
                   info = False
                End If
                isexist = True
             Else
                isexist = False
             End If
             rcd.Close
            End If
            '==================================================
            
            If K = 19 Then
               If action = vbRetry Then
                   sql = "Update tblH3C set CPN='" & xlSheet.Cells(j, 4) & "',EPN='" & xlSheet.Cells(j, 5) & "',Des='" & xlSheet.Cells(j, 6) & "',OS='" & xlSheet.Cells(j, 7) & "',GW='" & xlSheet.Cells(j, 8) & "',CE='" & xlSheet.Cells(j, 9) & "',WEEE='" & xlSheet.Cells(j, 10) & "',ChinaRoHS='" & xlSheet.Cells(j, 11) & "'," & _
                        "RoHS='" & xlSheet.Cells(j, 12) & "',TurkeyRohs='" & xlSheet.Cells(j, 13) & "',MS='" & xlSheet.Cells(j, 14) & "',MSValidFrom='" & xlSheet.Cells(j, 15) & "',NAL='" & xlSheet.Cells(j, 16) & "',ValidFrom='" & xlSheet.Cells(j, 17) & "',PrintSV='" & xlSheet.Cells(j, 18) & "',Remark='" & xlSheet.Cells(j, 19) & "'" & _
                        " where SN='" & xlSheet.Cells(j, 3) & "' and HV='" & xlSheet.Cells(j, 2) & "' "
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "�޸�H3C�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "�޸�ʧ��"
                   End If
'                   MsgBox "�޸�H3C�趨���ϳɹ�!"
               ElseIf isexist = False Then
                   sql = "Insert into tblH3C(ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark) " & _
                        " Values(" & getmaxID("tblH3C") & ",'" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "','" & xlSheet.Cells(j, 4) & "','" & xlSheet.Cells(j, 5) & "','" & xlSheet.Cells(j, 6) & "','" & xlSheet.Cells(j, 7) & "','" & xlSheet.Cells(j, 8) & "','" & xlSheet.Cells(j, 9) & "','" & xlSheet.Cells(j, 10) & "','" & xlSheet.Cells(j, 11) & "'," & _
                        "'" & xlSheet.Cells(j, 12) & "','" & xlSheet.Cells(j, 13) & "','" & xlSheet.Cells(j, 14) & "','" & xlSheet.Cells(j, 15) & "','" & xlSheet.Cells(j, 16) & "','" & xlSheet.Cells(j, 17) & "','" & xlSheet.Cells(j, 18) & "','" & xlSheet.Cells(j, 19) & "')"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "����H3C�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "�޸�ʧ��"
                   End If
'                   MsgBox "����H3C�趨���ϳɹ�!"
               End If
           End If
         Next K
         
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "H3C�趨���ϵ���ɹ�!"
      renovate
End Sub

Private Sub cmdInsert_Click()
   enable
   txtSN.Text = ""
   txtCPN.Text = ""
   txtEPN.Text = ""
   txtDes.Text = ""
   txtOS.Text = ""
   txtGW.Text = ""
   
   chkCE.Value = 1
   chkWEEE.Value = 1
   chkChinaROHS.Value = 1
   chkRoHS.Value = 1
   chkTurkeyROHS.Value = 1
   chkSVPrint.Value = 1
   
   txtNAL.Text = "N/A"
   dtpValidFrom.Value = Date
   
   txtMS.Text = "N/A"
   dtpMSValidFrom.Value = Date
   
   txtHV.Text = "N/A"
   txtRemark.Text = "N/A"
   op = "Insert"
End Sub

Private Sub cmdQuery_Click()
    If txtSN.Enabled = False Then
      MsgBox "�밴������ť��վͿ������ѯ����!", vbOKOnly + vbInformation, "�����ѯ����"
    End If
    If rec.State = 1 Then
        rec.Close
     End If
     sql = "select * from tblH3C Where 1=1"
     If txtSN.Text <> "" Then
        sql = sql & " and SN like '%" & txtSN.Text & "%'"
     End If
     If txtCPN.Text <> "" Then
        sql = sql & " and CPN like '%" & txtCPN.Text & "%'"
     End If
     If txtEPN.Text <> "" Then
        sql = sql & " and EPN='%" & txtEPN.Text & "%'"
     End If
     If txtDes.Text <> "" Then
        sql = sql & " and Des like '%" & txtDes.Text & "%'"
     End If
     If txtOS.Text <> "" Then
        sql = sql & " and OS like '%" & txtOS.Text & "%'"
     End If
     If txtGW.Text <> "" Then
        sql = sql & " and GW like '%" & txtGW.Text & "%'"
     End If
'     Dim CE, WEEE, ChinaRoHS, RoHS As String
'     If chkCE.Value = 1 Then
'        CE = "CE"
'     ElseIf chkNonCE.Value = 1 Then
'        CE = "/"
'     End If
'     If chkWEEE.Value = 1 Then
'        WEEE = "WEEE"
'     ElseIf chkNonWEEE.Value = 1 Then
'        WEEE = "/"
'     End If
'     If chkChinaRoHS.Value = 1 Then
'        ChinaRoHS = "China RoHS"
'     ElseIf chkNonChinaRoHS.Value = True Then
'        ChinaRoHS = "/"
'     End If
'     If optH3CRoHS.Value = 1 Then
'        RoHS = "H3C RoHS"
'     ElseIf opt3COMRoHS.Value = 1 Then
'        RoHS = "3COM RoHS"
'     ElseIf optNonRoHS.Value = 1 Then
'        RoHS = "/"
'     End If
'     If CE <> "" Then
'        sql = sql & " and CE='" & CE & "'"
'     End If
'     If WEEE <> "" Then
'        sql = sql & " and WEEE='" & WEEE & "'"
'     End If
'     If ChinaRoHS <> "" Then
'        sql = sql & " and ChinaRoHS='" & ChinaRoHS & "'"
'     End If
'     If RoHS <> "" Then
'        sql = sql & " and RoHS='" & RoHS & "'"
'     End If
'     If txtMS.Text <> "" Then
'        sql = sql & " and MS='" & txtMS.Text & "'"
'     End If
'     If txtNAL.Text <> "" Then
'        sql = sql & " and NAL='" & txtNAL.Text & "'"
'     End If
'     If txtChangNAL.Text <> "" Then
'        sql = sql & " and ChangNAL='" & txtChangNAL.Text & "'"
'     End If
'      If txtHV.Text <> "" Then
'        sql = sql & " and HV='" & txtHV.Text & "'"
'     End If
'     If txtRemark.Text <> "" Then
'        sql = sql & " and Remark='" & txtRemark.Text & "'"
'     End If
     sql = sql & " order by ID,SN"
     rec.Open sql, conn, adOpenKeyset, adLockOptimistic
     Set mfgH3C.DataSource = rec
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdSelect_Click()
   On Error Resume Next
   cdSelect.CancelError = True
   cdSelect.Filter = "*.xls|*.xls"
   cdSelect.action = 1
   If cdSelect.FileName <> "" Then txtPath.Text = cdSelect.FileName
End Sub

Private Sub cmdUpdate_Click()
   If mfgH3C.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫ�޸ĵ���!", vbInformation + vbOKOnly, "δѡ����"
      Exit Sub
   End If
   mfgH3C_Click
   enable
   txtSN.Enabled = False
   txtSN.BackColor = &HE0E0E0
   op = "Update"
End Sub

Private Sub Form_Load()
   unable
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   renovate
End Sub

Private Sub renovate()
   sql = "select * from tblH3C order by ID,SN"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgH3C.DataSource = rec
   With mfgH3C
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 3000
        .ColWidth(4) = 2500
        .ColWidth(5) = 3000
        .ColWidth(6) = 3500
        .ColWidth(7) = 3000
        .ColWidth(8) = 1500
        .ColWidth(9) = 1000
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        .ColWidth(13) = 1500
        .ColWidth(14) = 1500
        .ColWidth(15) = 2000
        .ColWidth(16) = 2000
        .ColWidth(17) = 1500
        .ColWidth(18) = 1000
        .ColWidth(19) = 2000
        
        .TextMatrix(0, 1) = "���(ID)"
        .TextMatrix(0, 2) = "Ӳ���汾(Hardware Version)"
        .TextMatrix(0, 3) = "��Ʒ����(Model Number)"
        .TextMatrix(0, 4) = "��Ʒ����(����)(Chinese Product Name)"
        .TextMatrix(0, 5) = "��Ʒ����(Ӣ��)(English Product Name)"
        .TextMatrix(0, 6) = "��Ʒ����(Description)"
        .TextMatrix(0, 7) = "����ߴ�(Outside Size)"
        .TextMatrix(0, 8) = "ë��(Gross Weight)"
        .TextMatrix(0, 9) = "��֤��ϢCE"
        .TextMatrix(0, 10) = "��֤��ϢWEEE"
        .TextMatrix(0, 11) = "��֤��ϢChinaRoHS"
        .TextMatrix(0, 12) = "��֤��ϢRoHS"
        .TextMatrix(0, 13) = "��֤��ϢTurkeyRoHS"
        .TextMatrix(0, 14) = "�����׼(China MFG Standards)"
        .TextMatrix(0, 15) = "�����׼��Ч��(Valid From)"
        .TextMatrix(0, 16) = "������ɺ�(China N.A.L.)"
        .TextMatrix(0, 17) = "������Ч��(Valid From)"
        .TextMatrix(0, 18) = "�Ƿ��ӡ����汾"
        .TextMatrix(0, 19) = "��ע(Remark)"
   End With
   rec.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rec.State = 1 Then
      rec.Close
      Set conn = Nothing
   End If
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub

Private Sub mfgH3C_Click()
   If mfgH3C.RowSel > 0 Then
      txtHV.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 2)
      txtSN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 3)
      txtCPN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 4)
      txtEPN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 5)
      txtDes.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 6)

      txtOS.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 7)
      txtGW.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 8)
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 9))) = "CE" Then
         chkCE.Value = 1
         chkNonCE.Value = 0
         chkCEAddr.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 9) = "N/A" Then
         chkCE.Value = 0
         chkNonCE.Value = 1
         chkCEAddr.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 9) = "CE+CE Addr" Then
         chkCE.Value = 1
         chkNonCE.Value = 0
         chkCEAddr.Value = 1
      End If
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 10))) = "WEEE" Then
         chkWEEE.Value = 1
         chkNonWEEE.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 10) = "N/A" Then
         chkWEEE.Value = 0
         chkNonWEEE.Value = 1
      End If
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 11))) = "CHINA ROHS" Then
         chkChinaROHS.Value = 1
         chkNonChinaROHS.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 11) = "N/A" Then
         chkChinaROHS.Value = 0
         chkNonChinaROHS.Value = 1
      End If
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 12))) = "H3C ROHS" Then
        chkRoHS.Value = 1
        chkNonRoHS.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 12) = "N/A" Then
         chkRoHS.Value = 0
         chkNonRoHS.Value = 1
      End If
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 13))) = "TURKEY ROHS" Then
         chkTurkeyROHS.Value = 1
         chkNonTurkeyROHS.Value = 0
      ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 13) = "N/A" Then
         chkTurkeyROHS.Value = 0
         chkNonTurkeyROHS.Value = 1
      End If
      
      txtMS.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 14)
      dtpMSValidFrom.Value = mfgH3C.TextMatrix(mfgH3C.RowSel, 15)
      
      txtNAL.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 16)
      dtpValidFrom.Value = mfgH3C.TextMatrix(mfgH3C.RowSel, 17)
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 18))) = "Y" Then
        chkSVPrint.Value = 1
        chkNoPrintSV.Value = 0
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 18))) = "N" Then
        chkSVPrint.Value = 0
        chkNoPrintSV.Value = 1
      End If
      
      txtRemark.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 19)
   End If
End Sub

Private Sub mfgH3C_SelChange()
   mfgH3C_Click
End Sub

