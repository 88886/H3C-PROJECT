VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmH3C_3COMSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H3C-3COM Setting"
   ClientHeight    =   9795
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
   Icon            =   "frmH3C_3COMSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   12120
   StartUpPosition =   2  '��Ļ����
   Begin MSComDlg.CommonDialog cdSelect 
      Left            =   2760
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   735
      Left            =   10740
      TabIndex        =   46
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   735
      Left            =   9300
      TabIndex        =   45
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "ȷ��(Confirm)"
      Height          =   735
      Left            =   7620
      TabIndex        =   44
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(Delete)"
      Height          =   735
      Left            =   10740
      TabIndex        =   43
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�޸�(Update)"
      Height          =   735
      Left            =   9300
      TabIndex        =   42
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "����(Insert)"
      Height          =   735
      Left            =   7620
      TabIndex        =   41
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "��ѯ(Query)"
      Height          =   735
      Left            =   5940
      TabIndex        =   40
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "����(Export)"
      Height          =   735
      Left            =   3420
      TabIndex        =   39
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "����(Import)"
      Height          =   735
      Left            =   3420
      TabIndex        =   38
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ѡ��(Select)"
      Height          =   495
      Left            =   1380
      TabIndex        =   37
      Top             =   9240
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   180
      TabIndex        =   36
      Top             =   8640
      Width           =   3015
   End
   Begin VB.Frame fmH3C 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.TextBox txtRemark 
         Height          =   495
         Left            =   9120
         TabIndex        =   17
         Top             =   4440
         Width           =   2775
      End
      Begin VB.TextBox txtHV 
         Height          =   495
         Left            =   2280
         TabIndex        =   16
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox txtChangNAL 
         Height          =   495
         Left            =   9120
         TabIndex        =   15
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox txtMS 
         Height          =   450
         Left            =   9120
         TabIndex        =   14
         Top             =   2520
         Width           =   2775
      End
      Begin VB.CheckBox chkNon3COMRoHS 
         Caption         =   "��RoHS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CheckBox chk3COMRoHS 
         Caption         =   "RoHS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         Caption         =   "��"
         Height          =   375
         Left            =   10680
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkChinaRoHS 
         Caption         =   "��"
         Height          =   375
         Left            =   9120
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox chkNonCE 
         Caption         =   "�� CE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkCE 
         Caption         =   "CE"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtNAL 
         Height          =   435
         Left            =   2280
         TabIndex        =   7
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox txtGW 
         Height          =   450
         Left            =   9120
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtOS 
         Height          =   450
         Left            =   2280
         TabIndex        =   5
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtDes 
         Height          =   450
         Left            =   9120
         TabIndex        =   4
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtEPN 
         Height          =   450
         Left            =   2280
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtCPN 
         Height          =   450
         Left            =   9120
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtSN 
         Height          =   450
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpValidTo 
         Height          =   495
         Left            =   2280
         TabIndex        =   18
         Top             =   3720
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
         Format          =   21495809
         UpDown          =   -1  'True
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpValidFrom 
         Height          =   495
         Left            =   9120
         TabIndex        =   19
         Top             =   3120
         Width           =   2775
         _ExtentX        =   4895
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
         Format          =   21495809
         UpDown          =   -1  'True
         CurrentDate     =   39757
      End
      Begin VB.Label lblOS 
         Caption         =   "��ߴ�(MM):"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         Caption         =   "������ɺ�:"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label lblRemark 
         Caption         =   "��ע:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   33
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label lblHV 
         Caption         =   "Ӳ���汾:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label lblChangNAL 
         Caption         =   "�滻������:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   31
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblValidTo 
         Caption         =   "������Ч��ֹ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblMS 
         Caption         =   "�����׼:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   29
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblRoHS 
         Caption         =   "3COM RoHS:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblWEEE 
         Caption         =   "China RoHS:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   27
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblCE 
         Caption         =   "��֤��ϢCE:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblValidFrom 
         Caption         =   "������Ч��ʼ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   25
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label lblGW 
         Caption         =   "ë��(KG):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   24
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblDes 
         Caption         =   "��Ʒ����:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   23
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblEPN 
         Caption         =   "��Ʒ����(Ӣ��):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblCPN 
         Caption         =   "��Ʒ����(����):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   21
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblSN 
         Caption         =   "��Ʒ����:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgH3C_3COM 
      Height          =   2775
      Left            =   60
      TabIndex        =   34
      Top             =   5160
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4895
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Left            =   180
      TabIndex        =   35
      Top             =   8160
      Width           =   2175
   End
End
Attribute VB_Name = "frmH3C_3COMSetting"
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
   chkNonCE.Enabled = True
   chkChinaRoHS.Enabled = True
   chkNonChinaRoHS.Enabled = True
   chk3COMRoHS.Enabled = True
   chkNon3COMRoHS.Enabled = True
   txtMS.Enabled = True
   txtMS.BackColor = &HFFFFFF
   txtNAL.Enabled = True
   txtNAL.BackColor = &HFFFFFF
   dtpValidFrom.Enabled = True
   dtpValidTo.Enabled = True
   txtChangNAL.Enabled = True
   txtChangNAL.BackColor = &HFFFFFF
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
   chkNonCE.Enabled = False
   chkChinaRoHS.Enabled = False
   chkNonChinaRoHS.Enabled = False
   chk3COMRoHS.Enabled = False
   chkNon3COMRoHS.Enabled = False
   txtMS.Enabled = False
   txtMS.BackColor = &HE0E0E0
   txtNAL.Enabled = False
   txtNAL.BackColor = &HE0E0E0
   dtpValidFrom.Enabled = False
   dtpValidTo.Enabled = False
   txtChangNAL.Enabled = False
   txtChangNAL.BackColor = &HE0E0E0
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

Private Sub chkNon3COMRoHS_Click()
   If chkNon3COMRoHS.Value = 1 Then
      chk3COMRoHS.Value = 0
   Else
      chk3COMRoHS.Value = 1
   End If
End Sub

Private Sub chkOS_Click()
   If chkOS.Value = 1 Then
      txtOS.Enabled = True
      txtOS.BackColor = &HFFFFFF
   Else
      txtOS.Enabled = False
      txtOS.BackColor = &HE0E0E0
      txtOS.Text = "/"
   End If
End Sub

Private Sub chkNonCE_Click()
   If chkNonCE.Value = 1 Then
      chkCE.Value = 0
   Else
      chkCE.Value = 1
   End If
End Sub

Private Sub chkNonChinaRoHS_Click()
   If chkNonChinaRoHS.Value = 1 Then
      chkChinaRoHS.Value = 0
   Else
      chkChinaRoHS.Value = 1
   End If
End Sub

Private Sub chk3COMRoHS_Click()
   If chk3COMRoHS.Value = 1 Then
      chkNon3COMRoHS.Value = 0
   Else
      chkNon3COMRoHS.Value = 1
   End If
End Sub

Private Sub chkChinaRoHS_Click()
   If chkChinaRoHS.Value = 1 Then
      chkNonChinaRoHS.Value = 0
   Else
      chkNonChinaRoHS.Value = 1
   End If
End Sub

Private Sub cmdCancel_Click()
   unable
   op = ""
End Sub

Private Sub cmdConfirm_Click()
   If txtSN.Text = "" Then
      MsgBox "��Ʒ���벻��Ϊ��!!", vbExclamation + vbOKOnly, "��Ʒ�����"
      txtSN.SetFocus
      Exit Sub
   End If
   If txtCPN.Text = "" Then
       MsgBox "��Ʒ����(����)����Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ����(����)�ſ�"
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
       MsgBox "��ߴ粻��Ϊ��!!", vbExclamation + vbOKOnly, "��ߴ��"
       txtOS.SetFocus
       Exit Sub
   End If
   If txtGW.Text = "" Then
      MsgBox "ë�ز���Ϊ��!", vbExclamation + vbOKOnly, "ë�ؿ�"
      txtGW.SetFocus
      Exit Sub
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
'   If txtChangNAL.Text = "" Then
'      MsgBox "�滻�����Ų���Ϊ��!"
'      txtChangNAL.SetFocus
'      Exit Sub
'   End If
   If txtHV.Text = "" Then
      MsgBox "Ӳ���汾����Ϊ��!", vbExclamation + vbOKOnly, "Ӳ���汾��"
      txtHV.SetFocus
      Exit Sub
   End If
   Dim CE, ChinaRoHS, RoHS As String
   If chkCE.Value = 1 Then
      CE = "CE"
   ElseIf chkNonCE.Value = 1 Then
      CE = "/"
   End If
   If chkChinaRoHS.Value = 1 Then
      ChinaRoHS = "China RoHS"
   ElseIf chkNonChinaRoHS.Value = 1 Then
      ChinaRoHS = "/"
   End If
   If chk3COMRoHS.Value = 1 Then
      RoHS = "3COM RoHS"
   ElseIf chkNon3COMRoHS.Value = 1 Then
      RoHS = "/"
   End If
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from H3C_3COM_3COM where SN='" & txtSN.Text & "'"
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "��Ʒ�����Ѵ���!", vbExclamation + vbOKOnly, "��Ʒ����ظ�"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close
      sql = "Insert into H3C_3COM(ID,SN,CPN,EPN,Des,OS,GW,CE,ChinaRoHS,RoHS,MS,NAL,ValidFrom,ValidTo,ChangNAL,HV,Remark) " & _
            "Values(" & getmaxID("H3C_3COM") & ",'" & txtSN.Text & "','" & txtCPN.Text & "','" & txtEPN.Text & "','" & txtDes.Text & "','" & txtOS.Text & "','" & txtGW.Text & "','" & CE & "','" & ChinaRoHS & "','" & RoHS & "'," & _
            "'" & txtMS.Text & "','" & txtNAL.Text & "','" & dtpValidFrom.Value & "','" & dtpValidTo.Value & "','" & txtChangNAL.Text & "','" & txtHV.Text & "','" & txtRemark.Text & "')"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "����H3C_3COM_3COM�趨����ʧ��!" & "ԭ����" & status, vbOKOnly + vbInformation, "����ʧ��"
      End If
      MsgBox "����H3C_3COM_3COM�趨���ϳɹ�!", vbOKOnly + vbInformation, "�����ɹ�"
      renovate
      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update H3C_3COM set CPN='" & txtCPN.Text & "',EPN='" & txtEPN.Text & "',Des='" & txtDes.Text & "',OS='" & txtOS.Text & "',GW='" & txtGW.Text & "',CE='" & CE & "',ChinaRoHS='" & ChinaRoHS & "',RoHS='" & RoHS & "'," & _
            "MS='" & txtMS.Text & "',NAL='" & txtNAL.Text & "',ValidFrom='" & dtpValidFrom.Value & "',ValidTo='" & dtpValidTo.Value & "',ChangNAL='" & txtChangNAL.Text & "',HV='" & txtHV.Text & "',Remark='" & txtRemark.Text & "'" & _
            " where ID=" & mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 1) & " and SN='" & txtSN.Text & "'"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "�޸�H3C_3COM_3COM�趨����ʧ��!" & "ԭ����" & status, vbOKOnly + vbInformation, "�޸�ʧ��"
      End If
      MsgBox "�޸�H3C_3COM_3COM�趨���ϳɹ�!", vbOKOnly + vbInformation, "�޸ĳɹ�"
      renovate
      cmdCancel_Click
   End If
   renovate
End Sub

Private Sub cmdDelete_Click()
   If mfgH3C_3COM_3COM.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫɾ������!", vbInformation + vbOKOnly, "δѡ����"
      Exit Sub
   End If
   sql = "delete from H3C_3COM_3COM where ID=" & mfgH3C_3COM_3COM.TextMatrix(mfgH3C_3COM_3COM.RowSel, 1) & " and SN='" & mfgH3C_3COM_3COM.TextMatrix(mfgH3C_3COM_3COM.RowSel, 2) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "ɾ��H3C_3COM_3COM�趨����ʧ��!" & "ԭ����" & status, vbInformation + vbOKOnly, "ɾ��ʧ��"
   End If
   MsgBox "ɾ��H3C_3COM_3COM�趨���ϳɹ�!", vbInformation + vbOKOnly, "ɾ���ɹ�"
   renovate
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If mfgH3C_3COM_3COM.Rows = 0 Then
      MsgBox "�����Ͽɻ��"
      Exit Sub
   End If
   If txtPath.Text <> "" Then
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets.Item(1)
       For i = 0 To mfgH3C_3COM_3COM.Rows - 1
         For j = 1 To mfgH3C_3COM_3COM.Cols - 1
          xlSheet.Cells(i + 1, j) = mfgH3C_3COM_3COM.TextMatrix(i, j)
       Next j
      Next i
      xlBook.SaveAs (txtPath.Text)
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "�����EXCEL���ϳɹ�!!"
    End If
End Sub

Private Sub cmdImport_Click()
   If txtPath.Text = "" Then
      MsgBox "����·������Ϊ��!", vbOKOnly + vbExclamation, "����·����"
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
          Dim isexist As Boolean
          If xlSheet.Cells(j, 17) = "" Then
             MsgBox "�������ϸ�ʽ����ȷ!", vbOKOnly + vbExclamation, "�����ʽ����"
             Exit Sub
          End If
          isexist = False
          For K = 1 To 17
           If K = 2 Then
             cellValue = xlSheet.Cells(j, K)
             If cellValue = "" Then
                MsgBox "�������ϸ�ʽ����ȷ!", vbOKOnly + vbExclamation, "�����ʽ����"
                Exit Sub
             End If
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from H3C_3COM where SN='" & cellValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("��Ʒ�����Ѵ���!", vbAbortRetryIgnore + vbExclamation, "�����ظ�")
                End If
                
                If action = vbAbort Then
                   MsgBox "���ϵ�������ֹ!!", vbOKOnly + vbInformation, "������ֹ"
                   rcd.Close
                   Exit Sub
                ElseIf action = vbIgnore And info = True Then
                   MsgBox "�ظ���Ʒ������ϲ��ᵼ��,���Ե�..!!", vbInformation + vbOKOnly, "�ظ��Ĳ�����"
                   rcd.Close
                   info = False
                   Exit For
                ElseIf action = vbRetry And info = True Then
                   MsgBox "�ظ���Ʒ������ϻ��Զ�����,���Ե�..!!", vbInformation + vbOKOnly, "�ظ��Ļ��Զ�����"
                   info = False
                End If
                isexist = True
             Else
                isexist = False
             End If
             rcd.Close
            End If
            
            If K = 17 Then
               If action = vbRetry Then
                   sql = "Update H3C_3COM_3COM set CPN='" & xlSheet.Cells(j, 3) & "',EPN='" & xlSheet.Cells(j, 4) & "',Des='" & xlSheet.Cells(j, 5) & "',OS='" & xlSheet.Cells(j, 6) & "',GW='" & xlSheet.Cells(j, 7) & "',CE='" & xlSheet.Cells(j, 8) & "',ChinaRoHS='" & xlSheet.Cells(j, 9) & "',RoHS='" & xlSheet.Cells(j, 10) & "'," & _
                        "MS='" & xlSheet.Cells(j, 11) & "',NAL='" & xlSheet.Cells(j, 12) & "',ValidFrom='" & xlSheet.Cells(j, 13) & "',ValidTo='" & xlSheet.Cells(j, 14) & "',ChangNAL='" & xlSheet.Cells(j, 15) & "',HV='" & xlSheet.Cells(j, 16) & "',Remark='" & xlSheet.Cells(j, 17) & "'" & _
                        " where SN='" & xlSheet.Cells(j, 2) & "'"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "�޸�H3C_3COM_3COM�趨����ʧ��!" & "ԭ����" & status, vbOKOnly + vbExclamation, "�޸�ʧ��"
                   End If
'                   MsgBox "�޸�H3C_3COM_3COM�趨���ϳɹ�!"
               ElseIf isexist = False Then
                   sql = "Insert into H3C_3COM(ID,SN,CPN,EPN,Des,OS,GW,CE,ChinaRoHS,RoHS,MS,NAL,ValidFrom,ValidTo,ChangNAL,HV,Remark) " & _
                        "Values(" & getmaxID("H3C_3COM") & ",'" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "','" & xlSheet.Cells(j, 4) & "','" & xlSheet.Cells(j, 5) & "','" & xlSheet.Cells(j, 6) & "','" & xlSheet.Cells(j, 7) & "','" & xlSheet.Cells(j, 8) & "','" & xlSheet.Cells(j, 9) & "','" & xlSheet.Cells(j, 10) & "'," & _
                        "'" & xlSheet.Cells(j, 11) & "','" & xlSheet.Cells(j, 12) & "','" & xlSheet.Cells(j, 13) & "','" & xlSheet.Cells(j, 14) & "','" & xlSheet.Cells(j, 15) & "','" & xlSheet.Cells(j, 16) & "','" & xlSheet.Cells(j, 17) & "')"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "����H3C_3COM_3COM�趨����ʧ��!" & "ԭ����" & status, vbOKOnly + vbInformation, "����ʧ��"
                   End If
'                   MsgBox "����H3C_3COM_3COM�趨���ϳɹ�!"
               End If
           End If
         Next K
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "H3C_3COM_3COM�趨���ϵ���ɹ�!", vbOKOnly + vbInformation, "����ɹ�"
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
   txtMS.Text = "/"
'   chkOS.Value = 1
   chkCE.Value = 1
   chkChinaRoHS.Value = 1
   chk3COMRoHS.Value = 1
   txtNAL.Text = "/"
   dtpValidFrom.Value = Date
   dtpValidTo.Value = "9999-12-31"
   txtChangNAL.Text = "/"
   txtHV.Text = "/"
   txtRemark.Text = "/"
   op = "Insert"
End Sub

Private Sub cmdQuery_Click()
   If txtSN.Enabled = False Then
      MsgBox "�밴������ť��վͿ������ѯ����!", vbOKOnly + vbInformation, "�����ѯ����"
   End If
   If rec.State = 1 Then
      rec.Close
   End If
   sql = "select * from H3C_3COM Where 1=1"
   If txtSN.Text <> "" Then
      sql = sql & " and SN like '%" & txtSN.Text & "%'"
   End If
   If txtCPN.Text <> "" Then
      sql = sql & " and CPN like '%" & txtCPN.Text & "%'"
   End If
   If txtEPN.Text <> "" Then
      sql = sql & " and EPN like '%" & txtEPN.Text & "%'"
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
'   Dim CE, ChinaRoHS, ThreeComRoHS As String
'   If chkCE.Value = 1 Then
'      CE = "CE"
'   ElseIf chkNonCE.Value = 1 Then
'      CE = "/"
'   End If
'   If chkChinaRoHS.Value = 1 Then
'      ChinaRoHS = "China RoHS"
'   ElseIf chkNonRoHS.Value = 1 Then
'      ChinaRoHS = "/"
'   End If
'   If chk3COMRoHS.Value = 1 Then
'      ThreeComRoHS = "3COM RoHS"
'   ElseIf chkNon3COMRoHS.Value = 1 Then
'      ThreeComRoHS = "/"
'   End If
'   If CE <> "" Then
'      sql = sql & " and CE='" & CE & "'"
'   End If
'   If WEEE <> "" Then
'      sql = sql & " and WEEE='" & WEEE & "'"
'   End If
'   If ChinaRoHS <> "" Then
'      sql = sql & " and ChinaRoHS='" & ChinaRoHS & "'"
'   End If
'   If ThreeComRoHS <> "" Then
'      sql = sql & " and RoHS='" & ThreeComRoHS & "'"
'   End If
'   If txtMS.Text <> "" Then
'      sql = sql & " and MS='" & txtMS.Text & "'"
'   End If
'   If txtNAL.Text <> "" Then
'      sql = sql & " and NAL='" & txtNAL.Text & "'"
'   End If
'   If txtChangNAL.Text <> "" Then
'      sql = sql & " and ChangNAL='" & txtChangNAL.Text & "'"
'   End If
'    If txtHV.Text <> "" Then
'      sql = sql & " and HV='" & txtHV.Text & "'"
'   End If
'   If txtRemark.Text <> "" Then
'      sql = sql & " and Remark='" & txtRemark.Text & "'"
'   End If
   sql = sql & " Order by ID,SN"
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgH3C_3COM.DataSource = rec
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdSelect_Click()
   cdSelect.CancelError = True
   cdSelect.Filter = "*.xls|*.xls"
   cdSelect.action = 1
   If cdSelect.FileName <> "" Then txtPath.Text = cdSelect.FileName
End Sub

Private Sub cmdUpdate_Click()
   If mfgH3C_3COM.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫ�޸ĵ���!", vbInformation + vbOKOnly, "δѡ����"
      Exit Sub
   End If
   mfgH3C_3COM_Click
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
   sql = "select * from H3C_3COM order by ID,SN"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgH3C_3COM.DataSource = rec
   With mfgH3C_3COM
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 3000
        .ColWidth(4) = 3000
        .ColWidth(5) = 2500
        .ColWidth(6) = 3500
        .ColWidth(7) = 1500
        .ColWidth(8) = 1000
        .ColWidth(9) = 2500
        .ColWidth(10) = 2500
        .ColWidth(11) = 1500
        .ColWidth(12) = 1500
        .ColWidth(13) = 2000
        .ColWidth(14) = 2000
        .ColWidth(15) = 2000
        .ColWidth(16) = 1000
        .ColWidth(17) = 2000
        
        .TextMatrix(0, 1) = "���(ID)"
        .TextMatrix(0, 2) = "��Ʒ����(Model Number)"
        .TextMatrix(0, 3) = "��Ʒ����(����)(Chinese Product Name)"
        .TextMatrix(0, 4) = "��Ʒ����(Ӣ��)(English Product Name)"
        .TextMatrix(0, 5) = "��Ʒ����(Description)"
        .TextMatrix(0, 6) = "����ߴ�(Outside Size)"
        .TextMatrix(0, 7) = "ë��(Gross Weight)"
        .TextMatrix(0, 8) = "��֤��ϢCE"
        .TextMatrix(0, 9) = "��֤��ϢChina RoHS"
        .TextMatrix(0, 10) = "��֤��Ϣ3COM RoHS"
        .TextMatrix(0, 11) = "�����׼(China MFG Standards)"
        .TextMatrix(0, 12) = "������ɺ�(China N.A.L.)"
        .TextMatrix(0, 13) = "������Ч��ʼ(Valid From)"
        .TextMatrix(0, 14) = "������Ч��ֹ(Valid To)"
        .TextMatrix(0, 15) = "�滻������(Changed N.A.L)"
        .TextMatrix(0, 16) = "Ӳ���汾(Hardware Version)"
        .TextMatrix(0, 17) = "��ע(Remark)"
   End With
   rec.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rec.State = 1 Then
      rec.Close
      Set rec = Nothing
   End If
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub

Private Sub mfgH3C_3COM_Click()
   If mfgH3C_3COM.RowSel > 0 Then
      txtSN.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 2)
      txtCPN.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 3)
      txtEPN.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 4)
      txtDes.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 5)
      txtOS.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 6)
'      If txtOS.Text <> "" And txtOS.Text <> "/" Then
'         chkOS.Value = 1
'         txtOS.Enabled = False
'         txtOS.BackColor = &HE0E0E0
'      Else
'         chkOS.Value = 0
'      End If
      txtGW.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 7)
      If Trim(mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 8)) = "CE" Then
         chkCE.Value = 1
         chkNonCE.Value = 0
      ElseIf mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 8) = "/" Then
         chkCE.Value = 0
         chkNonCE.Value = 1
      End If
      If UCase(mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 9)) = "CHINA ROHS" Then
         chkChinaRoHS.Value = 1
         chkNonChinaRoHS.Value = 0
      ElseIf mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 9) = "/" Then
         chkChinaRoHS.Value = 0
         chkNonChinaRoHS.Value = 1
      End If
      If UCase(mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 10)) = "3COM ROHS" Then
         chk3COMRoHS.Value = 1
         chkNon3COMRoHS.Value = 0
      ElseIf mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 10) = "/" Then
         chk3COMRoHS.Value = 0
         chkNon3COMRoHS.Value = 1
      End If
      txtMS.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 11)
      txtNAL.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 12)
      dtpValidFrom.Value = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 13)
      dtpValidTo.Value = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 14)
      txtChangNAL.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 15)
      txtHV.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 16)
      txtRemark.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 17)
   End If
End Sub

Private Sub mfgH3C_3COM_SelChange()
   mfgH3C_3COM_Click
End Sub


