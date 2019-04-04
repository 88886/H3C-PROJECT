VERSION 5.00
Begin VB.Form frmHUVshenchan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HUV Label Print"
   ClientHeight    =   11730
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   14130
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHUVshenchan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11730
   ScaleWidth      =   14130
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ӡ(Print) &p"
      Height          =   615
      Left            =   4080
      TabIndex        =   34
      Top             =   11100
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   615
      Left            =   9000
      TabIndex        =   33
      Top             =   11100
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   615
      Left            =   6480
      TabIndex        =   32
      Top             =   11100
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      TabIndex        =   19
      Top             =   5040
      Width           =   13935
      Begin VB.TextBox txtCNCN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2040
         TabIndex        =   65
         Top             =   4800
         Width           =   4095
      End
      Begin VB.TextBox txtCNCA 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   9600
         TabIndex        =   64
         Top             =   4800
         Width           =   4095
      End
      Begin VB.TextBox txtENCN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   450
         Left            =   2040
         TabIndex        =   63
         Top             =   5355
         Width           =   4095
      End
      Begin VB.TextBox txtENCA 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   450
         Left            =   9600
         TabIndex        =   62
         Top             =   5355
         Width           =   4095
      End
      Begin VB.CheckBox chkNonFCC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   60
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkFCC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   59
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtSalesArea 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6360
         TabIndex        =   57
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chkNonCN3C 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   12120
         TabIndex        =   55
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkCN3C 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10920
         TabIndex        =   54
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtPO 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ǻ���"
         Height          =   495
         Left            =   6360
         TabIndex        =   51
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox lblNALday 
         Height          =   450
         Left            =   6600
         TabIndex        =   50
         Top             =   4320
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox lblMSday 
         Height          =   450
         Left            =   6240
         TabIndex        =   49
         Top             =   4320
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CheckBox chkNonTurkey 
         BackColor       =   &H80000005&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   48
         Top             =   3120
         Width           =   855
      End
      Begin VB.CheckBox chkTurkey 
         BackColor       =   &H80000005&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   47
         Top             =   3120
         Width           =   735
      End
      Begin VB.CheckBox chkVer 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   8520
         TabIndex        =   44
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   7200
         TabIndex        =   43
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optNonRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   495
         Left            =   13080
         TabIndex        =   41
         Top             =   3120
         Width           =   615
      End
      Begin VB.OptionButton opt3COMRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3COMRoHS"
         Enabled         =   0   'False
         Height          =   615
         Left            =   11760
         TabIndex        =   40
         Top             =   3000
         Width           =   1215
      End
      Begin VB.OptionButton optH3CRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HUVRoHS"
         Enabled         =   0   'False
         Height          =   615
         Left            =   10560
         TabIndex        =   39
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtMN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   37
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ߴ�(MM):"
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   2160
         Width           =   855
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��CE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10920
         TabIndex        =   9
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   12120
         TabIndex        =   10
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox chkChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   16
         Top             =   4320
         Width           =   3135
      End
      Begin VB.TextBox txtNAL 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   14
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtMS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   13
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox txtHV 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   15
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox txtGW 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   6
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtOS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   5
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtDes 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   4
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtEPN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   3
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   3120
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtVer 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���Ĺ�˾��:"
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���ĵ�ַ��:"
         Height          =   375
         Left            =   7920
         TabIndex        =   68
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ӣ�Ĺ�˾��:"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   5355
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ӣ�ĵ�ַ��:"
         Height          =   375
         Left            =   7920
         TabIndex        =   66
         Top             =   5355
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FCC:"
         Height          =   375
         Left            =   5520
         TabIndex        =   61
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ǩ����:"
         Height          =   375
         Left            =   4920
         TabIndex        =   58
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��֤��ϢCCC:"
         Height          =   375
         Left            =   8640
         TabIndex        =   56
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblPO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ͬ��/PO"
         Height          =   375
         Left            =   6360
         TabIndex        =   52
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblTurkeyRohs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������RoHs:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�汾��Ϣ:"
         Height          =   375
         Left            =   8880
         TabIndex        =   45
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����:"
         Height          =   375
         Left            =   6360
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��֤��ϢRoHS:"
         Height          =   495
         Left            =   8400
         TabIndex        =   38
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ�ͺ�:"
         Height          =   375
         Left            =   8400
         TabIndex        =   36
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ע:"
         Height          =   375
         Left            =   8400
         TabIndex        =   31
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������ɺ�:"
         Height          =   375
         Left            =   8400
         TabIndex        =   30
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ִ�б�׼:"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��֤��ϢChinaRoHS:"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label lblHV 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ӳ���汾:"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label lblWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��֤��ϢWEEE:"
         Height          =   375
         Left            =   8400
         TabIndex        =   26
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��֤��ϢCE:"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblGW 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ë��(kg):"
         Height          =   375
         Left            =   8400
         TabIndex        =   24
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����(Ӣ��):"
         Height          =   375
         Left            =   8400
         TabIndex        =   23
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����(����):"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ��ʼ����:"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   120
      Picture         =   "frmHUVshenchan.frx":13652
      ScaleHeight     =   4785
      ScaleWidth      =   13905
      TabIndex        =   18
      Top             =   300
      Width           =   13935
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HUV ������"
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
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmHUVshenchan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim mydoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects

Private Sub Check1_Click()
If Check1.Value = 1 Then
    chkNonWEEE.Value = 1
    chkNonTurkey.Value = 1
    optNonRoHS.Value = True
    Check1.Enabled = False
End If

End Sub

Private Sub chkCE_Click()
   If chkCE.Value = 1 Then
      chkNonCE.Value = 0
   Else
      chkNonCE.Value = 1
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

Private Sub chkNonTurkey_Click()
    If chkNonTurkey.Value = 1 Then
      chkTurkey.Value = 0
   Else
      chkTurkey.Value = 1
   End If
End Sub

Private Sub chkTurkey_Click()
    If chkTurkey.Value = 1 Then
      chkNonTurkey.Value = 0
    Else
      chkNonTurkey.Value = 1
    End If
End Sub

Private Sub chkNonWEEE_Click()
   If chkNonWEEE.Value = 1 Then
      chkWEEE.Value = 0
   Else
      chkWEEE.Value = 1
   End If
End Sub

Private Sub chkOS_Click()
   If chkOS.Value = 1 Then
      txtOS.Enabled = True
      txtOS.BackColor = &H80000005
   Else
      txtOS.Enabled = False
      txtOS.BackColor = &HC0C0C0
   End If
End Sub

Private Sub chkChinaRoHS_Click()
   If chkChinaRoHS.Value = 1 Then
      chkNonChinaRoHS.Value = 0
   Else
      chkNonChinaRoHS.Value = 1
   End If
End Sub


Private Sub chkWEEE_Click()
   If chkWEEE.Value = 1 Then
      chkNonWEEE.Value = 0
   Else
      chkNonWEEE.Value = 1
   End If
End Sub

Private Sub cmdCancel_Click()
    txtPO.Text = ""
   txtSN.Text = ""
   txtVer.Text = ""
   txtCPN.Text = ""
   txtEPN.Text = ""
   txtDes.Text = ""
   txtMN.Text = ""
   txtOS.Text = ""
   txtGW.Text = ""
   txtQty.Text = ""
'   chkCE.Value = 0
   chkNonCE.Value = 0
'   chkWEEE.Value = 0
   chkNonWEEE.Value = 0
'   chkRoHS.Value = 0
   chkNonChinaRoHS.Value = 0
   chkTurkey.Value = 1
   optH3CRoHS.Value = True
   txtMS.Text = ""
   txtNAL.Text = ""
   txtHV.Text = ""
   txtRemark.Text = ""
   Me.lblMSday.Text = ""
   Me.lblNALday.Text = ""
   Me.chkFCC.Value = 0
   Me.chkNonFCC.Value = 0
   
   Check1.Value = 0
   Check1.Enabled = True
   
   txtCNCN.Text = ""
   txtCNCA.Text = ""
   txtENCN.Text = ""
   txtENCA.Text = ""
   Me.txtSalesArea.Text = ""
   
   txtSN.SetFocus
   
End Sub

Private Sub cmdPrint_Click()
'On Error Resume Next
   If Trim(txtSN.Text) = "" Then
      MsgBox "��Ʒ����δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ�����Ʒ����"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If txtQty.Text = "" Then
      MsgBox "����δ���룬���ܴ�ӡ��", vbInformation + vbOKOnly, "δ��������"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If CInt(txtQty.Text) = 0 Then
      MsgBox "��������ȷ��������", vbInformation + vbOKOnly, "��������"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If Trim(txtGW.Text) = "" Then
      MsgBox "��Ʒ����δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ����ë��"
      txtGW.SetFocus
      Exit Sub
   End If
   
   If Trim(txtVer.Text) = "" Then
      MsgBox "�汾δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ����汾"
      txtVer.SetFocus
      Exit Sub
   End If
   
   If txtHV.Text = "" Then
      MsgBox "Ӳ���汾δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ����Ӳ���汾"
      txtHV.SetFocus
      Exit Sub
   End If
   
  If DateDiff("d", CDate(Trim(Me.lblMSday.Text)), Date) >= 0 Then
      MsgBox "�����׼��Ч���ѹ���,���ܴ�ӡ!", vbInformation + vbOKOnly, "�����׼����"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If DateDiff("d", CDate(Trim(Me.lblNALday.Text)), Date) >= 0 Then
      MsgBox "���������Ч���ѹ���,���ܴ�ӡ!", vbInformation + vbOKOnly, "������ɹ���"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If Me.txtSalesArea.Text = "" Then
      MsgBox "û����������,���ܴ�ӡ!", vbInformation + vbOKOnly, "����������Ϊ��"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If Me.chkCN3C.Value + Me.chkNonCN3C.Value = 0 Or Me.chkCN3C.Value + Me.chkNonCN3C.Value = 2 Then
      MsgBox "CCC��֤û���趨,���ܴ�ӡ!", vbInformation + vbOKOnly, "CCC��֤����Ϊ��"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If Me.chkFCC.Value + Me.chkNonFCC.Value = 0 Or Me.chkFCC.Value + Me.chkNonFCC.Value = 2 Then
      MsgBox "FCC��֤û���趨,���ܴ�ӡ!", vbInformation + vbOKOnly, "FCC��֤����Ϊ��"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If Trim(txtPO.Text) <> "/" Then
        If IsNumeric(txtPO.Text) = False Then
            MsgBox "PO�ű���������!", vbInformation + vbOKOnly, "PO��"
            txtPO.SetFocus
            Exit Sub
        End If
        If Len(Trim(txtPO.Text)) <> 10 Then
            MsgBox "PO�ű�����10λ!", vbInformation + vbOKOnly, "PO��"
            txtPO.SetFocus
            Exit Sub
        End If
   End If
   
   
   Dim i, qty As Integer
   Dim leftstr, rightstr, str As String
   qty = CInt(txtQty.Text)
   leftstr = UCase(Left(Trim(txtSN.Text), 14))
   rightstr = Right(Trim(txtSN.Text), 6)
   
     OpenLppx
     
   For i = 0 To qty - 1
   
   
      str = leftstr & Right("000000" & CStr(CLng(rightstr) + i), 6)
      
   
 
   myVars.Item("SN").Value = str
   myVars.Item("Item").Value = UCase(Mid(Trim(txtSN.Text), 3, 8))
   

'   If chkVer.Value = 1 Then
'        If txtVer.Text = "" Or txtVer.Text = "/" Then
'            myObjs("BSver").Top = 100000
'            myVars.Item("SVer").Value = "N/A"
'        Else
'            myObjs("TSver").Top = 100000
'            myVars.Item("SVer").Value = UCase(txtVer.Text)
'        End If
'   Else
'        myObjs("BSver").Top = 100000
'        myVars.Item("SVer").Value = "N/A"
'   End If
'

'update by allen.yan for shun.huang's request 2014.06.27

    If Me.txtSalesArea.Text = "Ѷ��" Then
   
        If chkVer.Value = 1 Then
             If txtVer.Text = "" Or txtVer.Text = "/" Then
                 myObjs("BSver").Top = 100000
                 myVars.Item("SVer").Value = "N/A"
             Else
                 myObjs("TSver").Top = 100000
                 myVars.Item("SVer").Value = UCase(txtVer.Text)
             End If
        Else
             myObjs("BSver").Top = 100000
             myVars.Item("SVer").Value = "N/A"
        End If
    Else
         If chkVer.Value = 1 Then
             If txtVer.Text = "" Or txtVer.Text = "/" Then
                 myObjs("BSver").Top = 100000
                 myVars.Item("SVer").Value = ""
                 myObjs("TSver").Top = 100000
                 myObjs("Text1(2)").Top = 100000
             Else
                 myObjs("TSver").Top = 100000
                 myVars.Item("SVer").Value = UCase(txtVer.Text)
             End If
        Else
            myObjs("BSver").Top = 100000
            myVars.Item("SVer").Value = ""
            myObjs("TSver").Top = 100000
            myObjs("Text1(2)").Top = 100000
        End If
        
    End If
   
   
   
   If Me.chkCN3C.Value = 0 Then
        myObjs("CCC").Top = 100000
   End If
   
   If Me.chkNonFCC.Value = 1 Then
        myObjs("FCC").Top = 100000
   End If
   
    'add by allen.yan 2014/06/07 for the requirement from Shun.Huang
   
   If Me.txtCNCN.Text = "��" Then
        myVars.Item("Company-CN").Value = ""
   Else
        myVars.Item("Company-CN").Value = Trim(Me.txtCNCN.Text)
   End If
   
   
   If Me.txtCNCA.Text = "��" Then
        myVars.Item("Adress-CN").Value = ""
   Else
        myVars.Item("Adress-CN").Value = Trim(Me.txtCNCA.Text)
   End If
   
   
   If Me.txtENCN.Text = "��" Then
        myVars.Item("Company-EN").Value = ""
   Else
        myVars.Item("Company-EN").Value = Trim(Me.txtENCN.Text)
   End If
   
   
   If Me.txtENCA.Text = "��" Then
        myVars.Item("Adress-EN").Value = ""
   Else
        myVars.Item("Adress-EN").Value = Trim(Me.txtENCA.Text)
   End If
   
   
   myVars.Item("CPN").Value = txtCPN.Text
   myVars.Item("EPN").Value = txtEPN.Text
   myVars.Item("Des").Value = txtDes.Text
   myVars.Item("MN").Value = txtMN.Text
   
   myVars.Item("OD").Value = txtOS.Text
   
   'If chkOS.Value = 0 Or txtOS.Text = "/" Then
   '   myObjs("OD").Top = 100000
   '   myVars.Item("OD").Value = ""
   'Else
   '   myVars.Item("OD").Value = txtOS.Text
   'End If
   
   If Trim(txtPO.Text) = "/" Then
        myObjs("Text14(22)").Top = 100000
        myVars.Item("PO").Value = ""
   Else
        myVars.Item("PO").Value = Trim(txtPO.Text)
   End If
   
   myVars.Item("GW").Value = txtGW.Text
   If chkNonCE.Value = 1 Then
      myObjs("CE").Top = 100000
   End If
   If chkNonWEEE.Value = 1 Then
'      myObjs("Trash").Top = 10000
      myObjs("Trash").Top = 100000
   End If
   If chkNonChinaRoHS.Value = 1 Then
      myObjs("China RoHS").Top = 100000
   End If
   If optH3CRoHS.Value = True Then
      myObjs("3COM RoHS").Top = 100000
   End If
   If opt3COMRoHS.Value = True Then
      myObjs("3COM RoHS").Top = 2300
'      myObjs("H3C RoHS").Top = 100000
      myObjs("HUV RoHS").Top = 100000
   End If
   If optNonRoHS.Value = True Then
'      myObjs("H3C RoHS").Top = 100000
      myObjs("HUV RoHS").Top = 100000
      myObjs("3COM RoHS").Top = 100000
   End If
'   If Me.chkNonTurkey.Value = 1 Then
'      myObjs("Turkey RoHS").Top = 100000
'   End If
   myVars.Item("MS").Value = UCase(txtMS.Text)
   'sql = "select ChangNAL from H3C where SN='" & txtSN.Text & "' and ValidTo<='" & Date & "'"
   'rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   'If rec.EOF = True Then
   '   myVars.Item("NAL").Value = UCase(txtNAL.Text)
   'Else
   '   myVars.Item("NAL").Value = rec.Fields(0)
   'End If
   'rec.Close
   
   myVars.Item("NAL").Value = UCase(txtNAL.Text)
   
   If txtHV.Text = "" Or txtHV.Text = "/" Then
      myObjs("BHver(1)").Top = 100000
      myVars.Item("HVer").Value = "N/A"
   Else
      myObjs("THver").Top = 100000
      myVars.Item("HVer").Value = UCase(txtHV.Text)
   End If
   myVars.Item("Remark").Value = UCase(txtRemark.Text)
   'myApp.Visible = True
'   mydoc.CopyToClipboard
   mydoc.PrintLabel 1
   mydoc.FormFeed
   Next
   
   UnloadLppx
    

   cmdCancel_Click
End Sub

Private Sub cmdReturn_Click()
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




Private Sub txtHV_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 13) Then
     txtMS.SetFocus
  End If
End Sub

Private Sub txtMS_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtNAL.SetFocus
  End If
End Sub



Private Sub txtNAL_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtRemark.SetFocus
  End If
End Sub





Private Sub txtPO_GotFocus()
    Me.txtPO.Text = Clipboard.GetText
    Clipboard.Clear
End Sub

Private Sub txtQty_Change()
If txtQty.Text <> "" Then
    If Asc(Right(txtQty.Text, 1)) > 57 Or Asc(Right(txtQty.Text, 1)) < 48 Then
       MsgBox "ֻ���������֣�", vbInformation + vbOKOnly, "���벻��ȷ"
       SendKeys "{backspace}"
       txtQty.SetFocus
       Exit Sub
    End If
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtVer.SetFocus
  End If
End Sub



Private Sub txtRemark_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     cmdPrint_Click
  End If
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      If Len(Trim(txtSN.Text)) < 10 Then
         MsgBox "��Ʒ��ų��Ȳ���С��10!"
         txtSN.SetFocus
         Exit Sub
      End If

      sql = "select top 1 * from revset where model='" & Mid(Trim(txtSN.Text), 3, 8) & "' and firstall<='" & Trim(txtSN.Text) & "' and endall>='" & Trim(txtSN.Text) & "'order by ver desc"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "�˲�Ʒ���δ�ռ��汾!"
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
         txtHV.Text = Trim(rec.Fields(3))
      End If
      rec.Close
      
      
      sql = "select ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, CCC, SalesLocation, Remark,FCC,CNCN,CNCA,ENCN,ENCA from tblHUV where SN='" & Mid(Trim(txtSN.Text), 3, 8) & "' and HV='" & UCase(txtHV.Text) & "'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "�˲�Ʒ����δ��������!"
         txtVer.Text = ""
         txtSN.Text = ""
         txtHV.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
        txtCPN.Text = rec.Fields(3)
        txtEPN.Text = rec.Fields(4)
        txtDes.Text = rec.Fields(5)
        
        Dim psv As String
        psv = rec.Fields(17)
        If UCase(psv) = "N" Then
            chkVer.Value = 0
        Else
            chkVer.Value = 1
        End If
        chkVer.Enabled = False
        
        chkOS.Value = 1
        txtOS.Enabled = True
        'txtMN.Text = rec.Fields(5)
        'txtOS.BackColor = &HC0C0C0
        txtOS.Text = rec.Fields(6)
        txtGW.Text = rec.Fields(7)
        If UCase(Trim(rec.Fields(8))) = "CE" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
        ElseIf rec.Fields(8) = "/" Or rec.Fields(8) = "N/A" Then
           chkCE.Value = 0
           chkNonCE.Value = 1
        End If
        If UCase(Trim(rec.Fields(9))) = "WEEE" Then
           chkWEEE.Value = 1
           chkNonWEEE.Value = 0
        ElseIf rec.Fields(9) = "/" Or rec.Fields(9) = "N/A" Then
           chkWEEE.Value = 0
           chkNonWEEE.Value = 1
        End If
        If UCase(Trim(rec.Fields(10))) = "CHINA ROHS" Then
           chkChinaRoHS.Value = 1
           chkNonChinaRoHS.Value = 0
        ElseIf rec.Fields(10) = "/" Or rec.Fields(10) = "N/A" Then
           chkChinaRoHS.Value = 0
           chkNonChinaRoHS.Value = 1
        End If
'        If UCase(Trim(rec.Fields(11))) = "H3C ROHS" Then
        If UCase(Trim(rec.Fields(11))) = "HUV ROHS" Then
           optH3CRoHS.Value = 1
        ElseIf UCase(Trim(rec.Fields(11))) = "3COM ROHS" Then
           opt3COMRoHS.Value = 1
        ElseIf rec.Fields(11) = "/" Or rec.Fields(11) = "" Or rec.Fields(11) = "N/A" Then
           optNonRoHS.Value = 1
        End If
'        If UCase(Trim(rec.Fields(12))) = "TURKEY ROHS" Then
'            chkTurkey.Value = 1
'            'chkNonTurkey.Value = 0
'        ElseIf rec.Fields(12) = "/" Or rec.Fields(12) = "N/A" Then
'            'chkTurkey.Value = 0
'            chkNonTurkey.Value = 1
'        End If
        
        txtMS.Text = rec.Fields(13)
        txtNAL.Text = rec.Fields(15)
        
        Me.lblMSday.Text = rec.Fields(14)
        Me.lblNALday.Text = rec.Fields(16)
        
         If IsNull(rec.Fields(18)) Then
            MsgBox "û���趨CCC��֤,���ܴ�ӡ!", vbInformation + vbOKOnly, "CCC��֤����Ϊ��"
            txtSN.SetFocus
            Exit Sub
        ElseIf rec.Fields(18) = "Yes" Then
            Me.chkCN3C.Value = 1
            Me.chkNonCN3C.Value = 0
        ElseIf rec.Fields(18) = "No" Then
            Me.chkNonCN3C.Value = 1
            Me.chkCN3C.Value = 0
        End If
        If IsNull(rec.Fields(19)) Then
            MsgBox "û����������,���ܴ�ӡ!", vbInformation + vbOKOnly, "����������Ϊ��"
            txtSN.SetFocus
            Exit Sub
        Else
            Me.txtSalesArea.Text = rec.Fields(19)
        End If
        txtRemark.Text = rec.Fields(20)
        
        If IsNull(rec.Fields("FCC")) Then
            MsgBox "û���趨FCC��֤,���ܴ�ӡ!", vbInformation + vbOKOnly, "FCC��֤����Ϊ��"
            txtSN.SetFocus
            Exit Sub
        ElseIf rec.Fields("FCC") = "Yes" Then
            Me.chkFCC.Value = 1
            Me.chkNonFCC.Value = 0
        ElseIf rec.Fields("FCC") = "No" Then
            Me.chkFCC.Value = 0
            Me.chkNonFCC.Value = 1
        End If
        
        'add by allen.yan 2014/06/09 for shun.huang's requirement
        If IsNull(rec.Fields("CNCN")) Then
            MsgBox "û���趨���Ĺ�˾��,���ܴ�ӡ!", vbInformation + vbOKOnly, "Setting���Ĺ�˾��ֵ����Ϊ��"
            txtSN.SetFocus
            Exit Sub
        Else
            Me.txtCNCN.Text = rec.Fields("CNCN")
        End If
        
        If IsNull(rec.Fields("CNCA")) Then
            MsgBox "û���趨���ĵ�ַ��,���ܴ�ӡ!", vbInformation + vbOKOnly, "Setting���ĵ�ַ��ֵ����Ϊ��"
            txtSN.SetFocus
            Exit Sub
        Else
            Me.txtCNCA.Text = rec.Fields("CNCA")
        End If
        
        If IsNull(rec.Fields("ENCN")) Then
            MsgBox "û���趨Ӣ�Ĺ�˾��,���ܴ�ӡ!", vbInformation + vbOKOnly, "SettingӢ�Ĺ�˾��ֵ����Ϊ��"
            txtSN.SetFocus
            Exit Sub
        Else
            Me.txtENCN.Text = rec.Fields("ENCN")
        End If
        
        If IsNull(rec.Fields("ENCA")) Then
            MsgBox "û���趨Ӣ�ĵ�ַ��,���ܴ�ӡ!", vbInformation + vbOKOnly, "SettingӢ�ĵ�ַ��ֵ����Ϊ��"
            txtSN.SetFocus
            Exit Sub
        Else
            Me.txtENCA.Text = rec.Fields("ENCA")
        End If
      End If
      rec.Close
      txtQty.SetFocus
   End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
'   ����������\\sz-fs01\Public\Manufacture\�����ǩģ��\����20120615     Uniview-����.lab
'����������\\sz-fs01\Public\Manufacture\�����ǩģ��\����20120615    Uniview-����-����.lab
'�Ϻ���Ծ������\\sz-fs01\Public\Manufacture\�����ǩģ��\����20120615   Uniview-����-�Ϻ���Ծ.lab
'Uniview20140320
'   If Me.txtSalesArea.Text = "����" Then
'        Set mydoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\�����ǩģ��\����20120615\Uniview20140420\" & "Uniview-����-NEW.Lab")
'   ElseIf Me.txtSalesArea.Text = "����" Then
'        Set mydoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\�����ǩģ��\����20120615\Uniview20140420\" & "Uniview-����-����.lab")
'   ElseIf Me.txtSalesArea.Text = "�Ϻ���Ծ" Then
'        Set mydoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\�����ǩģ��\����20120615\Uniview20140420\" & "Uniview-����-�Ϻ���Ծ.lab")
'   Else
'      MsgBox "û����������,���ܴ�ӡ!", vbInformation + vbOKOnly, "����������Ϊ��"
'      Exit Sub
'      Unload Me
'   End If
    
   'Set mydoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\�����ǩģ��\����20120615\" & "Uniview-����.lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\��ǩģ��\" & "H3C.lab")
   
    If Me.txtSalesArea.Text = "��Ӣ��" Then
        Set mydoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\�����ǩģ��\����20120615\Uniview20140606\" & "Uniview-����-��Ӣ��.lab")
    ElseIf Me.txtSalesArea.Text = "��Ӣ��" Then
        Set mydoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\�����ǩģ��\����20120615\Uniview20140606\" & "Uniview-����-��Ӣ��.lab")
    ElseIf Me.txtSalesArea.Text = "Ѷ��" Then
        Set mydoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\�����ǩģ��\����20120615\Uniview20140606\" & "Uniview-����-Ѷ��.lab")
    Else
      MsgBox "û�б�ǩ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "����������Ϊ��"
      Exit Sub
      Unload Me
    End If
   
   
   Me.MousePointer = vbDefault
   Set myVars = mydoc.Variables
   Set myObjs = mydoc.DocObjects
End Sub

Private Sub txtVer_GotFocus()
    Me.txtVer.Text = Clipboard.GetText
    Clipboard.Clear
End Sub

Private Sub txtVer_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtSN.SetFocus
  End If
End Sub
