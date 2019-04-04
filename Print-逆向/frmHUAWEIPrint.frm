VERSION 5.00
Begin VB.Form frmHUAWEIPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HUAWEI Label Print"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHUAWEIPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   615
      Left            =   7680
      TabIndex        =   19
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   615
      Left            =   5160
      TabIndex        =   18
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(Print)&p"
      Height          =   615
      Left            =   2880
      TabIndex        =   17
      Top             =   8760
      Width           =   1695
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   11415
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   6120
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ߴ�(MM):"
         Height          =   375
         Left            =   6360
         TabIndex        =   36
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtVer 
         Height          =   405
         Left            =   8400
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtCPN 
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtEPN 
         Height          =   405
         Left            =   8400
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtDes 
         Height          =   405
         Left            =   2400
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtOS 
         Height          =   405
         Left            =   8400
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtGW 
         Height          =   405
         Left            =   2400
         TabIndex        =   6
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtHV 
         Height          =   405
         Left            =   2400
         TabIndex        =   15
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox txtMS 
         Height          =   405
         Left            =   2400
         TabIndex        =   13
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtNAL 
         Height          =   405
         Left            =   8400
         TabIndex        =   14
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtRemark 
         Height          =   405
         Left            =   8400
         TabIndex        =   16
         Top             =   3120
         Width           =   2895
      End
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NonRoHS"
         Height          =   375
         Left            =   9600
         TabIndex        =   12
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RoHS"
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��CE"
         Height          =   375
         Left            =   9600
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox chkCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE"
         Height          =   375
         Left            =   8400
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����:"
         Height          =   375
         Left            =   5280
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�汾��Ϣ:"
         Height          =   375
         Left            =   6960
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����(����):"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����(Ӣ��):"
         Height          =   375
         Left            =   6240
         TabIndex        =   31
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblGW 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ë��(kg):"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��֤��ϢCE:"
         Height          =   375
         Left            =   6240
         TabIndex        =   29
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��֤��ϢWEEE:"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblHV 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ӳ���汾:"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label lblRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��֤��ϢROSH:"
         Height          =   375
         Left            =   6240
         TabIndex        =   26
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ִ�б�׼:"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������ɺ�:"
         Height          =   375
         Left            =   6240
         TabIndex        =   24
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ע:"
         Height          =   375
         Left            =   6240
         TabIndex        =   23
         Top             =   3120
         Width           =   2175
      End
   End
   Begin VB.PictureBox picHUAWEI 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      Picture         =   "frmHUAWEIPrint.frx":13652
      ScaleHeight     =   4545
      ScaleWidth      =   11385
      TabIndex        =   21
      Top             =   360
      Width           =   11415
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HUAWEI��ǩ:"
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
      TabIndex        =   20
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmHUAWEIPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects

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

Private Sub chkNonRoHS_Click()
   If chkNonRoHS.Value = 1 Then
      chkRoHS.Value = 0
   Else
      chkRoHS.Value = 1
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

Private Sub chkRoHS_Click()
   If chkRoHS.Value = 1 Then
      chkNonRoHS.Value = 0
   Else
      chkNonRoHS.Value = 1
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
   txtSN.Text = ""
   txtVer.Text = ""
   txtCPN.Text = ""
   txtEPN.Text = ""
   txtDes.Text = ""
   txtOS.Text = ""
   txtGW.Text = ""
   txtQty.Text = ""
'   chkCE.Value = 0
   chkNonCE.Value = 0
'   chkWEEE.Value = 0
   chkNonWEEE.Value = 0
'   chkRoHS.Value = 0
   chkNonRoHS.Value = 0
   txtMS.Text = ""
   txtNAL.Text = ""
   txtHV.Text = ""
   txtRemark.Text = ""
   txtSN.SetFocus
End Sub

Private Sub cmdPrint_Click()
   If txtSN.Text = "" Then
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
   
   
   If txtVer.Text = "" Then
      MsgBox "�汾δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ����汾"
      txtVer.SetFocus
      Exit Sub
   End If
   
   If txtHV.Text = "" Then
      MsgBox "Ӳ���汾δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ����Ӳ���汾"
      txtHV.SetFocus
      Exit Sub
   End If
   
   
    Dim i, qty As Integer
   Dim leftstr, rightstr, str As String
   qty = CInt(txtQty.Text)
   leftstr = UCase(Left(txtSN.Text, 14))
   rightstr = Right(txtSN.Text, 6)
    OpenLppx
   For i = 0 To qty - 1
      str = leftstr & Right("000000" & CStr(CInt(rightstr) + i), 6)
   
  
   myVars.Item("SN").Value = str
   myVars.Item("Item").Value = UCase(Mid(txtSN.Text, 3, 8))
   If txtVer.Text = "" Or txtVer.Text = "/" Then
      myObjs("BSver").Top = 10000
      myVars.Item("SVer").Value = txtVer.Text
   Else
      myObjs("TSver").Top = 10000
      myVars.Item("SVer").Value = UCase(txtVer.Text)
   End If
   myVars.Item("PN").Value = txtCPN.Text & "/" & txtEPN.Text
   myVars.Item("Des").Value = txtDes.Text
   myVars.Item("OD").Value = txtOS.Text
'   If txtOS.Text = "" Then
'      myObjs("OD").Top = 10000
'      myVars.Item("OD").Value = txtOS.Text
'   Else
'      myVars.Item("OD").Value = txtOS.Text
'   End If
   myVars.Item("GW").Value = txtGW.Text
   If chkNonCE.Value = 1 Then
      myObjs("CE").Top = 10000
   End If
   If chkNonWEEE.Value = 1 Then
      myObjs("WEEE").Top = 10000
   End If
   If chkNonRoHS.Value = 1 Then
      myObjs("RoHS").Top = 10000
   End If
   myVars.Item("MS").Value = UCase(txtMS.Text)
   If txtNAL.Text = "" Or txtNAL.Text = "/" Then
      myObjs("NAL").Top = 10000
      myVars.Item("NAL").Value = ""
   Else
     sql = "select ChangNAL from HUAWEI where SN='" & txtSN.Text & "' and ValidTo<='" & Date & "'"
     rec.Open sql, conn, adOpenKeyset, adLockOptimistic
     If rec.EOF = True Then
        myVars.Item("NAL").Value = UCase(txtNAL.Text)
     Else
        myVars.Item("NAL").Value = rec.Fields(0)
     End If
     rec.Close
   End If
   If txtHV.Text = "" Or txtHV.Text = "/" Then
      myObjs("BHver(1)").Top = 10000
      myVars.Item("HVer").Value = txtHV.Text
   Else
      myObjs("THver").Top = 10000
      myVars.Item("HVer").Value = UCase(txtHV.Text)
   End If
   myVars.Item("Remark").Value = UCase(txtRemark.Text)
   'myApp.Visible = True
   myDoc.PrintLabel 1
   myDoc.FormFeed
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
      If Len(txtSN.Text) < 10 Then
         MsgBox "��Ʒ��ų��Ȳ���С��10!"
         txtSN.SetFocus
         Exit Sub
      End If
'      sql = "select ver from version where SN='" & txtSN.Text & "'"
'      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
'      If rec.EOF = True Then
'         MsgBox "�˲�Ʒ���δ�ռ��汾!"
'         txtSN.Text = ""
'         txtSN.SetFocus
'         rec.Close
'         Exit Sub
'      Else
'        Dim rcd As New ADODB.Recordset
'        sql = "select max(testtime) from version where model='" & Mid(txtSN.Text, 3, 8) & "'"
'        rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
'        If rec.EOF = True Then
'           MsgBox "�˲�Ʒ���δ�ռ��汾!"
'           txtSN.Text = ""
'           txtSN.SetFocus
'           rcd.Close
'           rec.Close
'           Exit Sub
'        Else
'          Dim rs As New ADODB.Recordset
'          sql = "select ver from version where testtime='" & rcd.Fields(0) & "' and model='" & Mid(txtSN.Text, 3, 8) & "'"
'          rs.Open sql, conn, adOpenKeyset, adLockOptimistic
'          If rs.EOF = False Then
'             txtVer.Text = rs.Fields(0)
'          Else
'             MsgBox "�˲�Ʒ���δ�ռ��汾!"
'             txtSN.Text = ""
'             txtSN.SetFocus
'             rs.Close
'             rcd.Close
'             rec.Close
'             Exit Sub
'          End If
'          rs.Close
'        End If
'        rcd.Close
'      End If
'      rec.Close
      sql = "select * from HUAWEI where SN='" & Mid(txtSN.Text, 3, 8) & "'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "�˲�Ʒ����δ��������!"
         txtVer.Text = ""
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
        txtCPN.Text = rec.Fields(2)
        txtEPN.Text = rec.Fields(3)
        txtDes.Text = rec.Fields(4)
        chkOS.Value = 1
        txtOS.Enabled = True
        txtOS.BackColor = &H80000005
        txtOS.Text = rec.Fields(5)
        txtGW.Text = rec.Fields(6)
        If UCase(Trim(rec.Fields(7))) = "CE" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
        ElseIf rec.Fields(7) = "/" Then
           chkCE.Value = 0
           chkNonCE.Value = 1
        End If
        If UCase(Trim(rec.Fields(8))) = "WEEE" Then
           chkWEEE.Value = 1
           chkNonWEEE.Value = 0
        ElseIf rec.Fields(8) = "/" Then
           chkWEEE.Value = 0
           chkNonWEEE.Value = 1
        End If
        If UCase(Trim(rec.Fields(9))) = "ROHS" Then
           chkRoHS.Value = 1
           chkNonRoHS.Value = 0
        ElseIf rec.Fields(9) = "/" Then
           chkRoHS.Value = 0
           chkNonRoHS.Value = 1
        End If
        txtMS.Text = rec.Fields(10)
        txtNAL.Text = rec.Fields(11)
'        dtpValidFrom.Value = rec.Fields(12)
'        dtpValidTo.Value = rec.Fields(13)
'        txtChangNAL.Text = rec.Fields(14)
       ' txtHV.Text = rec.Fields(15)
        txtRemark.Text = rec.Fields(16)
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
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HUAWEI.lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\��ǩģ��\" & "HUAWEI.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub txtVer_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtHV.SetFocus
  End If
End Sub
