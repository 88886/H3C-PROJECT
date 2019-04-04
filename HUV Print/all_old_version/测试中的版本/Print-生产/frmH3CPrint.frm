VERSION 5.00
Begin VB.Form frmH3CPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H3C Label Print"
   ClientHeight    =   10620
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
   Icon            =   "frmH3CPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10620
   ScaleWidth      =   14130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   4080
      TabIndex        =   35
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   9000
      TabIndex        =   34
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   6480
      TabIndex        =   33
      Top             =   9720
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   120
      TabIndex        =   19
      Top             =   5160
      Width           =   13935
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   7200
         TabIndex        =   44
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optNonRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Height          =   495
         Left            =   13080
         TabIndex        =   42
         Top             =   2640
         Width           =   615
      End
      Begin VB.OptionButton opt3COMRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3COMRoHS"
         Height          =   615
         Left            =   11760
         TabIndex        =   41
         Top             =   2520
         Width           =   1215
      End
      Begin VB.OptionButton optH3CRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "H3CRoHS"
         Height          =   615
         Left            =   10560
         TabIndex        =   40
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtMN 
         Height          =   405
         Left            =   10560
         TabIndex        =   38
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "外尺寸(MM):"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   2160
         Width           =   855
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无CE"
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Height          =   375
         Left            =   10560
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Height          =   375
         Left            =   12480
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtRemark 
         Height          =   405
         Left            =   10560
         TabIndex        =   16
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtNAL 
         Height          =   405
         Left            =   10560
         TabIndex        =   14
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox txtMS 
         Height          =   405
         Left            =   3120
         TabIndex        =   13
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtHV 
         Height          =   405
         Left            =   3120
         TabIndex        =   15
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox txtGW 
         Height          =   405
         Left            =   10560
         TabIndex        =   6
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtOS 
         Height          =   405
         Left            =   3120
         TabIndex        =   5
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtDes 
         Height          =   405
         Left            =   3120
         TabIndex        =   4
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtEPN 
         Height          =   405
         Left            =   10560
         TabIndex        =   3
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtCPN 
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
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   10560
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "数量:"
         Height          =   375
         Left            =   6360
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息RoHS:"
         Height          =   495
         Left            =   8400
         TabIndex        =   39
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品型号:"
         Height          =   375
         Left            =   8400
         TabIndex        =   37
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
         Height          =   375
         Left            =   8400
         TabIndex        =   32
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网许可号:"
         Height          =   375
         Left            =   8400
         TabIndex        =   31
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "执行标准:"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息ChinaRoHS:"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label lblHV 
         BackColor       =   &H00FFFFFF&
         Caption         =   "硬件版本:"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息WEEE:"
         Height          =   375
         Left            =   8400
         TabIndex        =   27
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息CE:"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblGW 
         BackColor       =   &H00FFFFFF&
         Caption         =   "毛重(kg):"
         Height          =   375
         Left            =   8400
         TabIndex        =   25
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(英文):"
         Height          =   375
         Left            =   8400
         TabIndex        =   24
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(中文):"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品起始编码:"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品描述:"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本信息:"
         Height          =   375
         Left            =   8400
         TabIndex        =   20
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      Picture         =   "frmH3CPrint.frx":13652
      ScaleHeight     =   4545
      ScaleWidth      =   13905
      TabIndex        =   18
      Top             =   480
      Width           =   13935
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "H3C 标签："
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
Attribute VB_Name = "frmH3CPrint"
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
   optH3CRoHS.Value = True
   txtMS.Text = ""
   txtNAL.Text = ""
   txtHV.Text = ""
   txtRemark.Text = ""
   txtSN.SetFocus
End Sub

Private Sub cmdPrint_Click()
'On Error Resume Next
   If txtSN.Text = "" Then
      MsgBox "产品编码未输入,不能打印!", vbInformation + vbOKOnly, "未输入产品编码"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If txtQty.Text = "" Then
      MsgBox "数量未输入，不能打印！", vbInformation + vbOKOnly, "未输入数量"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If CInt(txtQty.Text) = 0 Then
      MsgBox "请输入正确的数量！", vbInformation + vbOKOnly, "数量不对"
      txtQty.SetFocus
      Exit Sub
   End If
   
   
   If txtVer.Text = "" Then
      MsgBox "版本未输入,不能打印!", vbInformation + vbOKOnly, "未输入版本"
      txtVer.SetFocus
      Exit Sub
   End If
   
   If txtHV.Text = "" Then
      MsgBox "硬件版本未输入,不能打印!", vbInformation + vbOKOnly, "未输入硬件版本"
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
      myObjs("BSver").Top = 1000000
      myVars.Item("SVer").Value = "N/A"
   Else
      myObjs("TSver").Top = 1000000
      myVars.Item("SVer").Value = UCase(txtVer.Text)
   End If
   myVars.Item("CPN").Value = txtCPN.Text
   myVars.Item("EPN").Value = txtEPN.Text
   myVars.Item("Des").Value = txtDes.Text
   myVars.Item("MN").Value = txtMN.Text
   If chkOS.Value = 0 Or txtOS.Text = "/" Then
      myObjs("OD").Top = 100000
      myVars.Item("OD").Value = ""
   Else
      myVars.Item("OD").Value = txtOS.Text
   End If
   myVars.Item("GW").Value = txtGW.Text
   If chkNonCE.Value = 1 Then
      myObjs("CE").Top = 100000
   End If
   If chkNonWEEE.Value = 1 Then
      myObjs("WEEE").Top = 10000
   End If
   If chkNonChinaRoHS.Value = 1 Then
      myObjs("China RoHS").Top = 100000
   End If
   If optH3CRoHS.Value = True Then
      myObjs("3COM RoHS").Top = 10000
   End If
   If opt3COMRoHS.Value = True Then
      myObjs("3COM RoHS").Top = 2300
      myObjs("H3C RoHS").Top = 100000
   End If
   If optNonRoHS.Value = True Then
      myObjs("H3C RoHS").Top = 100000
      myObjs("3COM RoHS").Top = 100000
   End If
   myVars.Item("MS").Value = UCase(txtMS.Text)
   sql = "select ChangNAL from H3C where SN='" & txtSN.Text & "' and ValidTo<='" & Date & "'"
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   If rec.EOF = True Then
      myVars.Item("NAL").Value = UCase(txtNAL.Text)
   Else
      myVars.Item("NAL").Value = rec.Fields(0)
   End If
   rec.Close
   If txtHV.Text = "" Or txtHV.Text = "/" Then
      myObjs("BHver").Top = 1000000
      myVars.Item("HVer").Value = "N/A"
   Else
      myObjs("THver").Top = 1000000
      myVars.Item("HVer").Value = UCase(txtHV.Text)
   End If
   myVars.Item("Remark").Value = UCase(txtRemark.Text)
   'myApp.Visible = True
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

Private Sub txtQty_Change()
If txtQty.Text <> "" Then
    If Asc(Right(txtQty.Text, 1)) > 57 Or Asc(Right(txtQty.Text, 1)) < 48 Then
       MsgBox "只能输入数字！", vbInformation + vbOKOnly, "输入不正确"
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
         MsgBox "产品序号长度不能小于10!"
         txtSN.SetFocus
         Exit Sub
      End If
'      sql = "select ver from version where SN='" & txtSN.Text & "'"
'      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
'      If rec.EOF = True Then
'         MsgBox "此产品序号未收集版本!"
'         txtSN.Text = ""
'         txtSN.SetFocus
'         rec.Close
'         Exit Sub
'      Else
'        Dim rcd As New ADODB.Recordset
'        sql = "select max(testtime) from version where model='" & Mid(txtSN.Text, 3, 8) & "'"
'        rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
'        If rcd.EOF = True Then
'           MsgBox "此产品序号未收集版本!"
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
'             MsgBox "此产品序号未收集版本!"
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
      sql = "select * from H3C where SN='" & Mid(txtSN.Text, 3, 8) & "'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "此产品编码未进行设置!"
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
        txtMN.Text = rec.Fields(5)
        'txtOS.BackColor = &HC0C0C0
        txtOS.Text = rec.Fields(6)
        txtGW.Text = rec.Fields(7)
        If UCase(Trim(rec.Fields(8))) = "CE" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
        ElseIf rec.Fields(8) = "/" Then
           chkCE.Value = 0
           chkNonCE.Value = 1
        End If
        If UCase(Trim(rec.Fields(9))) = "WEEE" Then
           chkWEEE.Value = 1
           chkNonWEEE.Value = 0
        ElseIf rec.Fields(9) = "/" Then
           chkWEEE.Value = 0
           chkNonWEEE.Value = 1
        End If
        If UCase(Trim(rec.Fields(10))) = "CHINA ROHS" Then
           chkChinaRoHS.Value = 1
           chkNonChinaRoHS.Value = 0
        ElseIf rec.Fields(10) = "/" Then
           chkChinaRoHS.Value = 0
           chkNonChinaRoHS.Value = 1
        End If
        If UCase(Trim(rec.Fields(11))) = "H3C ROHS" Then
           optH3CRoHS.Value = 1
        ElseIf UCase(Trim(rec.Fields(11))) = "3COM ROHS" Then
           opt3COMRoHS.Value = 1
        ElseIf rec.Fields(11) = "/" Or rec.Fields(11) = "" Then
           optNonRoHS.Value = 1
        End If
        txtMS.Text = rec.Fields(12)
        txtNAL.Text = rec.Fields(13)
'        dtpValidFrom.Value = rec.Fields(14)
'        dtpValidTo.Value = rec.Fields(15)
'        txtChangNAL.Text = rec.Fields(16)
        'txtHV.Text = rec.Fields(17)
        txtRemark.Text = rec.Fields(18)
      End If
      rec.Close
      txtQty.SetFocus
   End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set mydoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\逆向标签模板\" & "H3C-备件.lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = mydoc.Variables
   Set myObjs = mydoc.DocObjects
End Sub

Private Sub txtVer_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtHV.SetFocus
  End If
End Sub
