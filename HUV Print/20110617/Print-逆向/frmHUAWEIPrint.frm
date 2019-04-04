VERSION 5.00
Begin VB.Form frmHUAWEIPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HUAWEI Label Print"
   ClientHeight    =   10170
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
   ScaleHeight     =   10170
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkChinarosh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "China Rosh"
      Height          =   375
      Left            =   8520
      TabIndex        =   40
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   7560
      TabIndex        =   19
      Top             =   9480
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   5160
      TabIndex        =   18
      Top             =   9480
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(Print)&p"
      Height          =   615
      Left            =   2880
      TabIndex        =   17
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   11415
      Begin VB.TextBox lblNALday 
         Height          =   450
         Left            =   6600
         TabIndex        =   43
         Top             =   3600
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox lblMSday 
         Height          =   450
         Left            =   6240
         TabIndex        =   42
         Top             =   3600
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CheckBox chkVer 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   41
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   6000
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "外尺寸(MM):"
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
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtEPN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   8400
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtDes 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtOS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   8400
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtGW 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   6
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtHV 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   15
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox txtMS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   13
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtNAL 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   8400
         TabIndex        =   14
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   16
         Top             =   3600
         Width           =   2775
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
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无CE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox chkCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8400
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "China Rosh:"
         Height          =   375
         Left            =   6360
         TabIndex        =   39
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "数量:"
         Height          =   375
         Left            =   5280
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本信息:"
         Height          =   375
         Left            =   6960
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品描述:"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(中文):"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(英文):"
         Height          =   375
         Left            =   6240
         TabIndex        =   31
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblGW 
         BackColor       =   &H00FFFFFF&
         Caption         =   "毛重(kg):"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息CE:"
         Height          =   375
         Left            =   6240
         TabIndex        =   29
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息WEEE:"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblHV 
         BackColor       =   &H00FFFFFF&
         Caption         =   "硬件版本:"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label lblRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息ROSH:"
         Height          =   375
         Left            =   6240
         TabIndex        =   26
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "执行标准:"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网许可号:"
         Height          =   375
         Left            =   6240
         TabIndex        =   24
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   3600
         Width           =   2175
      End
   End
   Begin VB.PictureBox picHUAWEI 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "HUAWEI标签:"
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
   chkChinaRosh.Value = 0
   txtMS.Text = ""
   txtNAL.Text = ""
   txtHV.Text = ""
   txtRemark.Text = ""
   
   Me.lblMSday.Text = ""
   Me.lblNALday.Text = ""
   
   txtSN.SetFocus
End Sub

Private Sub cmdPrint_Click()
   If Trim(txtSN.Text) = "" Then
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
   
   If Trim(txtGW.Text) = "" Then
      MsgBox "产品重量未带出,不能打印!", vbInformation + vbOKOnly, "未带出毛重"
      txtGW.SetFocus
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
   
  If DateDiff("d", CDate(Trim(Me.lblMSday.Text)), Date) >= 0 Then
      MsgBox "制造标准有效期已过期,不能打印!", vbInformation + vbOKOnly, "制造标准过期"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If DateDiff("d", CDate(Trim(Me.lblNALday.Text)), Date) >= 0 Then
      MsgBox "进网许可有效期已过期,不能打印!", vbInformation + vbOKOnly, "进网许可过期"
      txtSN.SetFocus
      Exit Sub
   End If
   
   
    Dim i, qty As Integer
   Dim leftstr, rightstr, str As String
   qty = CInt(txtQty.Text)
   leftstr = UCase(Left(Trim(txtSN.Text), 14))
   rightstr = Right(Trim(txtSN.Text), 6)
    OpenLppx
   For i = 0 To qty - 1
      str = leftstr & Right("000000" & CStr(CInt(rightstr) + i), 6)
   
  
   myVars.Item("SN").Value = str
   myVars.Item("Item").Value = UCase(Mid(Trim(txtSN.Text), 3, 8))

   If chkVer.Value = 1 Then
        If txtVer.Text = "" Or txtVer.Text = "/" Then
            myObjs("BSver").Top = 100000
            myVars.Item("SVer").Value = "N/A" 'txtVer.Text
        Else
            myObjs("TSver").Top = 100000
            myVars.Item("SVer").Value = UCase(txtVer.Text)
        End If
    Else
        myObjs("BSver").Top = 100000
        myVars.Item("SVer").Value = "N/A"
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
   If chkChinaRosh = 0 Then
      myObjs("China RoHS").Top = 100000
   End If
   If chkNonCE.Value = 1 Then
      myObjs("CE").Top = 100000
   End If
   If chkNonWEEE.Value = 1 Then
      myObjs("Trash").Top = 100000
   End If
   If chkNonRoHS.Value = 1 Then
      myObjs("RoHS").Top = 100000
   End If
   myVars.Item("MS").Value = UCase(txtMS.Text)
   If txtNAL.Text = "" Or txtNAL.Text = "/" Or txtNAL.Text = "N/A" Then
      myObjs("NAL").Top = 100000
      myVars.Item("NAL").Value = ""
   Else
     'sql = "select ChangNAL from HUAWEI where SN='" & txtSN.Text & "' and ValidTo<='" & Date & "'"
     'rec.Open sql, conn, adOpenKeyset, adLockOptimistic
     'If rec.EOF = True Then
     '   myVars.Item("NAL").Value = UCase(txtNAL.Text)
     'Else
     '   myVars.Item("NAL").Value = rec.Fields(0)
     'End If
     'rec.Close
     
     myVars.Item("NAL").Value = UCase(txtNAL.Text)
   End If
   If txtHV.Text = "" Or txtHV.Text = "/" Then
      myObjs("BHVer").Top = 100000
      myVars.Item("HVer").Value = "N/A" 'txtHV.Text
   Else
      myObjs("THver").Top = 100000
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
      If Len(Trim(txtSN.Text)) < 10 Then
         MsgBox "产品序号长度不能小于10!"
         txtSN.SetFocus
         Exit Sub
      End If

       
        sql = "select top 1 * from revset where model='" & Mid(Trim(txtSN.Text), 3, 8) & "' and firstall<='" & Trim(txtSN.Text) & "' and endall>='" & Trim(txtSN.Text) & "'order by ver desc"
        rec.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rec.EOF = True Then
            MsgBox "此产品序号未收集版本!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
        Else
            txtHV.Text = Trim(rec.Fields(3))
        End If
        rec.Close
        
      sql = "select * from tblHUAWEI where SN='" & Mid(Trim(txtSN.Text), 3, 8) & "' and HV='" & UCase(txtHV.Text) & "'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "此产品编码未进行设置!"
         txtVer.Text = ""
          txtHV.Text = ""
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
           If UCase(Trim(rec.Fields(10))) = "YES" Then
                chkChinaRosh.Value = 1
           Else
                chkChinaRosh.Value = 0
           End If
        txtCPN.Text = rec.Fields(3)
        txtEPN.Text = rec.Fields(4)
        txtDes.Text = rec.Fields(5)
        
        Dim psv As String
        psv = rec.Fields(16)
        If UCase(psv) = "N" Then
            chkVer.Value = 0
        Else
            chkVer.Value = 1
        End If
        chkVer.Enabled = False
        
        chkOS.Value = 1
        txtOS.Enabled = True
        txtOS.BackColor = &H80000005
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
           chkChinaRosh.Value = 1
        ElseIf rec.Fields(10) = "/" Or rec.Fields(10) = "N/A" Then
           chkChinaRosh.Value = 0
        End If
        
         If UCase(Trim(rec.Fields(11))) = "HUAWEI ROHS" Or UCase(Trim(rec.Fields(11))) = "ROHS" Then
           chkRoHS.Value = 1
           chkNonRoHS.Value = 0
        ElseIf rec.Fields(11) = "/" Or rec.Fields(11) = "N/A" Then
           chkRoHS.Value = 0
           chkNonRoHS.Value = 1
        End If
        If Trim(rec.Fields(12)) = "/" Then
            txtMS.Text = "N/A"
        Else
            txtMS.Text = rec.Fields(12)
        End If
        
        txtNAL.Text = rec.Fields(14)
        
        If rec.Fields(17) = "/" Then
            txtRemark.Text = "N/A"
        Else
            txtRemark.Text = rec.Fields(17)
        End If
        
        Me.lblMSday.Text = rec.Fields(13)
        Me.lblNALday.Text = rec.Fields(15)
        
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
   Set mydoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\逆向标签模板\" & "HUAWEI-逆向备件.lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "HUAWEI.lab")
   Me.MousePointer = vbDefault
   Set myVars = mydoc.Variables
   Set myObjs = mydoc.DocObjects
End Sub

Private Sub txtVer_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtSN.SetFocus
  End If
End Sub
