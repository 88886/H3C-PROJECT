VERSION 5.00
Begin VB.Form frmH3C_3COMPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H3C-3COM Label Print"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   14160
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmH3C_3COMPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   14160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   3480
      Picture         =   "frmH3C_3COMPrint.frx":13652
      TabIndex        =   17
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   8400
      TabIndex        =   19
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   5880
      TabIndex        =   18
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      TabIndex        =   22
      Top             =   5160
      Width           =   13935
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "外尺寸(MM):"
         Height          =   375
         Left            =   8640
         TabIndex        =   36
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox chkCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE"
         Height          =   375
         Left            =   10680
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无CE"
         Height          =   375
         Left            =   12000
         TabIndex        =   8
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox chk3COMRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RoHS"
         Height          =   375
         Left            =   10680
         TabIndex        =   11
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox chkNon3COMRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NonRoHS"
         Height          =   375
         Left            =   12000
         TabIndex        =   12
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtRemark 
         Height          =   405
         Left            =   10680
         TabIndex        =   16
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox txtNAL 
         Height          =   405
         Left            =   10680
         TabIndex        =   14
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox txtMS 
         Height          =   405
         Left            =   2400
         TabIndex        =   13
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtHV 
         Height          =   405
         Left            =   2400
         TabIndex        =   15
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox txtGW 
         Height          =   405
         Left            =   2400
         TabIndex        =   6
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtOS 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   10680
         TabIndex        =   5
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtDes 
         Height          =   405
         Left            =   2400
         TabIndex        =   4
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtEPN 
         Height          =   405
         Left            =   10680
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtCPN 
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtVer 
         Height          =   405
         Left            =   10680
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
         Height          =   375
         Left            =   8520
         TabIndex        =   35
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblNAL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网许可号:"
         Height          =   375
         Left            =   8520
         TabIndex        =   34
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "执行标准:"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3COM ROSH:"
         Height          =   375
         Left            =   8520
         TabIndex        =   32
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lblHV 
         BackColor       =   &H00FFFFFF&
         Caption         =   "硬件版本:"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "China RoHS:"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息CE:"
         Height          =   375
         Left            =   8520
         TabIndex        =   29
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblGW 
         BackColor       =   &H00FFFFFF&
         Caption         =   "毛重(kg):"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(英文):"
         Height          =   375
         Left            =   8520
         TabIndex        =   27
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(中文):"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品描述:"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本信息:"
         Height          =   375
         Left            =   8520
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      Picture         =   "frmH3C_3COMPrint.frx":26CA4
      ScaleHeight     =   4665
      ScaleWidth      =   13905
      TabIndex        =   21
      Top             =   360
      Width           =   13935
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "H3C-3COM 标签："
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
      Width           =   2055
   End
End
Attribute VB_Name = "frmH3C_3COMPrint"
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

Private Sub chkNon3COMRoHS_Click()
   If chkNon3COMRoHS.Value = 1 Then
      chk3COMRoHS.Value = 0
   Else
      chk3COMRoHS.Value = 1
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

Private Sub chkOS_Click()
   If chkOS.Value = 1 Then
      txtOS.Enabled = True
      txtOS.BackColor = &H80000005
   Else
      txtOS.Enabled = False
      txtOS.BackColor = &HC0C0C0
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
'   chkCE.Value = 0
   chkNonCE.Value = 0
'   chkChinaRoHS.Value = 0
   chkNonChinaRoHS.Value = 0
'   chk3COMRoHS.Value = 0
   chkNon3COMRoHS.Value = 0
   txtMS.Text = ""
   txtNAL.Text = ""
   txtHV.Text = ""
   txtRemark.Text = ""
   txtSN.SetFocus
End Sub

Private Sub cmdPrint_Click()
   If Trim(txtSN.Text) = "" Then
      MsgBox "产品编码未输入,不能打印!", vbInformation + vbOKOnly, "未输入产品编码"
      txtSN.SetFocus
      Exit Sub
   End If
   If txtVer.Text = "" Then
      MsgBox "版本未带出,不能打印,请重新输入产品编码!", vbInformation + vbOKOnly, "未带出版本"
      txtSN.SetFocus
      Exit Sub
   End If
   If Trim(txtGW.Text) = "" Then
      MsgBox "产品重量未带出,不能打印!", vbInformation + vbOKOnly, "未带出毛重"
      txtGW.SetFocus
      Exit Sub
   End If
   
   OpenLppx
   myVars.Item("SN").Value = Trim(txtSN.Text)
   myVars.Item("Item").Value = Mid(Trim(txtSN.Text), 3, 8)
   If txtVer.Text = "" Or txtVer.Text = "/" Then
      myObjs("BSver").Top = 10000
      myVars.Item("SVer").Value = txtVer.Text
   Else
      myObjs("TSver").Top = 10000
      myVars.Item("SVer").Value = txtVer.Text
   End If
   myVars.Item("CPN").Value = txtCPN.Text
   myVars.Item("EPN").Value = txtEPN.Text
   myVars.Item("Des").Value = txtDes.Text
   If chkOS.Value = 0 Or txtOS.Text = "/" Then
      myObjs("OD").Top = 10000
      myVars.Item("OD").Value = ""
   Else
      myVars.Item("OD").Value = txtOS.Text
   End If
   myVars.Item("GW").Value = txtGW.Text
   If chkNonCE.Value = 1 Then
      myObjs("CE").Top = 100000
   End If
   If chkNonChinaRoHS.Value = 1 Then
      myObjs("ChinaRoHS").Top = 10000
   End If
   If chkNon3COMRoHS.Value = 1 Then
      myObjs("3COMRoHS").Top = 10000
   End If
   myVars.Item("MS").Value = txtMS.Text
   sql = "select ChangNAL from H3C_3COM where SN='" & Trim(txtSN.Text) & "' and ValidTo<='" & Date & "'"
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   If rec.EOF = True Then
      myVars.Item("NAL").Value = txtNAL.Text
   Else
      myVars.Item("NAL").Value = rec.Fields(0)
   End If
   rec.Close
   If txtHV.Text = "" Or txtHV.Text = "/" Then
      myObjs("BHver").Top = 10000
      myVars.Item("HVer").Value = txtHV.Text
   Else
      myObjs("THver").Top = 10000
      myVars.Item("HVer").Value = txtHV.Text
   End If
   myVars.Item("Remark").Value = txtRemark.Text
   myApp.Visible = True
   mydoc.PrintLabel 1
   mydoc.FormFeed
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
   UnloadLppx
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      If Len(Trim(txtSN.Text)) < 10 Then
         MsgBox "产品序号长度不能小于10!"
         txtSN.SetFocus
         Exit Sub
      End If
      sql = "select ver from version where SN='" & Trim(txtSN.Text) & "'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "此产品序号未收集版本!"
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
        Dim rcd As New ADODB.Recordset
        sql = "select max(testtime) from version where model='" & Mid(Trim(txtSN.Text), 3, 8) & "'"
        rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rec.EOF = True Then
           MsgBox "此产品序号未收集版本!"
           txtSN.Text = ""
           txtSN.SetFocus
           rcd.Close
           rec.Close
           Exit Sub
        Else
          Dim rs As New ADODB.Recordset
          sql = "select ver from version where testtime='" & rcd.Fields(0) & "' and model='" & Mid(Trim(txtSN.Text), 3, 8) & "'"
          rs.Open sql, conn, adOpenKeyset, adLockOptimistic
          If rs.EOF = False Then
             txtVer.Text = rs.Fields(0)
          Else
             MsgBox "此产品序号未收集版本!"
             txtSN.Text = ""
             txtSN.SetFocus
             rs.Close
             rcd.Close
             rec.Close
             Exit Sub
          End If
          rs.Close
        End If
        rcd.Close
      End If
      rec.Close
      sql = "select * from H3C_3COM where SN='" & Mid(Trim(txtSN.Text), 3, 8) & "'"
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
        chkOS.Value = 0
        txtOS.Enabled = False
        txtOS.BackColor = &HC0C0C0
        txtOS.Text = rec.Fields(5)
        txtGW.Text = rec.Fields(6)
        If UCase(Trim(rec.Fields(7))) = "CE" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
        ElseIf rec.Fields(7) = "/" Then
           chkCE.Value = 0
           chkNonCE.Value = 1
        End If
        If UCase(Trim(rec.Fields(8))) = "CHINA ROHS" Then
           chkChinaRoHS.Value = 1
           chkNonChinaRoHS.Value = 0
        ElseIf rec.Fields(8) = "/" Then
           chkChinaRoHS.Value = 0
           chkNonChinaRoHS.Value = 1
        End If
        If UCase(Trim(rec.Fields(9))) = "3COM ROHS" Then
           chk3COMRoHS.Value = 1
           chkNon3COMRoHS.Value = 0
        ElseIf rec.Fields(9) = "/" Then
           chk3COMRoHS.Value = 0
           chkNon3COMRoHS.Value = 1
        End If
        txtMS.Text = rec.Fields(10)
        txtNAL.Text = rec.Fields(11)
'        dtpValidFrom.Value = rec.Fields(12)
'        dtpValidTo.Value = rec.Fields(13)
'        txtChangNAL.Text = rec.Fields(14)
        txtHV.Text = rec.Fields(15)
        txtRemark.Text = rec.Fields(16)
      End If
      rec.Close
   End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   'Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "H3C-3COM.lab")
   Set mydoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C-3COM.lab")
   Me.MousePointer = vbDefault
   Set myVars = mydoc.Variables
   Set myObjs = mydoc.DocObjects
End Sub
