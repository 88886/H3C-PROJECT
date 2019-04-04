VERSION 5.00
Begin VB.Form frmHUAWEIPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HUAWEI Label Print"
   ClientHeight    =   10365
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
   ScaleHeight     =   10365
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkchinarosh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ChinaRosh"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8520
      TabIndex        =   38
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   7680
      TabIndex        =   19
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   5160
      TabIndex        =   18
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(Print)&p"
      Height          =   615
      Left            =   2760
      TabIndex        =   17
      Top             =   9720
      Width           =   1695
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   11415
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "外尺寸(MM):"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   36
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtVer 
         Enabled         =   0   'False
         Height          =   405
         Left            =   8400
         Locked          =   -1  'True
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
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtEPN 
         Enabled         =   0   'False
         Height          =   405
         Left            =   8400
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtDes 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtOS 
         Enabled         =   0   'False
         Height          =   405
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtGW 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   6
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtHV 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox txtMS 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   13
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtNAL 
         Enabled         =   0   'False
         Height          =   405
         Left            =   8400
         TabIndex        =   14
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox txtRemark 
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
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         TabIndex        =   12
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RoHS"
         Enabled         =   0   'False
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
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "China Rosh:"
         Height          =   495
         Left            =   6360
         TabIndex        =   37
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本信息:"
         Height          =   375
         Left            =   6240
         TabIndex        =   35
         Top             =   240
         Width           =   1695
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
         Width           =   2055
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
      MsgBox "产品编码未输入,不能打印!", vbInformation + vbOKOnly, "未输入产品编码"
      txtSN.SetFocus
      Exit Sub
   End If
   If txtVer.Text = "" Then
      MsgBox "版本未带出,不能打印,请重新输入产品编码!", vbInformation + vbOKOnly, "未带出版本"
      txtSN.SetFocus
      Exit Sub
   End If
   'Verify hardware & software version before printing      modified by Jimmy Sun 06.14.2010
   If txtHV.Text = "" Then
      MsgBox "产品没有硬件版本,不能打印!", vbInformation + vbOKOnly, "没有硬件版本"
      txtHV.SetFocus
      Exit Sub
   End If
   If txtVer.Text = "" Then
      MsgBox "产品没有软件版本,不能打印!", vbInformation + vbOKOnly, "没有软件版本"
      txtVer.SetFocus
      Exit Sub
   End If
   OpenLppx
   myVars.Item("SN").Value = UCase(txtSN.Text)
   myVars.Item("Item").Value = Mid(UCase(txtSN.Text), 3, 8)
   If UCase(txtVer.Text) = "" Or txtVer.Text = "/" Then
      myObjs("BSver").Top = 10000
      myVars.Item("SVer").Value = "N/A"
   Else
      myObjs("TSver").Top = 10000
      myVars.Item("SVer").Value = UCase(txtVer.Text)
   End If
   myVars.Item("PN").Value = UCase(txtCPN.Text) & "/" & txtEPN.Text
   'myVars.Item("Des").Value = UCase(txtDes.Text)
    myVars.Item("Des").Value = txtDes.Text
   'myVars.Item("OD").Value = UCase(txtOS.Text)          //modified by Jimmy Sun 06.11.2010
    myVars.Item("OD").Value = txtOS.Text
'   If txtOS.Text = "" Then
'      myObjs("OD").Top = 10000
'      myVars.Item("OD").Value = txtOS.Text
'   Else
'      myVars.Item("OD").Value = txtOS.Text
'   End If
' myVars.Item("GW").Value = UCase(txtGW.Text)
    myVars.Item("GW").Value = txtGW.Text
   If chkchinarosh.Value = 0 Then
      myObjs("China RoHS(1)").Top = 10000
   End If
   If chkNonCE.Value = 1 Then
      myObjs("CE").Top = 10000
   End If
   If chkNonWEEE.Value = 1 Then
      myObjs("WEEE").Top = 10000
   End If
   If chkNonRoHS.Value = 1 Then
      myObjs("RoHS").Top = 10000
   End If
   myVars.Item("MS").Value = txtMS.Text
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
      myObjs("BHver").Top = 10000
      myVars.Item("HVer").Value = "N/A" 'UCase(txtHV.Text)
   Else
      myObjs("THver").Top = 10000
      myVars.Item("HVer").Value = UCase(txtHV.Text)
   End If
   myVars.Item("Remark").Value = UCase(txtRemark.Text)
   'myApp.Visible = True
   myDoc.PrintLabel 1
   myDoc.FormFeed
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

Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      If Len(txtSN.Text) < 10 Then
         MsgBox "产品序号长度不能小于10!"
         txtSN.SetFocus
         Exit Sub
      End If
      sql = "select ver from version where SN='" & txtSN.Text & "'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "此产品序号未收集版本!"
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
        Dim rcd As New ADODB.Recordset
        sql = "select max(testtime) from version where sn='" & Trim(txtSN.Text) & "'"
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
          sql = "select ver from version where testtime='" & rcd.Fields(0) & "' and sn='" & Trim(txtSN.Text) & "'"
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
      sql = "select * from HUAWEI where SN='" & Mid(txtSN.Text, 3, 8) & "'"
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
        txtOS.BackColor = &H80000005
        txtOS.Text = rec.Fields(5)
        txtGW.Text = rec.Fields(6)
        If UCase(Trim(rec.Fields(17))) = "YES" Then
        chkchinarosh = 1
        Else
        chkchinarosh = 0
        End If
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
        If Trim(rec.Fields(10)) = "/" Then
        txtMS.Text = "N/A"
        Else
        txtMS.Text = rec.Fields(10)
        End If
        txtNAL.Text = rec.Fields(11)
'        dtpValidFrom.Value = rec.Fields(12)
'        dtpValidTo.Value = rec.Fields(13)
'        txtChangNAL.Text = rec.Fields(14)
        'txtHV.Text = rec.Fields(15)          //modified by Jimmy Sun 06.11.2010
        If rec.Fields(16) = "/" Then
        txtRemark.Text = "N/A"
        Else
        txtRemark.Text = rec.Fields(16)
        End If
        sql1 = "select top 1 * from revset where model='" & Mid(txtSN.Text, 3, 8) & "' and firstall<='" & txtSN.Text & "' and endall>='" & txtSN.Text & "' order by ver desc"
        If rec.State = 1 Then rec.Close
        rec.Open sql1, conn, adOpenKeyset, adLockOptimistic
        If Not rec.EOF Then txtHV.Text = rec.Fields(3)
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
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HUAWEI-生产.lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "HUAWEI-生产.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub
