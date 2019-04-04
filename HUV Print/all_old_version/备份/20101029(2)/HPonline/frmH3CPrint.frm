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
      Begin VB.OptionButton optNonRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   495
         Left            =   13080
         TabIndex        =   42
         Top             =   2640
         Width           =   615
      End
      Begin VB.OptionButton opt3COMRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3COMRoHS"
         Enabled         =   0   'False
         Height          =   615
         Left            =   11760
         TabIndex        =   41
         Top             =   2520
         Width           =   1215
      End
      Begin VB.OptionButton optH3CRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "H3CRoHS"
         Enabled         =   0   'False
         Height          =   615
         Left            =   10560
         TabIndex        =   40
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtMN 
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   38
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "外尺寸(MM):"
         Enabled         =   0   'False
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
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   2160
         Width           =   855
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无CE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10560
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   12480
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtRemark 
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   16
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtNAL 
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   14
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox txtMS 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   13
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtHV 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   15
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox txtGW 
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   6
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtOS 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtDes 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   4
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtEPN 
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   3
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtCPN 
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
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   1
         Top             =   240
         Width           =   3135
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
         Caption         =   "产品编码:"
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
Dim rec1 As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim hpsn As String
Dim checkhp As New ADODB.Recordset
Public HP_pack_label As Boolean

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
   If txtVer.Text = "" Or txtVer.Text = "/" Then
      myObjs("BSver").Top = 10000
      myVars.Item("SVer").Value = "N/A"
   Else
      myObjs("TSver").Top = 10000
      myVars.Item("SVer").Value = UCase(txtVer.Text)
   End If
   myVars.Item("CPN").Value = UCase(txtCPN.Text)
   'myVars.Item("EPN").Value = UCase(txtEPN.Text)
   myVars.Item("EPN").Value = txtEPN.Text
   'myVars.Item("Des").Value = UCase(txtDes.Text)
   myVars.Item("Des").Value = txtDes.Text
   'myVars.Item("MN").Value = UCase(txtMN.Text)
    myVars.Item("MN").Value = txtMN.Text
   If chkOS.Value = 0 Or txtOS.Text = "/" Then
      myObjs("OD").Top = 10000
      myVars.Item("OD").Value = ""
   Else
      'myVars.Item("OD").Value = UCase(txtOS.Text)
      myVars.Item("OD").Value = txtOS.Text
   End If
   'myVars.Item("GW").Value = UCase(txtGW.Text)
    myVars.Item("GW").Value = txtGW.Text
   If chkNonCE.Value = 1 Then
      myObjs("CE").Top = 10000
   End If
   If chkNonWEEE.Value = 1 Then
      myObjs("WEEE").Top = 10000
   End If
   If chkNonChinaRoHS.Value = 1 Then
      myObjs("China RoHS").Top = 10000
   End If
   If optH3CRoHS.Value = True Then
      myObjs("3COM RoHS").Top = 10000
   End If
   If opt3COMRoHS.Value = True Then
      myObjs("3COM RoHS").Top = 2300
      myObjs("H3C RoHS").Top = 10000
   End If
   If optNonRoHS.Value = True Then
      myObjs("H3C RoHS").Top = 10000
      myObjs("3COM RoHS").Top = 10000
   End If
   myVars.Item("MS").Value = txtMS.Text
   sql = "select ChangNAL from H3C where SN='" & txtSN.Text & "' and ValidTo<='" & Date & "'"
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   If rec.EOF = True Then
      myVars.Item("NAL").Value = UCase(txtNAL.Text)
   Else
      myVars.Item("NAL").Value = rec.Fields(0)
   End If
   rec.Close
   If txtHV.Text = "" Or txtHV.Text = "/" Then
      myObjs("BHver").Top = 10000
      myVars.Item("HVer").Value = UCase(txtHV.Text)
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
   If HP_pack_label = True Then
    frmH3CPrint.Hide
    'add hp print
   
    FormHPFahuo.txtSN = hpsn
    FormHPFahuo.Show
    Call FormHPFahuo.cmdMPrint_Click
   End If

End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   If conn1.State = 0 Then
      conn1.ConnectionString = "Provider=SQLOLEDB;User ID=datasweep;PWD=datasweep;Initial Catalog=dsActive;Data Source=DS-DB"
      conn1.Open
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
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
      If Len(txtSN.Text) < 10 Then
         MsgBox "产品序号长度不能小于10!"
         txtSN.SetFocus
         Exit Sub
      End If
      hpsn = ""
      If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
      End If
      Dim checkhp As New ADODB.Recordset
      'Edited by mike 2010.06.11
       
      HP_pack_label = False
      sql = "select * from hp where charindex(h3c_bom_code,'" & Trim(txtSN.Text) & "')<>0"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If Not rec.EOF Then
        If rec("pack_label") = "Y" Then HP_pack_label = True
      End If
      If rec.State = 1 Then rec.Close
     
     If HP_pack_label = True Then
     
        Dim rdgettime As New ADODB.Recordset
        sql = "select max(last_modified_time) from vH3C_HP_test where serial_number='" & Trim(txtSN.Text) & "' "
        rdgettime.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rdgettime.EOF = True Then
            MsgBox ("没有对应的HP条码！")
            txtSN.Text = ""
            txtSN.SetFocus
            rdgettime.Close
            Exit Sub
        Else
            'sql = "select top 1 * from vH3C_HP_new where serial_number='" & Trim(txtSN.Text) & "' order by last_modified_time DESC"
            sql = "select * from vH3C_HP_test where serial_number='" & Trim(txtSN.Text) & "' and datediff(second,convert(varchar(100), '" & rdgettime.Fields(0) & "',120),last_modified_time)=0"
            checkhp.Open sql, conn, adOpenKeyset, adLockOptimistic
            If checkhp.EOF = True Then
                MsgBox ("没有对应的HP条码！")
                txtSN.Text = ""
                txtSN.SetFocus
                checkhp.Close
                rdgettime.Close
                Exit Sub
            Else
                hpsn = checkhp.Fields(1)
                checkhp.Close
            End If
            rdgettime.Close
        End If
        
     End If
     
     
      sql = "select top 1 ver from version where SN='" & txtSN.Text & "' order by testtime desc"
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
        If rcd.EOF = True Then
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
'        txtHV.Text = rec.Fields(17)
        txtRemark.Text = rec.Fields(18)
      sql1 = "select top 1 * from revset where model='" & Mid(txtSN.Text, 3, 8) & "' and firstall<='" & txtSN.Text & "' and endall>='" & txtSN.Text & "'order by ver desc"
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
   'sql99 = "select order_number from work_order  where order_key in (select order_key from unit where serial_number='" & txtSN.Text & "')"
   'rec1.Open sql99, conn1, adOpenKeyset, adLockOptimistic
   'If Trim(rec1.Fields(0) > "30000000") And Trim(rec1.Fields(0) < "40000000") Then
   '     Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "NPI-H3C.lab")
   'Else
   '     Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "H3C.lab")
   'End If
   'rec1.Close
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C.lab")

   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "H3C.lab")
   
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
   
End Sub

Private Sub OpenLppx_hp()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP发货标签.lab")
   
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

