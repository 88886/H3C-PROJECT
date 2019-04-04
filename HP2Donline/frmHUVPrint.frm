VERSION 5.00
Begin VB.Form frmHUVPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HUV Label Print"
   ClientHeight    =   11280
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
   Icon            =   "frmHUVPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11280
   ScaleWidth      =   14130
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox lblNALday 
      Height          =   450
      Left            =   600
      TabIndex        =   44
      Top             =   9480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox lblMSday 
      Height          =   450
      Left            =   240
      TabIndex        =   43
      Top             =   9480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   4080
      TabIndex        =   34
      Top             =   10300
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   9000
      TabIndex        =   33
      Top             =   10300
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   6480
      TabIndex        =   32
      Top             =   10300
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   120
      TabIndex        =   19
      Top             =   5600
      Width           =   13935
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   495
         Left            =   12240
         TabIndex        =   42
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox chkRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10800
         TabIndex        =   41
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox chkVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本信息:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8400
         TabIndex        =   40
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkNonTurkey 
         BackColor       =   &H80000005&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   12240
         TabIndex        =   39
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chkTurkey 
         BackColor       =   &H80000005&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10800
         TabIndex        =   38
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "外尺寸(MM):"
         Enabled         =   0   'False
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
         Left            =   10800
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   12240
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
         Width           =   855
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   16
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtNAL 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   14
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox txtMS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   13
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtHV 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   15
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox txtGW 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   6
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtOS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         Locked          =   -1  'True
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
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblTurkeyRohs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "土耳其RoHs:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8400
         TabIndex        =   37
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息RoHS:"
         Height          =   495
         Left            =   8400
         TabIndex        =   36
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
         Height          =   375
         Left            =   8400
         TabIndex        =   31
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网许可号:"
         Height          =   375
         Left            =   8400
         TabIndex        =   30
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "执行标准:"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息ChinaRoHS:"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label lblHV 
         BackColor       =   &H00FFFFFF&
         Caption         =   "硬件版本:"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息WEEE:"
         Height          =   375
         Left            =   8400
         TabIndex        =   26
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息CE:"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblGW 
         BackColor       =   &H00FFFFFF&
         Caption         =   "毛重(kg):"
         Height          =   375
         Left            =   8400
         TabIndex        =   24
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(英文):"
         Height          =   375
         Left            =   8400
         TabIndex        =   23
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(中文):"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品描述:"
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
      Height          =   5295
      Left            =   120
      Picture         =   "frmHUVPrint.frx":13652
      ScaleHeight     =   5265
      ScaleWidth      =   13905
      TabIndex        =   18
      Top             =   300
      Width           =   13935
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HUV 标签："
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
Attribute VB_Name = "frmHUVPrint"
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
Dim serial_number As String
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

Private Sub chkChinaRoHS_Click()
   If chkChinaRoHS.Value = 1 Then
      chkNonChinaRoHS.Value = 0
   Else
      chkNonChinaRoHS.Value = 1
   End If
End Sub

Private Sub chkNonChinaRoHS_Click()
   If chkNonChinaRoHS.Value = 1 Then
      chkChinaRoHS.Value = 0
   Else
      chkChinaRoHS.Value = 1
   End If
End Sub

Private Sub chkNonRoHS_Click()
    If chkNonRoHS.Value = 1 Then
        chkRoHS.Value = 0
    Else
        chkRoHS.Value = 1
    End If
End Sub

Private Sub chkRoHS_Click()
    If chkRoHS.Value = 1 Then
        chkNonRoHS.Value = 0
    Else
        chkNonRoHS.Value = 1
    End If
End Sub

Private Sub chkTurkey_Click()
    If chkTurkey.Value = 1 Then
        chkNonTurkey.Value = 0
    Else
        chkNonTurkey.Value = 1
    End If
End Sub

Private Sub chkNonTurkey_Click()
     If chkNonTurkey.Value = 1 Then
        chkTurkey.Value = 0
     Else
        chkTurkey.Value = 1
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
   'txtMN.Text = ""
   txtOS.Text = ""
   txtGW.Text = ""

   chkNonCE.Value = 0
   chkNonWEEE.Value = 0
   chkNonChinaRoHS.Value = 0
   chkNonTurkey.Value = 0
   chkNonRoHS.Value = 0
   
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
   If Trim(txtVer.Text) = "" Then
      MsgBox "软件版本未带出,不能打印,请重新输入产品编码!", vbInformation + vbOKOnly, "未带出版本"
      txtSN.SetFocus
      Exit Sub
   End If
   If Trim(txtHV.Text) = "" Then
      MsgBox "产品没有硬件版本,不能打印!", vbInformation + vbOKOnly, "没有硬件版本"
      txtHV.SetFocus
      Exit Sub
   End If
   If Trim(txtGW.Text) = "" Then
      MsgBox "产品重量未带出,不能打印!", vbInformation + vbOKOnly, "未带出毛重"
      txtGW.SetFocus
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
   
   '===============add by ben start===============
   If Trim(txtOS.Text) = "" Then
      MsgBox "外尺寸未带出，不能打印!", vbInformation + vbOKOnly, "未带出外尺寸"
      txtSN.SetFocus
      Exit Sub
   End If
   '===============add by ben end  ===============
   
'===============add by ben 2012-02-05 start===============
    If reprint = False Then
        If Connect.isPrintedLabel(Me.txtSN.Text, Me.Name) Then
            MsgBox ("此序列号已打印！")
            txtSN.SetFocus
            Exit Sub
        End If
    End If
'===============add by ben 2012-02-05 end=================
   
   OpenLppx
   
   myVars.Item("SN").Value = UCase(txtSN.Text)
   myVars.Item("Item").Value = Mid(UCase(txtSN.Text), 3, 8)
   
   If chkVer.Value = 0 Then
      myObjs("BSver").Top = 10000
      myVars.Item("SVer").Value = "N/A"
   Else
       If txtVer.Text = "" Or txtVer.Text = "/" Or txtVer.Text = "N/A" Then
         myObjs("BSver").Top = 10000
         myVars.Item("SVer").Value = "N/A"
       Else
         myObjs("TSver").Top = 10000
         myVars.Item("SVer").Value = UCase(Trim(Replace(txtVer.Text, vbCrLf, "")))
       End If
   End If
   
   myVars.Item("CPN").Value = UCase(txtCPN.Text)
   myVars.Item("EPN").Value = txtEPN.Text
   myVars.Item("Des").Value = txtDes.Text

   'If chkOS.Value = 0 Or txtOS.Text = "/" Then
   '   myObjs("OD").Top = 10000
   '   myVars.Item("OD").Value = ""
   'Else
   '   myVars.Item("OD").Value = txtOS.Text
   'End If
   myVars.Item("OD").Value = txtOS.Text

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
   If chkNonRoHS.Value = 1 Then
      myObjs("HUV RoHS").Top = 10000
   End If
'   If Me.chkNonTurkey.Value = 1 Then
'      myObjs("Turkey RoHS").Top = 10000
'   End If
   'If optH3CRoHS.Value = True Then
   '   myObjs("3COM RoHS").Top = 10000
   'End If
   'If opt3COMRoHS.Value = True Then
   '   myObjs("3COM RoHS").Top = 2300
   '   myObjs("H3C RoHS").Top = 10000
   'End If
   'If optNonRoHS.Value = True Then
   '   myObjs("H3C RoHS").Top = 10000
   '   myObjs("3COM RoHS").Top = 10000
   'End If
   
   myVars.Item("MS").Value = txtMS.Text
   myVars.Item("NAL").Value = UCase(txtNAL.Text)
   
   If txtHV.Text = "" Or txtHV.Text = "/" Or txtHV.Text = "N/A" Then
      myObjs("BHver").Top = 10000
      myVars.Item("HVer").Value = "N/A"
   Else
      myObjs("THver").Top = 10000
      myVars.Item("HVer").Value = UCase(Trim(Replace(txtHV.Text, vbCrLf, "")))
   End If
   
   myVars.Item("Remark").Value = UCase(txtRemark.Text)
   
   sql = "Insert Into tblHPonline_PrintLog(SN,PTime,Printer) values ('" & UCase(txtSN.Text) & "',getdate(),'" & golUSERNAME & "')"
   conn.Execute sql
   
   Dim smodel As String
   smodel = Mid(Trim(txtSN.Text), 3, 8)
   
   'myApp.Visible = True
   myDoc.PrintLabel 1
   myDoc.FormFeed
   
'===============add by ben 2012-02-05 start===============
                Call Connect.addPrintedLabel(Me.txtSN.Text, Me.Name)
'===============add by ben 2012-02-05 end=================
   
   UnloadLppx
   
   
   cmdCancel_Click
   If HP_pack_label = True Then
    frmH3CPrint.Hide
    'add hp print
    
    FormHPFahuo.txtSN = hpsn
    FormHPFahuo.txtModel_hid = smodel
    
    
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

Public Function get_nextchar(strRemark As String, pipei As String) As String

    If InStr(strRemark, pipei) > 0 Then
        get_nextchar = UCase(Mid(strRemark, InStr(strRemark, pipei) + Len(pipei), 1))
    End If

End Function

Public Function get_ver(strVer As String) As String

    If InStr(strVer, "-") > 1 Then
        get_ver = Mid(strVer, 1, InStr(strVer, "-") - 1)
    Else
        get_ver = strVer
    End If
    

End Function

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
      sql = "select * from hp where charindex(h3c_bom_code,'" & Trim(txtSN.Text) & "')<>0 "
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If Not rec.EOF Then
        If rec("pack_label") = "Y" Then HP_pack_label = True
      End If
      If rec.State = 1 Then rec.Close
     
      If HP_pack_label = True Then
      
        If conn1.State = 1 Then
             conn1.Close
        End If
   
        conn1.Open
        'sql = "select top 1 * from vH3C_HP_Label where serial_number='" & Trim(txtSN.Text) & "'  order by last_modified_time DESC"
        sql = "select top 1 * from vH3C_HP_New where serial_number='" & Trim(txtSN.Text) & "'  order by last_modified_time DESC"
                
        checkhp.Open sql, conn1, adOpenKeyset, adLockOptimistic
        If checkhp.EOF = True Then
            MsgBox ("没有对应的HP条码！")
            txtSN.Text = ""
            txtSN.SetFocus
            checkhp.Close
            Exit Sub
        Else
            hpsn = checkhp.Fields(1)
            checkhp.Close
        End If
        
        If conn1.State = 1 Then
             conn1.Close
        End If
        
      End If
      
      '=========================================================================
       
            Dim con13 As ADODB.Connection
            Dim rs13 As ADODB.Recordset
            Dim com As ADODB.Command
            Dim str As String
            Set con13 = New ADODB.Connection
            Set rs13 = New ADODB.Recordset
            strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
            'con13.ConnectionTimeout = 50
            con13.Open ConnectionString:=strConn
            Set com = New ADODB.Command
            com.ActiveConnection = con13
'            str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtSN.Text) & "'"
             str = " select top 1 part_number,part_revision,creation_time,order_number,serial_number from (" & _
            "select a.part_number,a.part_revision,a.creation_time,a.serial_number,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
            "select top 1 a.part_number,a.part_revision,a.creation_time,a.serial_number,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
            "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
            com.CommandText = str
            rs13.Open Source:=com
            'rs13.Open str
            If rs13.EOF = True Then
                Dim con3 As ADODB.Connection
                Dim rs3 As ADODB.Recordset
        
                Set con3 = New ADODB.Connection
                Set rs3 = New ADODB.Recordset
                con3.ConnectionString = "Provider=SQLOLEDB;User ID=datasweep;PWD=datasweep;Initial Catalog=dsActive;Data Source=DS-DB"
                con3.ConnectionTimeout = 50
                con3.Open
                Set rs3.ActiveConnection = con3
                rs3.CursorType = adOpenDynamic
                
                str = " select top 1 part_number,part_revision,creation_time from (" & _
                " select part_number,part_revision,creation_time from dsactive.dbo.unit nolock " & _
                " where serial_number='" & Trim(txtSN.Text) & "'" & _
                " union" & _
                " select part_number,part_rev as part_revision,creation_time from dsactive.dbo.dc_task_order NOLOCK  " & _
                " where order_number=(select order_number from dsactive.dbo.taskorder_unit NOLOCK" & _
                " where serial_number='" & Trim(txtSN.Text) & "')" & _
                " ) as t " & _
                " order by t.creation_time desc"
                
                rs3.Open str
                If rs3.EOF = True Then
                    MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
                    cmdCancel_Click
                    rs13.Close
                    con13.Close
                    Exit Sub
                Else
                    txtHV.Text = rs3.Fields(1)
                End If
                rs3.Close
            Else
                txtHV.Text = rs13.Fields(1)
                serial_number = rs13.Fields(4)
            End If
            If rs13.State = 1 Then
                rs13.Close
            End If
            
            If con13.State = 1 Then
                con13.Close
            End If
     
      '============add by ben start=============
        Dim con4 As ADODB.Connection
        Dim rs4 As ADODB.Recordset
        Dim con5 As ADODB.Connection
        Dim rs5 As ADODB.Recordset
        Dim flagTaskOrder As Boolean, flagHasBOM As Boolean, stringBOM As String
        
        Set con4 = New ADODB.Connection
        Set rs4 = New ADODB.Recordset
        con4.ConnectionString = "Provider=SQLOLEDB;User ID=datasweep;PWD=datasweep;Initial Catalog=dsActive;Data Source=DS-DB"
        con4.ConnectionTimeout = 50
        con4.Open
        Set con5 = New ADODB.Connection
        Set rs5 = New ADODB.Recordset
        con5.ConnectionString = "Provider=SQLOLEDB;User ID=sa;PWD=Flash123;Initial Catalog=afg_active_90;Data Source=10.11.1.130"
        con5.ConnectionTimeout = 50
        con5.Open
        Dim stringSQL As String
        
        Set rs4.ActiveConnection = con4
        rs4.CursorType = adOpenDynamic
        Set rs5.ActiveConnection = con5
        rs5.CursorType = adOpenDynamic
        stringSQL = "select 1 from taskorder_unit with (nolock) where serial_number = '" & txtSN.Text & "' "
        If rs4.State = 1 Then rs4.Close
        rs4.Open stringSQL
        If rs4.EOF = False Then
            flagTaskOrder = True
        Else
            flagTaskOrder = False
        End If
        
        If flagTaskOrder = True Then
            txtOS.Text = ""
        Else
'            stringSQL = "select A.bom_name from BOM as A where A.bom_name in ( " & _
'            "select W.order_number + '_' + U.part_number from unit as U with(nolock) " & _
'            "left join work_order as W with(nolock) " & _
'            "on U.order_key = W.order_key where U.serial_number = '" & txtSN.Text & "' " & _
'                 ")  or A.bom_name in (" & _
'            "select '_DEL_' + W.order_number + '_' + U.part_number from unit as U with(nolock) " & _
'            "left join work_order as W with(nolock) " & _
'            "on U.order_key = W.order_key  where U.serial_number = '" & txtSN.Text & "' " & _
'            ") "
            stringSQL = "select A.bom_name from [dbo].[BOM] as A where A.bom_name in ( " & _
            "select W.order_number + '_' + U.part_number from [dbo].[UNIT] as U  " & _
            "left join [dbo].[WORK_ORDER] as W " & _
            "on U.order_key = W.order_key where U.serial_number = '" & serial_number & "' " & _
            ")  or A.bom_name in (" & _
            "select '_DEL_' + W.order_number + '_' + U.part_number from [dbo].[UNIT] as U " & _
            "left join [dbo].[WORK_ORDER] as W " & _
            "on U.order_key = W.order_key  where U.serial_number = '" & serial_number & "' " & _
            ") "
            If rs5.State = 1 Then rs5.Close
            rs5.Open stringSQL
            If rs5.EOF = False Then
                flagHasBOM = True
                stringBOM = rs5.Fields(0)
            Else
                flagHasBOM = False
            End If
'            stringSQL = "select A.bom_name from [10.11.1.130].[afg_active_90].[dbo].[BOM] as A where A.bom_name in ( " & _
'            "select W.order_number + '_' + U.part_number from unit as U with(nolock) " & _
'            "left join work_order as W with(nolock) " & _
'            "on U.order_key = W.order_key  where U.serial_number = '" & txtSN.Text & "' " & _
'            ")  or A.bom_name in (" & _
'            "select '_DEL_' + W.order_number + '_' + U.part_number from unit as U with(nolock) " & _
'            "left join work_order as W with(nolock) " & _
'            "on U.order_key = W.order_key  where U.serial_number = '" & txtSN.Text & "' " & _
'            ") "
'
'            If rs4.State = 1 Then rs4.Close
'            rs4.Open stringSQL
'            If rs4.EOF = False Then
'                flagHasBOM = True
'                stringBOM = rs4.Fields(0)
'            Else
'                flagHasBOM = False
'            End If
            
            If flagHasBOM = True Then
'                stringSQL = " select C.size_of_part from [10.11.1.130].[afg_active_90].[dbo].[BOM] as A " & _
'                "left join [10.11.1.130].[afg_active_90].[dbo].[BOM_PART_LIST] as B " & _
'                "on A.bom_key = B.bom_key " & _
'                "left join [10.11.1.130].[afg_active_90].[dbo].[BOM_PART_3003] as C " & _
'                "on B.part_number = C.part_number " & _
'                "where A.bom_name in ( " & _
'                "select W.order_number + '_' + U.part_number from unit as U with(nolock) " & _
'                "left join work_order as W with(nolock) " & _
'                "on U.order_key = W.order_key where U.serial_number = '" & txtSN.Text & "' " & _
'                ") and C.size_of_part is not null "
                stringSQL = " select C.size_of_part from [BOM] as A with (nolock) " & _
                "left join [BOM_PART_LIST] as B with (nolock) " & _
                "on A.bom_key = B.bom_key " & _
                "left join [BOM_PART_3003] as C with (nolock) " & _
                "on B.part_number = C.part_number " & _
                "where A.bom_name = '" & stringBOM & "' " & _
                "and C.size_of_part is not null "
                If rs5.State = 1 Then rs5.Close
                rs5.Open stringSQL
                If rs5.EOF = True Then
                    txtOS.Text = ""
                Else
                    txtOS.Text = rs5.Fields(0)
                End If
            Else
                If reprint = True Then
                Else
                    MsgBox "此正常品缺少工单BOM，禁止打印!"
                    txtSN.Text = ""
                    txtSN.SetFocus
                    rs4.Close
                    Exit Sub
                End If
            End If
            
        End If
        rs4.Close
      '============add by ben end  =============
      
      
      '===============================
      Dim rcDavid As New ADODB.Recordset
      sql = "select PrintSV from tblHUV where  SN='" & Mid(txtSN.Text, 3, 8) & "'  and HV = '" & txtHV.Text & "'"
      rcDavid.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcDavid.EOF Then
            MsgBox "此产品序号未收集版本!"
            txtSN.Text = ""
            txtSN.SetFocus
            rcDavid.Close
            Exit Sub
      Else
            If rcDavid.Fields(0) = "N" Then
                txtVer.Text = "N/A"
            Else
                '--------------
                Set con = New ADODB.Connection
                con.CursorLocation = adUseClient
                con.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
                con.ConnectionTimeout = 100
                
                
                sql = "select * from tblSoftVersion where model='" & Mid(txtSN.Text, 3, 8) & "'"
    
                If con.State = 1 Then
                    con.Close
                End If
   
                con.Open
    
                Set rs3 = New ADODB.Recordset
                rs3.ActiveConnection = con
                rs3.Open sql, con, adOpenKeyset, adLockOptimistic
                
                If rs3.EOF Then
                    MsgBox "此产品序号未进行发货标签软件版本维护!"
                    txtSN.Text = ""
                    txtSN.SetFocus
                    rs3.Close
                    rcDavid.Close
                    Exit Sub
                Else
                    If rs3.Fields("searchFlag") = "Y" Then
                        Set con2 = New ADODB.Connection
                        con2.CursorLocation = adUseClient
                        con2.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=dataT"
                        con2.ConnectionTimeout = 100
                        
                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (rtrim(remark) <> '' and remark is not null AND testtime >= dateadd(month,-1,getdate())) ORDER BY testtime DESC "
'                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (rtrim(remark) <> '' and remark is not null AND testtime >= dateadd(month,-1,getdate()) ORDER BY testtime DESC "
                        If con2.State = 1 Then
                            con2.Close
                        End If
                        con2.Open
                        Set rs2 = New ADODB.Recordset
                        rs2.ActiveConnection = con2
                        rs2.Open sql, con2, adOpenKeyset, adLockOptimistic
                        If rs2.EOF Then
                            MsgBox "查询软件版本资料时错误!"
                            txtSN.Text = ""
                            txtSN.SetFocus
                            rs2.Close
                            rs3.Close
                            rcDavid.Close
                            Exit Sub
                        Else
                            Dim stmp As String
                            Dim stmp2 As String
                            Dim stmp3 As String
                            Dim nowver As String
                            Dim beforver As String
                            Dim endDate As String
                            
                            stmp2 = rs3.Fields("nowVer")
                            stmp3 = rs3.Fields("beforeVer")
                            
                            nowver = Mid(stmp2, 2)
                            beforver = Mid(stmp3, 2)
                            nowver = get_ver(nowver)
                            beforver = get_ver(beforver)
                            
                            endDate = rs3.Fields("endDate")
                            
                            stmp = rs2.Fields("remark")
                            
                            If InStr(stmp, nowver) > 0 Then
                                Dim ttt As String
                                ttt = get_nextchar(stmp, nowver)
                                
                                If ttt = "L" Or ttt = "P" Then
                                    MsgBox "查询软件版本资料时错误(版本匹配错误)!"
                                    txtSN.Text = ""
                                    txtSN.SetFocus
                                    rs2.Close
                                    rs3.Close
                                    rcDavid.Close
                                    Exit Sub
                                Else
                                    txtVer.Text = stmp2
                                End If
                                
                            Else
                                If Trim(beforver) = "" Then
                                    MsgBox "查询软件版本资料时错误(版本匹配错误)!"
                                    txtSN.Text = ""
                                    txtSN.SetFocus
                                    rs2.Close
                                    rs3.Close
                                    rcDavid.Close
                                    Exit Sub
                                Else
                                    '***********
                                    
                                If InStr(stmp, beforver) > 0 Then
                                    Dim st As String
                                    st = get_nextchar(stmp, beforver)
                                    If st = "L" Or st = "P" Then
                                        MsgBox "查询软件版本资料时错误(版本匹配错误)!"
                                        txtSN.Text = ""
                                        txtSN.SetFocus
                                        rs2.Close
                                        rs3.Close
                                        rcDavid.Close
                                        Exit Sub
                                    Else
                                        If DateDiff("d", Now, CDate(endDate)) < 0 Then
                                            MsgBox "查询软件版本资料时错误(超过有效期)!"
                                            txtSN.Text = ""
                                            txtSN.SetFocus
                                            rs2.Close
                                            rs3.Close
                                            rcDavid.Close
                                            Exit Sub
                                        Else
                                            txtVer.Text = stmp3
                                        End If
                                    End If

                                Else
                                        MsgBox "查询软件版本资料时错误(版本匹配错误)!"
                                        txtSN.Text = ""
                                        txtSN.SetFocus
                                        rs2.Close
                                        rs3.Close
                                        rcDavid.Close
                                        Exit Sub
                                End If
                                    '**********
                                
                                End If
                                
                            End If
                            
                        End If
                        rs2.Close
                        con2.Close
                        
                    Else
                        If rs3.Fields("searchFlag") = "N" Then
    '=====================================================================
                            Dim stmp2_2 As String
                            Dim stmp3_2 As String
                            Dim endDate_2 As String
                            Dim nowver_2 As String
                            Dim beforver_2 As String
                            Dim stmp_2 As String
                            
                            stmp2_2 = rs3.Fields("nowVer")
                            stmp3_2 = rs3.Fields("beforeVer")
                            endDate_2 = rs3.Fields("endDate")
                            nowver_2 = Trim(stmp2_2)
                            beforver_2 = Trim(stmp3_2)
    
    sql = "select top 1 ver from version where SN='" & txtSN.Text & "' order by testtime desc"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "此产品序号未收集版本!"
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
                                      rs3.Close
                    rcDavid.Close

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
                                          rs3.Close
                    rcDavid.Close

           Exit Sub
        Else
          Dim rs8 As New ADODB.Recordset
          sql = "select ver from version where testtime='" & rcd.Fields(0) & "' and sn='" & Trim(txtSN.Text) & "'"
          rs8.Open sql, conn, adOpenKeyset, adLockOptimistic
          If rs8.EOF = False Then
'             txtVer.Text = rs8.Fields(0)
                stmp_2 = rs8.Fields(0)
                If checkVersion(stmp_2, beforver_2, nowver_2, endDate_2) Then
                    txtVer.Text = rs8.Fields(0)
                Else
                    txtSN.Text = ""
                    txtSN.SetFocus
                    rs8.Close
                    rcd.Close
                    rec.Close
                    rs3.Close
                    rcDavid.Close
                    Exit Sub
                End If
          Else
             MsgBox "此产品序号未收集版本!"
             txtSN.Text = ""
             txtSN.SetFocus
             rs8.Close
             rcd.Close
             rec.Close
             rs3.Close
             rcDavid.Close

             Exit Sub
          End If
          rs8.Close
        End If
        rcd.Close
      End If
      rec.Close
      
      '==============================================
                        End If
                    End If
                End If
                
                rs3.Close
                con.Close
                
                '--------------
            End If
      End If
      
      
      
      '===========================================================
      
      
      sql = "select ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark from tblHUV where SN='" & Mid(txtSN.Text, 3, 8) & "' and HV='" & txtHV.Text & "'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "此产品编码未进行设置!"
         txtVer.Text = ""
         txtSN.Text = ""
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
        
        chkOS.Value = 1
        'txtOS.Enabled = True
        'txtMN.Text = rec.Fields(5)
        'txtOS.BackColor = &HC0C0C0
        '============edit by ben start=============
        If flagTaskOrder = True Then
            txtOS.Text = rec.Fields(6)
        Else
            If flagHasBOM = True Then
                If txtOS.Text = "" Then
                    txtOS.Text = rec.Fields(6)
                Else
                    If Trim(txtOS.Text) <> Trim(rec.Fields(6)) Then
                        MsgBox "后台尺寸数据维护不一致,请找ME确认!"
                        txtSN.Text = ""
                        txtSN.SetFocus
                        rec.Close
                        Exit Sub
                    End If
                End If
            Else
                If reprint = True Then
                    txtOS.Text = rec.Fields(6)
                Else
                    MsgBox "此正常品缺少工单BOM，禁止打印!"
                    txtSN.Text = ""
                    txtSN.SetFocus
                    rec.Close
                    Exit Sub
                End If
            End If
        End If
'        txtOS.Text = rec.Fields(6)
        '============edit by ben end  =============
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
        If UCase(Trim(rec.Fields(11))) = "HUV ROHS" Then
            chkRoHS.Value = 1
            chkNonRoHS.Value = 0
        ElseIf rec.Fields(11) = "/" Or rec.Fields(11) = "N/A" Then
            chkRoHS.Value = 0
            chkNonRoHS.Value = 1
        End If
        If UCase(Trim(rec.Fields(12))) = "TURKEY ROHS" Then
            chkTurkey.Value = 1
            chkNonTurkey.Value = 0
        ElseIf rec.Fields(12) = "/" Or rec.Fields(12) = "N/A" Then
            chkTurkey.Value = 0
            chkNonTurkey.Value = 1
        End If
        'If UCase(Trim(rec.Fields(11))) = "H3C ROHS" Then
        '   optH3CRoHS.Value = 1
        'ElseIf UCase(Trim(rec.Fields(11))) = "3COM ROHS" Then
        '   opt3COMRoHS.Value = 1
        'ElseIf rec.Fields(11) = "/" Or rec.Fields(11) = "" Then
        '   optNonRoHS.Value = 1
        'End If
        txtMS.Text = rec.Fields(13)
        Me.lblMSday.Text = rec.Fields(14)
        txtNAL.Text = rec.Fields(15)
        Me.lblNALday.Text = rec.Fields(16)
'        dtpValidFrom.Value = rec.Fields(14)
'        dtpValidTo.Value = rec.Fields(15)
'        txtChangNAL.Text = rec.Fields(16)
'        txtHV.Text = rec.Fields(17)
        txtRemark.Text = rec.Fields(18)
        
      End If
      '==================================================
       If rec.State = 1 Then
            rec.Close
       End If
       
       cmdPrint_Click
       
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
   'If rec1.EOF = True Then
   '     Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "H3C.lab")
   'Else
   '     If Trim(rec1.Fields(0) > "30000000") And Trim(rec1.Fields(0) < "40000000") Then
   '         Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "NPI-H3C.lab")
   '     Else
   '         Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "H3C.lab")
   '     End If
   'End If
   'rec1.Close

   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "Uniview.Lab")
   
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

