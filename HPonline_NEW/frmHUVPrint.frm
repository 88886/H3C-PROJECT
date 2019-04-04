VERSION 5.00
Begin VB.Form frmHUVPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HUV 打印程序"
   ClientHeight    =   12555
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   14220
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
   ScaleHeight     =   12555
   ScaleWidth      =   14220
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox lblNALday 
      Height          =   450
      Left            =   600
      TabIndex        =   44
      Top             =   11880
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox lblMSday 
      Height          =   450
      Left            =   240
      TabIndex        =   43
      Top             =   11880
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   4080
      TabIndex        =   34
      Top             =   11760
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   9000
      TabIndex        =   33
      Top             =   11745
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   6480
      TabIndex        =   32
      Top             =   11745
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   13935
      Begin VB.TextBox txtENCA 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   9600
         TabIndex        =   60
         Top             =   5280
         Width           =   4095
      End
      Begin VB.TextBox txtENCN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2040
         TabIndex        =   58
         Top             =   5280
         Width           =   4095
      End
      Begin VB.TextBox txtCNCA 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   9600
         TabIndex        =   56
         Top             =   4680
         Width           =   4095
      End
      Begin VB.TextBox txtCNCN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2040
         TabIndex        =   53
         Top             =   4680
         Width           =   4095
      End
      Begin VB.CheckBox chkNonFCC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7320
         TabIndex        =   52
         Top             =   3120
         Width           =   615
      End
      Begin VB.CheckBox chkFCC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6480
         TabIndex        =   51
         Top             =   3120
         Width           =   855
      End
      Begin VB.CheckBox chkNonCN3C 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   48
         Top             =   3120
         Width           =   975
      End
      Begin VB.CheckBox chkCN3C 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   47
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtSalesArea 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   45
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11640
         TabIndex        =   42
         Top             =   3120
         Width           =   615
      End
      Begin VB.CheckBox chkRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10800
         TabIndex        =   41
         Top             =   3120
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
         Left            =   11640
         TabIndex        =   39
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox chkTurkey 
         BackColor       =   &H80000005&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10800
         TabIndex        =   38
         Top             =   2160
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
         Top             =   2640
         Width           =   615
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11640
         TabIndex        =   10
         Top             =   2640
         Width           =   615
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
         Top             =   4080
         Width           =   3135
      End
      Begin VB.TextBox txtNAL 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   14
         Top             =   3600
         Width           =   3135
      End
      Begin VB.TextBox txtMS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   13
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox txtHV 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   15
         Top             =   4080
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
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "英文地址名:"
         Height          =   375
         Left            =   7920
         TabIndex        =   59
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "英文公司名:"
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "中文地址名:"
         Height          =   375
         Left            =   7920
         TabIndex        =   55
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "中文公司名:"
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FCC:"
         Height          =   375
         Left            =   5640
         TabIndex        =   50
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息CCC:"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "标签类型:"
         Height          =   375
         Left            =   8520
         TabIndex        =   46
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblTurkeyRohs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "土耳其RoHs:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8400
         TabIndex        =   37
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息RoHS:"
         Height          =   495
         Left            =   8400
         TabIndex        =   36
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
         Height          =   375
         Left            =   8400
         TabIndex        =   31
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网许可号:"
         Height          =   375
         Left            =   8400
         TabIndex        =   30
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "执行标准:"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   3600
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
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label lblWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息WEEE:"
         Height          =   375
         Left            =   8400
         TabIndex        =   26
         Top             =   2640
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
   txtCNCN.Text = ""
   txtCNCA.Text = ""
   txtENCN.Text = ""
   txtENCA.Text = ""
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
   
   If Me.txtSalesArea.Text = "" Then
      MsgBox "没有销售区域,不能打印!", vbInformation + vbOKOnly, "销售区域不能为空"
      txtSN.SetFocus
      Exit Sub
   End If
   If Me.chkCN3C.Value + Me.chkNonCN3C.Value = 0 Or Me.chkCN3C.Value + Me.chkNonCN3C.Value = 2 Then
      MsgBox "CCC认证没有设定,不能打印!", vbInformation + vbOKOnly, "CCC认证不能为空"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If Me.chkFCC.Value + Me.chkNonFCC.Value = 0 Or Me.chkFCC.Value + Me.chkNonFCC.Value = 2 Then
      MsgBox "FCC认证没有设定或者设定错误,不能打印!", vbInformation + vbOKOnly, "FCC认证不正确"
      txtSN.SetFocus
      Exit Sub
   End If
   
   
   '===============add by ben start===============
   'Marked by carson 20150928 for guodong.ma require
   'If Trim(txtOS.Text) = "" Then
   '   MsgBox "外尺寸未带出，不能打印!", vbInformation + vbOKOnly, "未带出外尺寸"
   '   txtSN.SetFocus
   '   Exit Sub
   'End If
   
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
   
'   If chkVer.Value = 0 Then
'      myObjs("BSver").Top = 100000
'      myVars.Item("SVer").Value = "N/A"
'   Else
'       If txtVer.Text = "" Or txtVer.Text = "/" Or txtVer.Text = "N/A" Then
'         myObjs("BSver").Top = 100000
'         myVars.Item("SVer").Value = "N/A"
'       Else
'         myObjs("TSver").Top = 100000
'         myVars.Item("SVer").Value = UCase(Trim(Replace(txtVer.Text, vbCrLf, "")))
'       End If
'   End If


    If Me.txtSalesArea.Text = "讯美" Then
   
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
   
   
   myVars.Item("CPN").Value = UCase(txtCPN.Text)
   
   'myVars.Item("EPN").Value = txtEPN.Text
   'myVars.Item("Des").Value = txtDes.Text
   'add by carson 20150730 for guodong.ma require
   If InStr(txtDes.Text, "@") > 0 Then
        myVars.Item("Des").Value = Left(txtDes.Text, InStr(txtDes.Text, "@") - 1)
        'If Mid(UCase(txtSN.Text), 3, 8) <> "0235C169" And Mid(UCase(txtSN.Text), 3, 8) <> "0235C168" Then
        'If Mid(UCase(txtSN.Text), 3, 8) <> "0235C169" Then
        '    myObjs("PZH").Value = Right(txtDes.Text, Len(txtDes.Text) - InStr(txtDes.Text, "@"))
        'End If
         'add by carson 20150928 for guodong.ma require
        If txtSalesArea.Text = "讯美" Then
        Else
            myObjs("PZH").Value = Right(txtDes.Text, Len(txtDes.Text) - InStr(txtDes.Text, "@"))
        End If

   Else
        myVars.Item("Des").Value = txtDes.Text
        'If Mid(UCase(txtSN.Text), 3, 8) <> "0235C169" And Mid(UCase(txtSN.Text), 3, 8) <> "0235C168" Then
        'If Mid(UCase(txtSN.Text), 3, 8) <> "0235C169" Then
        '    myObjs("Text10(3)").Top = 100000
        '    myVars.Item("PZH").Value = ""
        'End If
        'add by carson 20150928 for guodong.ma require
        If txtSalesArea.Text = "讯美" Then
        Else
            myObjs("Text10(3)").Top = 100000
            myVars.Item("PZH").Value = ""
        End If
        'myObjs("PZH").Top = 100000
   End If
   'If chkOS.Value = 0 Or txtOS.Text = "/" Then
   '   myObjs("OD").Top = 10000
   '   myVars.Item("OD").Value = ""
   'Else
   '   myVars.Item("OD").Value = txtOS.Text
   'End If
   'If Mid(UCase(txtSN.Text), 3, 8) <> "0235C169" And Mid(UCase(txtSN.Text), 3, 8) <> "0235C168" Then
   '     myVars.Item("OD").Value = txtOS.Text
   'End If
   'myVars.Item("OD").Value = txtOS.Text
   'add by carson 20150928 for guodong.ma require
   If txtSalesArea.Text = "讯美" Then
        myVars.Item("OD").Value = txtOS.Text
   Else
        myVars.Item("OD").Value = ""
   End If
   
   myVars.Item("GW").Value = txtGW.Text
   
   If chkNonCE.Value = 1 Then
      myObjs("CE").Top = 100000
   End If
   If chkNonWEEE.Value = 1 Then
      myObjs("WEEE").Top = 100000
   End If
   If chkNonChinaRoHS.Value = 1 Then
      myObjs("China RoHS").Top = 100000
   End If
   If chkNonRoHS.Value = 1 Then
      myObjs("HUV RoHS").Top = 100000
   End If
   
   If Me.chkNonFCC.Value = 1 Then
      myObjs("FCC").Top = 100000
   End If
   
   'add by allen.yan 2014/06/07 for the requirement from Shun.Huang
   'If Mid(UCase(txtSN.Text), 3, 8) <> "0235C169" And Mid(UCase(txtSN.Text), 3, 8) <> "0235C168" Then
   If Mid(UCase(txtSN.Text), 3, 8) <> "0235C169" Then
        If Me.txtCNCN.Text = "无" Then
             myVars.Item("Company-CN").Value = ""
        Else
             myVars.Item("Company-CN").Value = Trim(Me.txtCNCN.Text)
        End If
        
        
        If Me.txtCNCA.Text = "无" Then
             myVars.Item("Adress-CN").Value = ""
        Else
             myVars.Item("Adress-CN").Value = Trim(Me.txtCNCA.Text)
        End If
        
        If Me.txtSalesArea.Text = "中英文" Then
            myVars.Item("Company-EN").Value = ""
            myVars.Item("Adress-EN").Value = ""
            myVars.Item("EPN").Value = ""
        Else
            If Me.txtENCN.Text = "无" Then
                 myVars.Item("Company-EN").Value = ""
            Else
                 myVars.Item("Company-EN").Value = Trim(Me.txtENCN.Text)
            End If
            
            
            If Me.txtENCA.Text = "无" Then
                 myVars.Item("Adress-EN").Value = ""
            Else
                 myVars.Item("Adress-EN").Value = Trim(Me.txtENCA.Text)
            End If
            
            myVars.Item("EPN").Value = txtEPN.Text
        End If
   Else '0235C169
        If Me.txtCNCN.Text = "无" Then
             myVars.Item("Company-CN").Value = ""
        Else
             myVars.Item("Company-CN").Value = Trim(Me.txtCNCN.Text)
        End If
        
        If Me.txtCNCA.Text = "无" Then
             myVars.Item("Adress-CN").Value = ""
        Else
             myVars.Item("Adress-CN").Value = Trim(Me.txtCNCA.Text)
        End If
        myVars.Item("EPN").Value = ""
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
        'If Mid(UCase(txtSN.Text), 3, 8) <> "0235C169" And Mid(UCase(txtSN.Text), 3, 8) <> "0235C168" Then
        If Mid(UCase(txtSN.Text), 3, 8) <> "0235C169" Then
            'myObjs("THver").Top = 10000
        End If
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
'   If conn1.State = 0 Then
'      conn1.ConnectionString = "Provider=SQLOLEDB;User ID=datasweep;PWD=datasweep;Initial Catalog=dsActive;Data Source=DS-DB"
'      conn1.Open
'   End If
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

'抓取到系统版本,从第一位到 分隔符-这里

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
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
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
                
        checkhp.Open sql, conn1, adOpenKeyset, adLockReadOnly
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
             str = " select top 1 t.part_number,t.part_revision,t.creation_time,t.order_number,t.serial_number from (" & _
            "select a.part_number,a.part_revision,a.creation_time,a.serial_number,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
            "select top 1 a.part_number,a.part_revision,a.creation_time,a.serial_number,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
            "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
            com.CommandText = str
            rs13.Open Source:=com
            'rs13.Open str
            If rs13.EOF = True Then
               
                    MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
                    cmdCancel_Click
                    rs13.Close
                    con13.Close
                    Exit Sub
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
            If IsValidECOVersion("HUV" & Mid(UCase(txtSN.Text), 3, 8), Me.txtHV.Text) = False Then
                cmdCancel_Click
                Exit Sub
            End If
     
      '============add by ben start=============
        Dim con4 As ADODB.Connection
        Dim rs4 As ADODB.Recordset
        Dim con5 As ADODB.Connection
        Dim rs5 As ADODB.Recordset
        Dim flagTaskOrder As Boolean, flagHasBOM As Boolean, stringBOM As String
        
       
        Set con5 = New ADODB.Connection
        Set rs5 = New ADODB.Recordset
        con5.ConnectionString = "Provider=SQLOLEDB;User ID=sa;PWD=Flash123;Initial Catalog=afg_active_90;Data Source=10.11.1.130"
        con5.ConnectionTimeout = 50
        con5.Open
        Dim stringSQL As String
        
        Set rs5.ActiveConnection = con5
        rs5.CursorType = adOpenForwardOnly
          stringSQL = "select 1 from UNIT a left join UDA_Unit b on a.unit_key = b.object_key "
        stringSQL = stringSQL + "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK' and a.serial_number = '" + UCase(Trim(Me.txtSN.Text)) + "'"
        If rs5.State = 1 Then rs5.Close
        rs5.Open stringSQL
        If rs5.EOF = True Then
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
            If Not rs5.EOF Then
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
'        rs4.Close
      '============add by ben end  =============
      
      
      '===============================
      Dim rcDavid As New ADODB.Recordset
      sql = "select PrintSV from tblHUV where  SN='" & Mid(txtSN.Text, 3, 8) & "'  and HV = '" & txtHV.Text & "'"
      rcDavid.Open sql, conn, adOpenKeyset, adLockReadOnly
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
                rs3.Open sql, con, adOpenKeyset, adLockReadOnly
                
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
                        
                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (rtrim(remark) <> '' and remark is not null AND testtime >= dateadd(month,-24,getdate())) ORDER BY testtime DESC "
'                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (rtrim(remark) <> '' and remark is not null AND testtime >= dateadd(month,-1,getdate()) ORDER BY testtime DESC "
                        If con2.State = 1 Then
                            con2.Close
                        End If
                        con2.Open
                        Set rs2 = New ADODB.Recordset
                        rs2.ActiveConnection = con2
                        rs2.Open sql, con2, adOpenKeyset, adLockReadOnly
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
'add by allen.yan for requirement by DongXu 2014/10/08
'if the remark in the table test_equ which is uploaded by test software is the same with nowver filed of setting table
'先精确匹配，如果失败采用模糊匹配
                            If stmp = Trim(stmp2) Then
                                txtVer.Text = stmp2
                            ElseIf stmp = Trim(stmp3) Then
                                If DateDiff("d", Now, CDate(endDate)) < 0 Then
                                    MsgBox "系统上传软件版本资料时错误(超过有效期)!"
                                    txtSN.Text = ""
                                    txtSN.SetFocus
                                    rs2.Close
                                    rs3.Close
                                    rcDavid.Close
                                    Exit Sub
                                Else
                                    txtVer.Text = stmp3
                                End If
                                
                            ElseIf InStr(stmp, nowver) > 0 Then
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
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
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
        rcd.Open sql, conn, adOpenKeyset, adLockReadOnly
        
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
          rs8.Open sql, conn, adOpenKeyset, adLockReadOnly
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
      
      
      sql = "select ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV,CCC,SalesLocation,Remark,FCC,CNCN,CNCA,ENCN,ENCA from tblHUV where SN='" & Mid(txtSN.Text, 3, 8) & "' and HV='" & txtHV.Text & "'"
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      If rec.EOF = True Then
         MsgBox "此产品编码未进行设置!"
         txtVer.Text = ""
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
        txtCPN.Text = rec.Fields("CPN")
        txtEPN.Text = rec.Fields("EPN")
        txtDes.Text = rec.Fields("Des")

        
        Dim psv As String
        psv = rec.Fields("PrintSV")
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
            txtOS.Text = rec.Fields("OS")
        Else
            If flagHasBOM = True Then
                If txtOS.Text = "" Then
                    txtOS.Text = rec.Fields("OS")
                Else
                    If Trim(txtOS.Text) <> Trim(rec.Fields("OS")) Then
                        MsgBox "后台尺寸数据维护不一致,请找ME确认!"
                        txtSN.Text = ""
                        txtSN.SetFocus
                        rec.Close
                        Exit Sub
                    End If
                End If
            Else
                If reprint = True Then
                    txtOS.Text = rec.Fields("OS")
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
        txtGW.Text = rec.Fields("GW")
        If UCase(Trim(rec.Fields("CE"))) = "CE" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
        ElseIf rec.Fields("CE") = "/" Or rec.Fields("CE") = "N/A" Then
           chkCE.Value = 0
           chkNonCE.Value = 1
        End If
        If UCase(Trim(rec.Fields("WEEE"))) = "WEEE" Then
           chkWEEE.Value = 1
           chkNonWEEE.Value = 0
        ElseIf rec.Fields("WEEE") = "/" Or rec.Fields("WEEE") = "N/A" Then
           chkWEEE.Value = 0
           chkNonWEEE.Value = 1
        End If
        If UCase(Trim(rec.Fields("ChinaRoHS"))) = "CHINA ROHS" Then
           chkChinaRoHS.Value = 1
           chkNonChinaRoHS.Value = 0
        ElseIf rec.Fields("ChinaRoHS") = "/" Or rec.Fields("ChinaRoHS") = "N/A" Then
           chkChinaRoHS.Value = 0
           chkNonChinaRoHS.Value = 1
        End If
        If UCase(Trim(rec.Fields("RoHS"))) = "HUV ROHS" Then
            chkRoHS.Value = 1
            chkNonRoHS.Value = 0
        ElseIf rec.Fields("RoHS") = "/" Or rec.Fields("RoHS") = "N/A" Then
            chkRoHS.Value = 0
            chkNonRoHS.Value = 1
        End If
        If UCase(Trim(rec.Fields("TurkeyRohs"))) = "TURKEY ROHS" Then
            chkTurkey.Value = 1
            chkNonTurkey.Value = 0
        ElseIf rec.Fields("TurkeyRohs") = "/" Or rec.Fields("TurkeyRohs") = "N/A" Then
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
        txtMS.Text = rec.Fields("MS")
        Me.lblMSday.Text = rec.Fields("MSValidFrom")
        txtNAL.Text = rec.Fields("NAL")
        Me.lblNALday.Text = rec.Fields("ValidFrom")
'        dtpValidFrom.Value = rec.Fields(14)
'        dtpValidTo.Value = rec.Fields(15)
'        txtChangNAL.Text = rec.Fields(16)
'        txtHV.Text = rec.Fields(17)
        If IsNull(rec.Fields("CCC")) Then
            MsgBox "没有设定CCC认证,不能打印!", vbInformation + vbOKOnly, "CCC认证不能为空"
            txtSN.SetFocus
            Exit Sub
        ElseIf rec.Fields("CCC") = "Yes" Then
            Me.chkCN3C.Value = 1
            Me.chkNonCN3C.Value = 0
        ElseIf rec.Fields("CCC") = "No" Then
            Me.chkNonCN3C.Value = 1
            Me.chkCN3C.Value = 0
        End If
        If IsNull(rec.Fields("SalesLocation")) Then
            MsgBox "没有销售区域,不能打印!", vbInformation + vbOKOnly, "销售区域不能为空"
            txtSN.SetFocus
            Exit Sub
        Else
            Me.txtSalesArea.Text = rec.Fields("SalesLocation")
        End If
        
        txtRemark.Text = rec.Fields("Remark")
        
        If IsNull(rec.Fields("FCC")) Then
            MsgBox "没有设定FCC,不能打印!", vbInformation + vbOKOnly, "FCC值不能为空"
            txtSN.SetFocus
            Exit Sub
        Else
            If UCase(rec.Fields("FCC")) = "YES" Then
                Me.chkFCC.Value = 1
                Me.chkNonFCC.Value = 0
            ElseIf UCase(rec.Fields("FCC")) = "NO" Then
                Me.chkFCC.Value = 0
                Me.chkNonFCC.Value = 1
            Else
                MsgBox "FCC设定值不对,不能打印!", vbInformation + vbOKOnly, "FCC值不能设定为非Yes和No以外的值,请重新设定"
                txtSN.SetFocus
                Exit Sub
            End If
            
        End If
        
        'add by allen.yan for shun.huang requirement 2014/06/09
        If IsNull(rec.Fields("CNCN")) Then
            MsgBox "没有设定中文公司名,不能打印!", vbInformation + vbOKOnly, "Setting中文公司名值不能为空"
            txtSN.SetFocus
            Exit Sub
        Else
            Me.txtCNCN.Text = rec.Fields("CNCN")
        End If
        
        If IsNull(rec.Fields("CNCA")) Then
            MsgBox "没有设定中文地址名,不能打印!", vbInformation + vbOKOnly, "Setting中文地址名值不能为空"
            txtSN.SetFocus
            Exit Sub
        Else
            Me.txtCNCA.Text = rec.Fields("CNCA")
        End If
        
        If IsNull(rec.Fields("ENCN")) Then
            MsgBox "没有设定英文公司名,不能打印!", vbInformation + vbOKOnly, "Setting英文公司名值不能为空"
            txtSN.SetFocus
            Exit Sub
        Else
            Me.txtENCN.Text = rec.Fields("ENCN")
        End If
        
        If IsNull(rec.Fields("ENCA")) Then
            MsgBox "没有设定英文地址名,不能打印!", vbInformation + vbOKOnly, "Setting英文地址名值不能为空"
            txtSN.SetFocus
            Exit Sub
        Else
            Me.txtENCA.Text = rec.Fields("ENCA")
        End If
        
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
   'If Mid(UCase(txtSN.Text), 3, 8) = "0235C168" Then
   '     Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\宇视发货标签NEW\宇视发货标签20140606\" & "专用机种模板2.lab")
   If Mid(UCase(txtSN.Text), 3, 8) = "0235C169" Then
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\宇视发货标签NEW\宇视发货标签20140606\" & "专用机种模板3.lab")
   Else
   
        If Me.txtSalesArea.Text = "中英文" Then
             Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\宇视发货标签NEW\宇视发货标签20140606\" & "Uniview-中英文新.lab")
        ElseIf Me.txtSalesArea.Text = "纯英文" Then
             Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\宇视发货标签NEW\宇视发货标签20140606\" & "Uniview-纯英文.lab")
        ElseIf Me.txtSalesArea.Text = "讯美" Then
             Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\宇视发货标签NEW\宇视发货标签20140606\" & "Uniview-讯美.lab")
        Else
           MsgBox "没有标签类型请ME更新Setting后台数据,不能打印!", vbInformation + vbOKOnly, "标签类型不能为空"
           Exit Sub
           Unload Me
        End If
   End If
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub OpenLppx_hp()
   Me.MousePointer = vbHourglass
   '\\sz-fs01\Public\Manufacture\标签模板\宇视发货标签NEW
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP发货标签.lab")
   
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

