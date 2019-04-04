VERSION 5.00
Begin VB.Form frmNewNECPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New NEC Print"
   ClientHeight    =   11895
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11895
   ScaleWidth      =   14160
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      TabIndex        =   5
      Top             =   5880
      Width           =   13935
      Begin VB.TextBox txtRCM 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   11880
         TabIndex        =   53
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CheckBox chkNonRCM 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   10200
         TabIndex        =   51
         Top             =   4080
         Width           =   615
      End
      Begin VB.CheckBox chkRCM 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
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
         Left            =   9480
         TabIndex        =   50
         Top             =   4080
         Width           =   855
      End
      Begin VB.CheckBox chkVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本信息:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8520
         TabIndex        =   48
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txtSZ 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   47
         Top             =   4080
         Width           =   2775
      End
      Begin VB.CheckBox chkN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N*"
         Enabled         =   0   'False
         Height          =   330
         Left            =   12480
         TabIndex        =   45
         Top             =   3600
         Width           =   855
      End
      Begin VB.CheckBox chkN4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N4"
         Enabled         =   0   'False
         Height          =   330
         Left            =   11640
         TabIndex        =   44
         Top             =   3600
         Width           =   735
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y*"
         Enabled         =   0   'False
         Height          =   330
         Left            =   10920
         TabIndex        =   43
         Top             =   3600
         Width           =   855
      End
      Begin VB.CheckBox chkY2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y2"
         Enabled         =   0   'False
         Height          =   330
         Left            =   10080
         TabIndex        =   42
         Top             =   3600
         Width           =   855
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   450
         Left            =   3120
         TabIndex        =   40
         Top             =   3600
         Width           =   975
      End
      Begin VB.CheckBox chkWEEE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2400
         TabIndex        =   39
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtVer 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10680
         TabIndex        =   24
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2400
         TabIndex        =   23
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtHV 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   22
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtDecripe1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10680
         TabIndex        =   21
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtDecripe2 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   20
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtDecripe3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10680
         TabIndex        =   19
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtGW 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   18
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtG15 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   17
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txtEAN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   450
         Left            =   10680
         TabIndex        =   16
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   450
         Left            =   10680
         TabIndex        =   15
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CheckBox chkSVPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "是"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2400
         TabIndex        =   14
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkNonSVPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "否"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3120
         TabIndex        =   13
         Top             =   2640
         Width           =   615
      End
      Begin VB.CheckBox chkCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE"
         Enabled         =   0   'False
         Height          =   330
         Left            =   9360
         TabIndex        =   12
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无CE"
         Enabled         =   0   'False
         Height          =   330
         Left            =   10560
         TabIndex        =   11
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CheckBox chkNECAddr 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NEC Addr"
         Enabled         =   0   'False
         Height          =   330
         Left            =   11880
         TabIndex        =   10
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CheckBox chkChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2400
         TabIndex        =   9
         Top             =   3120
         Width           =   615
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3120
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CheckBox chkTurkey 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   330
         Left            =   10560
         TabIndex        =   7
         Top             =   3120
         Width           =   735
      End
      Begin VB.CheckBox chkNonTurkey 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   330
         Left            =   11400
         TabIndex        =   6
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RCM ID:"
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
         Left            =   10920
         TabIndex        =   52
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RCM:"
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
         Left            =   8520
         TabIndex        =   49
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SZ:"
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "环保属性:"
         Height          =   375
         Left            =   8520
         TabIndex        =   41
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "项目描述2:"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "硬件版本:"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "项目描述1:"
         Height          =   375
         Left            =   8520
         TabIndex        =   35
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblGW 
         BackColor       =   &H00FFFFFF&
         Caption         =   "毛重(kg):"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblHV 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ChinaRoHS:"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "软件版本打印:"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
         Height          =   495
         Left            =   8520
         TabIndex        =   31
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "项目描述3:"
         Height          =   375
         Left            =   8520
         TabIndex        =   30
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "EAN码:"
         Height          =   375
         Left            =   8520
         TabIndex        =   29
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "G15版本:"
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CE:"
         Height          =   375
         Left            =   8520
         TabIndex        =   27
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Turkey RoHS:"
         Height          =   375
         Left            =   8520
         TabIndex        =   26
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WEEE:"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   3600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   11040
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   8280
      TabIndex        =   2
      Top             =   11040
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   5640
      TabIndex        =   1
      Top             =   11040
      Width           =   1815
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      Picture         =   "frmNewNECPrint.frx":0000
      ScaleHeight     =   5505
      ScaleWidth      =   13905
      TabIndex        =   4
      Top             =   240
      Width           =   13935
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NEC 标签："
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
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmNewNECPrint"
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
Dim str As String
Dim myApp2 As New LabelManager2.Application
Dim myDoc2 As LabelManager2.Document
Dim myVars2 As LabelManager2.Variables
Dim myObjs2 As LabelManager2.DocObjects
Dim myApp3 As New LabelManager2.Application
Dim myDoc3 As LabelManager2.Document
Dim myVars3 As LabelManager2.Variables
Dim myObjs3 As LabelManager2.DocObjects

Private Sub Check10_Click()

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



Private Sub cmdCancel_Click()

   txtSN.Text = ""
   txtVer.Text = ""
   txtHV.Text = ""
   txtDecripe1.Text = ""
   txtDecripe2.Text = ""
   txtDecripe3.Text = ""
   txtEAN.Text = ""
   txtG15.Text = ""
   txtGW.Text = ""
   chkNonCE.Value = 1
   chkNonChinaRoHS = 1
   chkNonTurkey = 1
   chkNonWEEE = 1
   chkNonRCM = 1
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
   
     OpenLppx

Dim PB As String
If (chkY2.Value = 1) Then
    PB = "Y2"
ElseIf (chkY.Value = 1) Then
    PB = "Y*"
ElseIf (chkN.Value = 1) Then
    PB = "N*"
ElseIf (chkN4.Value = 1) Then
    PB = "N4"
End If

    If txtG15.Text = "" Or txtG15.Text = "/" Then
      myObjs("G15(1)").Top = 10000
      Else
      myVars.Item("G15").Value = UCase(txtG15.Text)
   End If

   myVars.Item("SN").Value = UCase(txtSN.Text)

    myVars.Item("Y2").Value = PB

   myVars.Item("Part Number").Value = Mid(UCase(txtSN.Text), 3, 8)

   If txtVer.Text = "" Or txtVer.Text = "/" Then
      myObjs("Swev").Top = 10000
      Else
      myVars.Item("soft").Value = UCase(txtVer.Text)
   End If

     If txtHV.Text = "" Or txtHV.Text = "/" Then
      myObjs("Rev").Top = 10000
      Else
      myVars.Item("Rev").Value = UCase(txtHV.Text)
   End If

   If chkNonCE.Value = 1 Then
      myObjs("CE").Top = 10000
      myObjs("NEC").Top = 10000
   Else
        If chkNECAddr.Value = 0 Then
           myObjs("NEC").Top = 10000
       End If
   End If

    If chkNonRCM = 1 Then
     myObjs("RCM").Top = 10000
   End If

   If chkNonWEEE = 1 Then
   myObjs("WEEE").Top = 10000
   End If

    If chkNonChinaRoHS.Value = 1 Then
      myObjs("China RoHS(1)").Top = 10000
   End If

   If chkNonTurkey.Value = 1 Then
     myObjs("Turkey RoHS(1)").Top = 10000
     End If
    If Trim(txtGW.Text) = "" Or Trim(txtGW.Text) = "N/A" Then
        myVars.Item("GW").Value = ""
   Else
        myVars.Item("GW").Value = txtGW.Text
   End If

   If Trim(txtEAN.Text) = "" Or Trim(txtEAN.Text) = "N/A" Then
   myObjs("EAN").Top = 10000
   Else
   myVars.Item("EAN").Value = txtEAN.Text
   End If

    If Trim(txtDecripe1.Text) = "" Or Trim(txtDecripe1.Text) = "N/A" Then
   myObjs("项目描述1").Top = 10000
   Else
   myVars.Item("项目描述1").Value = txtDecripe1.Text
   End If


    If Trim(txtDecripe2.Text) = "" Or Trim(txtDecripe2.Text) = "N/A" Then
   myObjs("项目描述2").Top = 10000
   Else
   myVars.Item("项目描述2").Value = txtDecripe2.Text
   End If


    If Trim(txtDecripe3.Text) = "" Or Trim(txtDecripe3.Text) = "N/A" Then
   myObjs("项目描述3").Top = 10000
   Else
   myVars.Item("项目描述3").Value = txtDecripe3.Text
   End If

    If txtSZ.Text <> "SZ" Then
        myObjs("SZ").Top = 10000
   End If

    myApp.Visible = True
    myDoc.PrintLabel 1
    myDoc.FormFeed
    UnloadLppx
    printRe_Click
    printJapan_Click
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
      
      
    Dim lh As New Label_History
    Dim sn As String
    sn = txtSN.Text
    If (lh.Init(sn)) Then
        If lh.PB = "Y*" Then
            chkY.Value = 1
            chkY2.Value = 0
            chkN4.Value = 0
            chkN.Value = 0
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
        ElseIf lh.PB = "Y2" Then
            chkY.Value = 0
            chkY2.Value = 1
            chkN4.Value = 0
            chkN.Value = 0
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
        ElseIf lh.PB = "N*" Then
            chkY.Value = 0
            chkY2.Value = 0
            chkN4.Value = 0
            chkN.Value = 1
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
        ElseIf lh.PB = "N4" Then
            chkY.Value = 0
            chkY2.Value = 0
            chkN4.Value = 1
            chkN.Value = 0
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
        End If
    Else
        chkY.Enabled = True
        chkY2.Enabled = True
        chkN.Enabled = True
        chkN4.Enabled = True
    End If
      
'      hpsn = ""
'      If conn.State = 0 Then
'      conn.ConnectionString = Connect.getConnectionstring
'      conn.Open
'      End If
'      Dim checkhp As New ADODB.Recordset
'      'Edited by mike 2010.06.11
'
'      HP_pack_label = False
'      sql = "select * from hp where charindex(h3c_bom_code,'" & Trim(txtSN.Text) & "')<>0 "
'      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
'      If Not rec.EOF Then
'        If rec("pack_label") = "Y" Then HP_pack_label = True
'      End If
'      If rec.State = 1 Then rec.Close
'
'      If HP_pack_label = True Then
'
'        If conn11.State = 1 Then
'             conn11.Close
'        End If
'
'        strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
'            'con13.ConnectionTimeout = 50
'        conn11.Open ConnectionString:=strConn
''        conn11.Open
'        sql = " SELECT component_SN from dc_component_sn where unit_key IN (SELECT unit_key from UNIT WHERE SERIAL_NUMBER = '" & Trim(txtSN.Text) & "')" & " AND Remark = 'HP'"
'        checkhp.Open sql, conn11, adOpenKeyset, adLockReadOnly
'        If checkhp.EOF = True Then
'            MsgBox ("没有对应的HP条码！")
'            txtSN.Text = ""
'            txtSN.SetFocus
'            checkhp.Close
'            Exit Sub
'        Else
'            hpsn = checkhp.Fields(0)
'            checkhp.Close
'        End If
'
'        If conn11.State = 1 Then
'             conn11.Close
'        End If
'
'      End If
'
      '=========================================================================
            Dim con13 As ADODB.Connection
            Dim rs13 As ADODB.Recordset
            Dim com As ADODB.Command
            
            Dim part_number As String
            Dim part_revision As String
            Dim order_number As String

            Set con13 = New ADODB.Connection
            Set rs13 = New ADODB.Recordset
            strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
            con13.Open ConnectionString:=strConn
            Set com = New ADODB.Command
            com.ActiveConnection = con13
            str = " select top 1 part_number,part_revision,creation_time,order_number from (" & _
            "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
            "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
            "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK') as t where order_number is not null order by t.creation_time desc "
            com.CommandText = str
            rs13.Open Source:=com
            'rs13.Open str
            If rs13.EOF = True Then
               
                MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
                rs13.Close
                cmdCancel_Click
                Exit Sub
               
            Else
                txtHV.Text = rs13.Fields(1)
                part_number = rs13.Fields(0)
                part_revision = rs13.Fields(1)
                order_number = rs13.Fields(3)
            End If
            If rs13.State = 1 Then
                rs13.Close
            End If
            If con13.State = 1 Then
                con13.Close
            End If
            
            'add by allen yan 2014/05/20
            'the main purpose of this function is to block the ECO versions that are disabled.
            If IsValidECOVersion(part_number, Me.txtHV.Text) = False Then
                cmdCancel_Click
                Exit Sub
            End If

       
        Dim conSZ As ADODB.Connection
        Dim rsSZ As ADODB.Recordset
        Set conSZ = New ADODB.Connection
        Set rsSZ = New ADODB.Recordset
        conSZ.ConnectionString = "Provider=SQLOLEDB;User ID=sa;PWD=Flash123;Initial Catalog=afg_active_90;Data Source=10.11.1.130"
        conSZ.ConnectionTimeout = 50
        conSZ.Open
'        Dim stringSQL As String
        Set rsSZ.ActiveConnection = conSZ
        rsSZ.CursorType = adOpenDynamic

        stringSQL = " select TOP 1 'SZ' from C_NoTR5_Part where EFFE_FLAG='1' AND  Part_Number ='" & Mid(txtSN.Text, 3, 8) & "'  "

        rsSZ.Open stringSQL
        If rsSZ.EOF = True Then
            txtSZ.Text = ""
        Else
            txtSZ.Text = rsSZ.Fields(0)
        End If
        rsSZ.Close
  
      Dim rcDavid As New ADODB.Recordset
      sql = "select case when Print_SV = 1 then 'Y' else 'N' end from tblNEC where  PartNum ='" & Mid(txtSN.Text, 3, 8) & "'  and HardWare_Revision = '" & txtHV.Text & "'"
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
                        
                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (rtrim(remark) <> '' and left(equipment,3) <> 'MTP' and remark is not null AND testtime >= dateadd(month,-3,getdate())) ORDER BY testtime DESC "
'                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ_ATE where barcode='" & Trim(txtSN.Text) & "' AND (ISNULL(remark, '') <> '') ORDER BY testtime DESC "
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
                            
                            'stmp2,stmp3 is ME settings sv
                            stmp2 = rs3.Fields("nowVer")
                            stmp3 = rs3.Fields("beforeVer")
                            
                            nowver = Mid(stmp2, 2)
                            beforver = Mid(stmp3, 2)
                            nowver = get_ver(nowver)
                            beforver = get_ver(beforver)
                            
                            endDate = rs3.Fields("endDate")
                            
                            'stmp is test sv
                            stmp = rs2.Fields("remark")
'update by allen.yan for the DongXu 2014/10/9
'exactly match first, if not then try faintly match
'先采用精确匹配，若匹配不成功，使用模糊匹配
                            If stmp2 = stmp Then
                                txtVer.Text = stmp
                            ElseIf stmp = smtp3 Then
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
                            ElseIf InStr(stmp, nowver) > 0 Then
                                Dim ttt As String
                                '查询测试记录中，ME维护的软件版本后的下一位字符
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
                                    
                            Else '''''''''stmp2 = stmp
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
                            End If ''''''''stmp2 = stmp end
                            
                        End If 'rs2 end
                        
                        rs2.Close
                        con2.Close
                        
                    Else '''''If rs3.Fields("searchFlag") = "Y" de else
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
                    '               txtVer.Text = rs8.Fields(0)
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
                                End If 'rcd.EOF = True
                                rcd.Close
                            End If 'rec.EOF = True
                            rec.Close
      '==============================================
                        End If 'rs3.Fields("searchFlag") = "N"
                        
                    End If ''''''If rs3.Fields("searchFlag") = "Y" end
                    
                End If 'rs3.EOF
                
                rs3.Close
                con.Close
                
                '--------------
            End If 'rcDavid.Fields(0) = "N"
      End If 'rcDavid.EOF
      
      
      
      '===========================================================
      
     sql = "SELECT [ID],[PartNum],[HardWare_Revision],[Product_Describe1],[Product_Describe2],[Product_Describe3],[EAN],[G15_Revision],[GW] " & _
        ",case when Print_SV is null then 'N/A' when Print_SV = 0 then 'No' when Print_SV = 1 then 'Yes' end as 'Print_SV'" & _
        ",case when [CE] = 0 then 'Non CE' when CE = 1 then 'CE' when CE = 2 then 'NEC Addr' end as 'CE'" & _
        ",case when ChinaRoHS is null then 'N/A' when ChinaRoHS = 0 then 'No' when ChinaRoHS = 1 then 'Yes' end as 'ChinaRoHS'" & _
        ",case when WEEE is null then 'N/A' when WEEE = 0 then 'No' when WEEE = 1 then 'Yes' end as 'WEEE'" & _
        ",case when [TurkeyRoHS] is null then 'N/A' when [TurkeyRoHS] = 0 then 'No' when TurkeyRoHS = 1 then 'Yes' end as '[TurkeyRoHS]'" & _
        ",case when RCM is null then 'N/A' when RCM = 0 then 'No' when RCM = 1 then 'Yes' end as 'RCM'" & _
        ",[RCM_ID],[Remark]" & _
        " FROM [Print].[dbo].[tblNEC] where PartNum = '" & Mid(txtSN.Text, 3, 8) & "'and  HardWare_Revision = '" & txtHV.Text & "'"
     
      

      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      If rec.EOF = True Then
         MsgBox "此产品编码未进行设置!"
         txtVer.Text = ""
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
        Me.txtDecripe1.Text = rec.Fields(3)
        txtDecripe2.Text = rec.Fields(4)
        txtDecripe3.Text = rec.Fields(5)

        If IsNull(rec.Fields(8)) = True Then
            MsgBox "此正常品缺少毛重数据，禁止打印!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
        Else
            txtGW.Text = rec.Fields(8)
        End If
    
        Me.txtEAN.Text = rec.Fields(6)
        Me.txtG15.Text = rec.Fields(7)

        If UCase(Trim(rec.Fields(10))) = "CE" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
           chkNECAddr.Value = 0
        ElseIf rec.Fields(10) = "Non CE" Then
           chkCE.Value = 0
           chkNonCE.Value = 1
          chkNECAddr.Value = 0
        ElseIf rec.Fields(10) = "NEC Addr" Then
           chkCE.Value = 1
           chkNonCE.Value = 0
           chkNECAddr.Value = 1
        End If
        
        If UCase(Trim(rec.Fields(12))) = "YES" Then
           chkWEEE.Value = 1
           chkNonWEEE.Value = 0
        ElseIf rec.Fields(12) = "No" Or rec.Fields(12) = "N/A" Then
           chkWEEE.Value = 0
           chkNonWEEE.Value = 1
        End If
        If UCase(Trim(rec.Fields(11))) = "YES" Then
           chkChinaRoHS.Value = 1
           chkNonChinaRoHS.Value = 0
        ElseIf rec.Fields(11) = "No" Or rec.Fields(16) = "N/A" Then
           chkChinaRoHS.Value = 0
           chkNonChinaRoHS.Value = 1
        End If
   
        If UCase(Trim(rec.Fields(13))) = "YES" Then
            chkTurkey.Value = 1
            chkNonTurkey.Value = 0
        ElseIf rec.Fields(13) = "No" Or rec.Fields(18) = "N/A" Then
            chkTurkey.Value = 0
            chkNonTurkey.Value = 1
        End If
        
      
        
        If UCase(Trim(rec.Fields(14))) = "YES" Then
            Me.chkRCM.Value = 1
            Me.chkNonRCM.Value = 0
        Else
            Me.chkRCM.Value = 0
            Me.chkNonRCM.Value = 1
            Me.txtRCM.Text = ""
        End If
        Me.txtRCM.Text = rec.Fields(15)
        
       
        
        If UCase(Trim(rec.Fields(9))) = "YES" Then
            Me.chkSVPrint.Value = 1
            Me.chkNonSVPrint.Value = 0
        Else
            Me.chkSVPrint.Value = 0
            Me.chkNonSVPrint.Value = 1
        End If
        
        txtHV.Text = rec.Fields(2)
        txtRemark.Text = rec.Fields(16)
        
      End If
      '==================================================
       If rec.State = 1 Then
            rec.Close
       End If
       
       
       If chkY2.Value + chkY.Value + chkN.Value + chkN4.Value > 0 Then
           cmdPrint_Click
       Else
            MsgBox "Please select the value of PB"
       End If
       
   End If
   
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\10.11.1.25\Public\Manufacture\标签模板\" & "NEC发货模板.lab")
   
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub OpenLppx_Japan()
   Me.MousePointer = vbHourglass
   Set myDoc2 = myApp2.Documents.Open("\\10.11.1.25\Public\Manufacture\标签模板\" & "NEC日文提示.lab")
   Me.MousePointer = vbDefault
   Set myVars2 = myDoc2.Variables
   Set myObjs2 = myDoc2.DocObjects
End Sub

Private Sub UnloadLppx2()
    myApp2.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp2.Quit
    Set myApp2 = Nothing
End Sub

Private Sub printJapan_Click()
    OpenLppx_Japan
    myApp2.Visible = True
    myDoc2.PrintLabel 1
    myDoc2.FormFeed
    UnloadLppx2
    
End Sub
Private Sub OpenLppx_Reprint()
   Me.MousePointer = vbHourglass
   Set myDoc3 = myApp3.Documents.Open("\\10.11.1.25\Public\Manufacture\标签模板\" & "NEC发货模板.lab")
   
   Me.MousePointer = vbDefault
   Set myVars3 = myDoc3.Variables
   Set myObjs3 = myDoc3.DocObjects
End Sub
Private Sub UnloadLppx3()
    myApp3.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp3.Quit
    Set myApp3 = Nothing
End Sub

Private Sub printRe_Click()
    OpenLppx_Reprint
    
Dim PB1 As String
If (chkY2.Value = 1) Then
    PB1 = "Y2"
ElseIf (chkY.Value = 1) Then
    PB1 = "Y*"
ElseIf (chkN.Value = 1) Then
    PB1 = "N*"
ElseIf (chkN4.Value = 1) Then
    PB1 = "N4"
End If

    If txtG15.Text = "" Or txtG15.Text = "/" Then
      myObjs3("G15(1)").Top = 10000
      Else
      myVars3.Item("G15").Value = UCase(txtG15.Text)
   End If

   myVars3.Item("SN").Value = UCase(txtSN.Text)

    myVars3.Item("Y2").Value = PB1

   myVars3.Item("Part Number").Value = Mid(UCase(txtSN.Text), 3, 8)

   If txtVer.Text = "" Or txtVer.Text = "/" Then
      myObjs3("Swev").Top = 10000
      Else
      myVars3.Item("soft").Value = UCase(txtVer.Text)
   End If

     If txtHV.Text = "" Or txtHV.Text = "/" Then
      myObjs3("Rev").Top = 10000
      Else
      myVars3.Item("Rev").Value = UCase(txtHV.Text)
   End If

   If chkNonCE.Value = 1 Then
      myObjs3("CE").Top = 10000
      myObjs3("NEC").Top = 10000
   Else
        If chkNECAddr.Value = 0 Then
           myObjs3("NEC").Top = 10000
       End If
   End If

    If chkNonRCM = 1 Then
     myObjs3("RCM").Top = 10000
   End If

   If chkNonWEEE = 1 Then
   myObjs3("WEEE").Top = 10000
   End If

    If chkNonChinaRoHS.Value = 1 Then
      myObjs3("China RoHS(1)").Top = 10000
   End If

   If chkNonTurkey.Value = 1 Then
     myObjs3("Turkey RoHS(1)").Top = 10000
     End If
    If Trim(txtGW.Text) = "" Or Trim(txtGW.Text) = "N/A" Then
        myVars3.Item("GW").Value = ""
   Else
        myVars3.Item("GW").Value = txtGW.Text
   End If

   If Trim(txtEAN.Text) = "" Or Trim(txtEAN.Text) = "N/A" Then
   myObjs3("EAN").Top = 10000
   Else
   myVars3.Item("EAN").Value = txtEAN.Text
   End If

    If Trim(txtDecripe1.Text) = "" Or Trim(txtDecripe1.Text) = "N/A" Then
   myObjs3("项目描述1").Top = 10000
   Else
   myVars3.Item("项目描述1").Value = txtDecripe1.Text
   End If


    If Trim(txtDecripe2.Text) = "" Or Trim(txtDecripe2.Text) = "N/A" Then
   myObjs3("项目描述2").Top = 10000
   Else
   myVars3.Item("项目描述2").Value = txtDecripe2.Text
   End If


    If Trim(txtDecripe3.Text) = "" Or Trim(txtDecripe3.Text) = "N/A" Then
   myObjs3("项目描述3").Top = 10000
   Else
   myVars3.Item("项目描述3").Value = txtDecripe3.Text
   End If

    If txtSZ.Text <> "SZ" Then
        myObjs3("SZ").Top = 10000
   End If


    myApp3.Visible = True
    myDoc3.PrintLabel 1
    myDoc3.FormFeed
    UnloadLppx3
End Sub
