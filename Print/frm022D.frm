VERSION 5.00
Begin VB.Form frm022D 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HUV 单板阶二维码标签"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10920
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
   ScaleHeight     =   6435
   ScaleWidth      =   10920
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdGoon 
      Caption         =   "继续"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9120
      TabIndex        =   21
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7680
      TabIndex        =   20
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   5280
      TabIndex        =   14
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   2760
      TabIndex        =   13
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   10455
      Begin VB.CheckBox chkY3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y3"
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtWO 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   1680
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   9120
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtXH 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   17
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y1"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkY2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y2"
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtQty1 
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   3
         Text            =   "1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtVer 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   450
         Left            =   3840
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "工单号:"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "打印数量:"
         Height          =   375
         Left            =   7800
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品型号:"
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "环保属性:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "起始条码:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "一式几份:"
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本:"
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5940
      Left            =   0
      Picture         =   "frm022D.frx":0000
      ScaleHeight     =   5910
      ScaleWidth      =   10875
      TabIndex        =   6
      Top             =   0
      Width           =   10905
   End
End
Attribute VB_Name = "frm022D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim pb As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myFormat As LabelManager2.Format
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Dim bRun As Boolean

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
   'Else
   '   chkChinaRoHS.Value = 1
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
   'Else
   '   chkNonChinaRoHS.Value = 1
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
   txtXH.Text = ""
   chkY.Value = 0
   chkY2.Value = 0
   chkY.Enabled = False
   chkY2.Enabled = False
   txtWO.Text = ""
   txtQty.Text = ""
   txtSN.SetFocus
End Sub

Private Sub cmdGoon_Click()
    bRun = True
    cmdPrint.Enabled = False
    cmdCancel.Enabled = True
    cmdReturn.Enabled = True
    cmdStop.Enabled = True
    cmdGoon.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    If chkY.Value + Me.chkY2.Value + chkY3.Value = 0 Then
        MsgBox "环保属性未输入,不能打印!", vbInformation + vbOKOnly, "未输入环保属性"
        txtSN.SetFocus
        Exit Sub
    End If
    
    sql = "select active from tblECO_Ver where PartNumber='" & Trim(txtCPN.Text) & "' and Version='" & Trim(txtVer.Text) & "'"
    If rec.State = 1 Then
      rec.Close
    End If
   
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   If rec.EOF = False Then
        If rec.Fields(0).Value = "False" Then
            MsgBox "此版本已经被禁用,不能打印!", vbInformation + vbOKOnly, "版本已经被禁用"
            txtSN.SetFocus
            Exit Sub
        End If
   End If
   rec.Close


  If txtSN.Text = "" Then
      MsgBox "产品条码未输入,不能打印!", vbInformation + vbOKOnly, "未输入产品条码"
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
   
      If txtQty1.Text = "" Then
      MsgBox "一式几份未输入，不能打印！", vbInformation + vbOKOnly, "未输入数量"
      txtQty1.SetFocus
      Exit Sub
   End If
   
   If CInt(txtQty1.Text) = 0 Then
      MsgBox "请输入正确的数量！", vbInformation + vbOKOnly, "数量不对"
      txtQty1.SetFocus
      Exit Sub
   End If
   
   
   If txtVer.Text = "" Then
      MsgBox "版本未输入,不能打印!", vbInformation + vbOKOnly, "未输入版本"
      txtWO.SetFocus
      Exit Sub
   End If
   
   If txtXH.Text = "" Then
      MsgBox "型号未输入,不能打印!", vbInformation + vbOKOnly, "未输入型号"
      txtXH.SetFocus
      Exit Sub
   End If
   
   cmdPrint.Caption = "执行中..."
   cmdPrint.Enabled = False
   cmdStop.Enabled = True
    
   Dim i, j, qty, qty1 As Integer
   Dim leftstr, rightstr, str As String, str1 As String, str2 As String, str3 As String, str4 As String, endStr As String
   
   endStr = "XXXXXXXXXXXXXXXXXXXX"
   
   
   qty = CInt(txtQty.Text)
   qty1 = CInt(txtQty1.Text)
   leftstr = UCase(Left(txtSN.Text, Len(txtSN.Text) - 6))
   If chkY2.Value = 1 Then
        pb = "Y2"
   ElseIf chkY.Value = 1 Then
        pb = "Y1"
   ElseIf chkY3.Value = 1 Then
        pb = "Y3"
   End If
    
    '' To double check the value of pb
'    If (MsgBox("环保属性为<" & pb & ">,是否继续打印", vbYesNo, "确认信息") <> vbYes) Then
'        cmdPrint.Caption = "打印(Print) &p"
'        cmdPrint.Enabled = True
'        Exit Sub
'    End If
    
   
   If pb = "Y2" Then
      rightstr = "9" + Right(txtSN.Text, 5)
   Else
      rightstr = "0" + Right(txtSN.Text, 5)
   End If
   
   OpenLppx
     
    bRun = True
    Dim k As Integer
    k = 0
    Dim strPreviousLength As Integer, strFinalLength As Integer
    Dim strFinal As String, strFinal1 As String, strFinal2 As String, strFinal3 As String, strFinal4 As String
    '===============add by Carson 2015-12-15 start===============
        Dim normalSNFlag, normalSNFlag1, normalSNFlag2, normalSNFlag3, normalSNFlag4 As Boolean
    '===============add by Carson 2015-12-15 end===============

   For i = 0 To qty - 1 Step 5
   
    '===============add by Carson 2015-12-15 start===============
        normalSNFlag = False
        normalSNFlag1 = False
        normalSNFlag2 = False
        normalSNFlag3 = False
        normalSNFlag4 = False
    '===============add by Carson 2015-12-15 end===============

'      str = leftstr & Right("000000" & CStr(CInt(rightstr) + i), 6)
'==================edit by ben 2011-10-14 start========================
       strPreviousLength = Len(rightstr)
       strFinal = CStr(CLng(rightstr) + i)
       strFinalLength = Len(strFinal)
       For m = strprevisouslength To strFinalLength - 1
            strFinal = "0" + strFinal
       Next
       str = leftstr & Right("000000" & strFinal, 6)
       normalSNFlag = True
       
       'str1
       If i + 1 > qty - 1 Then
            strFinal1 = endStr
            strFinalLength = Len(strFinal1)
       Else
            strFinal1 = CStr(CLng(rightstr) + i + 1)
            strFinalLength = Len(strFinal1)
            normalSNFlag1 = True
       End If
       For m = strprevisouslength To strFinalLength - 1
            strFinal1 = "0" + strFinal1
       Next
       str1 = leftstr & Right("000000" & strFinal1, 6)
       
       'str2
       If i + 2 > qty - 1 Then
            strFinal2 = endStr
            strFinalLength = Len(strFinal2)
       Else
            strFinal2 = CStr(CLng(rightstr) + i + 2)
            strFinalLength = Len(strFinal2)
            normalSNFlag2 = True
       End If
       
       For m = strprevisouslength To strFinalLength - 1
            strFinal2 = "0" + strFinal2
       Next
       str2 = leftstr & Right("000000" & strFinal2, 6)
       
       'str3
       If i + 3 > qty - 1 Then
            strFinal3 = endStr
            strFinalLength = Len(strFinal3)
       Else
            strFinal3 = CStr(CLng(rightstr) + i + 3)
            strFinalLength = Len(strFinal3)
            normalSNFlag3 = True
       End If
       
       For m = strprevisouslength To strFinalLength - 1
            strFinal3 = "0" + strFinal3
       Next
       str3 = leftstr & Right("000000" & strFinal3, 6)
       
       'str4
       If i + 4 > qty - 1 Then
            strFinal4 = endStr
            strFinalLength = Len(strFinal4)
       Else
            strFinal4 = CStr(CLng(rightstr) + i + 4)
            strFinalLength = Len(strFinal4)
            normalSNFlag4 = True
       End If
       
       For m = strprevisouslength To strFinalLength - 1
            strFinal4 = "0" + strFinal4
       Next
       str4 = leftstr & Right("000000" & strFinal4, 6)
       
       '======Add by mike 2015.3.24 for data upload to FTPC============
       If UploadH3CInfo(pb, Trim(str), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
            Or UploadH3CInfo(pb, Trim(str1), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
            Or UploadH3CInfo(pb, Trim(str2), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
            Or UploadH3CInfo(pb, Trim(str3), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
            Or UploadH3CInfo(pb, Trim(str4), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False Then
           MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
           txtSN.SetFocus
           UnloadLppx
           Exit Sub
       End If
       
       If UploadH3C_PB(pb, Trim(str), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
            Or UploadH3C_PB(pb, Trim(str1), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
            Or UploadH3C_PB(pb, Trim(str2), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
            Or UploadH3C_PB(pb, Trim(str3), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
            Or UploadH3C_PB(pb, Trim(str4), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False Then
           MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
           txtSN.SetFocus
           UnloadLppx
           Exit Sub
       End If
      
       '======Add by mike 2015.3.24 for data upload to FTPC============
       
'===============add by Carson 2015-12-15 start===============
        If reprint = False Then
            If Connect.isPrintedLabel(Trim(str), Me.Name) = True _
                Or Connect.isPrintedLabel(Trim(str1), Me.Name) = True _
                Or Connect.isPrintedLabel(Trim(str2), Me.Name) = True _
                Or Connect.isPrintedLabel(Trim(str3), Me.Name) = True _
                Or Connect.isPrintedLabel(Trim(str4), Me.Name) = True Then
                MsgBox ("此序列号已打印！")
                txtSN.SetFocus
                UnloadLppx
                cmdCancel_Click
                cmdPrint.Caption = "打印(Print) &p"
                cmdPrint.Enabled = True
                Exit Sub
            End If

        End If
 '===============add by Carson 2015-12-15 end=================
       
'==================edit by ben 2011-10-14 end==========================
    For j = 0 To qty1 - 1
 
        If bRun = True Then
            If k > 0 And k Mod 100 = 0 Then
                Savetime = timeGetTime '记下开始时的时间
                While timeGetTime < Savetime + 30000 '循环等待
                    DoEvents '转让控制权，以便让操作系统处理其它的事件。
                Wend
            End If
keepprint:
'            myVars.Item("2D").Value = str
'            myVars.Item("2D2").Value = str1
'            myVars.Item("2D3").Value = str2
            myVars.Item("SN1").Value = str
            myVars.Item("SN2").Value = str1
            myVars.Item("SN3").Value = str2
            myVars.Item("SN4").Value = str3
            myVars.Item("SN5").Value = str4
            
            
            'myVars.Item("Item").Value = "03" & UCase(Left(txtSN.Text, 6))
            If txtVer.Text = "" Or txtVer.Text = "/" Then
                'myObjs("Sver").Top = 5
                myVars.Item("Rev").Value = "N/A"
            ElseIf txtVer.Text = "00" Then
                myVars.Item("Rev").Value = ""
            Else
                'myObjs("Sver").Top = 5
                myVars.Item("Rev").Value = UCase(txtVer.Text)
            End If
            myVars.Item("Type").Value = txtXH.Text
            
 
            If chkY.Value = 1 Then
                myVars.Item("Rohs").Value = "Y1"
            ElseIf chkY2.Value = 1 Then
                myVars.Item("Rohs").Value = "Y2"
            ElseIf chkY3.Value = 1 Then
                myVars.Item("Rohs").Value = "Y3"
            End If
 
            'myApp.Visible = True
            myDoc.PrintLabel 1
            myDoc.FormFeed
            
'===============add by Carson 2015-12-15 start===============
            If normalSNFlag = True Then
                Call Connect.addPrintedLabel(str, Me.Name)
            End If
            If normalSNFlag1 = True Then
                Call Connect.addPrintedLabel(str1, Me.Name)
            End If
            If normalSNFlag2 = True Then
                Call Connect.addPrintedLabel(str2, Me.Name)
            End If
            If normalSNFlag3 = True Then
                Call Connect.addPrintedLabel(str3, Me.Name)
            End If
            If normalSNFlag4 = True Then
                Call Connect.addPrintedLabel(str4, Me.Name)
            End If

'===============add by Carson 2015-12-15 end=================
   
            k = k + 1
            
            DoEvents
            
        Else
            While (bRun = False)
                'sleep 1000
                DoEvents
            Wend
            
            GoTo keepprint
        End If
   Next
   Next
   
   UnloadLppx
    

   cmdCancel_Click
   
   cmdPrint.Caption = "打印(Print) &p"
   cmdPrint.Enabled = True
   
End Sub

Private Sub cmdStop_Click()
    bRun = False
    cmdPrint.Enabled = False
    cmdCancel.Enabled = False
    cmdReturn.Enabled = False
    cmdStop.Enabled = False
    cmdGoon.Enabled = True
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   If connFTPC.State = 0 Then
      connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
      connFTPC.Open
   End If
   Me.chkY.Enabled = False
   Me.chkY2.Enabled = False
   Me.chkY3.Enabled = False
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   bRun = False
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
     'txtVer.SetFocus
     cmdPrint_Click
  End If
End Sub



Private Sub txtRemark_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     cmdPrint_Click
  End If
End Sub

Private Sub txtQty1_Change()
If txtQty1.Text <> "" Then
    If Asc(Right(txtQty1.Text, 1)) > 57 Or Asc(Right(txtQty1.Text, 1)) < 48 Then
       MsgBox "只能输入数字！", vbInformation + vbOKOnly, "输入不正确"
       SendKeys "{backspace}"
       txtQty1.SetFocus
       Exit Sub
    End If
End If
End Sub
Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
    Me.txtSN.Text = Trim(Me.txtSN.Text)
      If Len(Trim(txtSN.Text)) <> 16 Then
         MsgBox "产品序号长度必须为16位的02阶单板!"
         txtSN.SetFocus
         Exit Sub
      End If
      
      If Left(Trim(txtSN.Text), 2) <> "02" Then
         MsgBox "产品序号起始必须02!"
         txtSN.SetFocus
         Exit Sub
      End If
      
      If Left(txtSN.Text, 2) = "21" Then
           sql = "select ID,SN,TYPE,CASE PB WHEN 1 THEN 'Yes' when 0 then 'No' else 'Non' end from SingleUnit where SN= '" & Mid(txtSN.Text, 3, 8) & "'"
      Else
           sql = "select ID,SN,TYPE,CASE PB WHEN 1 THEN 'Yes' when 0 then 'No' else 'Non' end from SingleUnit where SN='03" & Mid(txtSN.Text, 1, 6) & "'"
      End If
             
              rec.Open sql, conn, adOpenKeyset, adLockOptimistic
              If rec.EOF = True Then
                 MsgBox "此产品编码未进行设置!"
                 txtVer.Text = ""
                 txtSN.Text = ""
                 txtSN.SetFocus
                 rec.Close
                 Exit Sub
              Else
                txtCPN.Text = Trim(rec.Fields(1))
                txtXH.Text = Trim(rec.Fields(2))
                If (rec.Fields(3) = "Non") Then
                     MsgBox "此产品编码未进行设置有铅/无铅,不能打印!"
                     rec.Close
                     Exit Sub
                Else
'                    If (rec.Fields(3) = "No") Then
'                        Me.chkY2.Value = 1
'                        Me.chkY.Value = 0
'                        Me.chkY3.Value = 0
'                    Else
'                        Me.chkY2.Value = 0
'                        Me.chkY2.Enabled = False
'                        Me.chkY.Enabled = True
'                        Me.chkY3.Enabled = True
'                        Me.chkY.Value = 0
'                        Me.chkY3.Value = 0
'
'                    End If
                    If (rec.Fields(3) = "No") Then
                        Me.chkY2.Value = 1
                        Me.chkY2.Enabled = False
                        Me.chkY.Enabled = False
                        Me.chkY3.Enabled = False
                    Else
                        Me.chkY2.Enabled = False
                        Me.chkY2.Value = 0
                        If (Connect.GetPBStatusfromDB(txtCPN.Text) = "Y3") Then
                            Me.chkY3.Value = 1
                            Me.chkY2.Value = 0
                            Me.chkY2.Enabled = False
                            Me.chkY.Value = 0
                            Me.chkY.Enabled = False
                        ElseIf (Connect.GetPBStatusfromDB(txtCPN.Text) = "Y1") Then
                            Me.chkY.Value = 1
                            Me.chkY2.Value = 0
                            Me.chkY2.Enabled = False
                            Me.chkY3.Value = 0
                            Me.chkY3.Enabled = False
                        End If
                
                    End If
                    
                    
                    
                End If
                
               
            
              End If
              rec.Close
              txtWO.SetFocus
           Else
                txtWO.Text = ""
                txtVer.Text = ""
                txtCPN.Text = ""
                txtXH.Text = ""
           End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\打印中心\" & "0302二维码标签.Lab")
   Set myFormat = myDoc.Format
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub txtWO_KeyPress(KeyAscii As Integer)
    Dim tempWO As String
    If (KeyAscii = 13) Then
        If Len(Trim(txtCPN.Text)) <> 8 Then
            MsgBox "产品编码长度必须为8位!"
            txtSN.SetFocus
            Exit Sub
        Else
            tempWO = txtWO.Text
            sql = "select part_revision from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A ,[10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B " & _
                "WHERE A.order_key = B.order_key AND A.order_number ='" & tempWO & "' and (part_number like 'HUV%' or part_number like 'HUVC%' )"
 
            rec.Open sql, conn, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                MsgBox "SAP中此工单的编码号与此产品编码不一致或者此工单是HWF工单!"
                txtWO.Text = ""
                txtVer.Text = ""
                txtWO.SetFocus
                rec.Close
                Exit Sub
            Else
                txtVer.Text = Trim(rec.Fields(0))
                rec.Close
            End If
        End If
    Else
        txtVer.Text = ""
    End If
End Sub
