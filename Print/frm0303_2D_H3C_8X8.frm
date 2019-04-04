VERSION 5.00
Begin VB.Form frm0303_2D_H3C_8X8 
   Caption         =   "0303 2D标签打印(H3C)"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   14640
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5940
      Left            =   0
      Picture         =   "frm0303_2D_H3C_8X8.frx":0000
      ScaleHeight     =   5910
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   0
      Width           =   10905
      Begin VB.TextBox tbFirst 
         Enabled         =   0   'False
         Height          =   450
         Left            =   0
         TabIndex        =   24
         Top             =   5400
         Width           =   495
      End
      Begin VB.Frame fmVar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   2900
         Width           =   10455
         Begin VB.CheckBox chkN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N*"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3240
            TabIndex        =   26
            Top             =   1320
            Width           =   735
         End
         Begin VB.CheckBox chkN4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N4"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3960
            TabIndex        =   25
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtVer 
            BackColor       =   &H80000011&
            Enabled         =   0   'False
            Height          =   450
            Left            =   3840
            TabIndex        =   15
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtSN 
            BackColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   1680
            TabIndex        =   14
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtCPN 
            BackColor       =   &H80000011&
            Enabled         =   0   'False
            Height          =   405
            Left            =   6840
            TabIndex        =   13
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtQty1 
            BackColor       =   &H80000011&
            Enabled         =   0   'False
            Height          =   405
            Left            =   6840
            TabIndex        =   12
            Text            =   "1"
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox chkY2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Y2"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            TabIndex        =   11
            Top             =   1320
            Width           =   735
         End
         Begin VB.CheckBox chkY 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Y*"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   10
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtXH 
            BackColor       =   &H80000011&
            Enabled         =   0   'False
            Height          =   405
            Left            =   6840
            TabIndex        =   9
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox txtQty 
            BackColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   9120
            TabIndex        =   8
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtWO 
            BackColor       =   &H00FFFFFF&
            Height          =   450
            Left            =   1680
            TabIndex        =   7
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblVer 
            BackColor       =   &H00FFFFFF&
            Caption         =   "版本:"
            Height          =   375
            Left            =   3000
            TabIndex        =   23
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lblDes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "一式几份:"
            Height          =   375
            Left            =   5160
            TabIndex        =   22
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblSN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "起始条码:"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblCPN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "产品编码:"
            Height          =   375
            Left            =   5160
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblChinaRoHS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "环保属性:"
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblMN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "产品型号:"
            Height          =   375
            Left            =   5160
            TabIndex        =   18
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "打印数量:"
            Height          =   375
            Left            =   7800
            TabIndex        =   17
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblWO 
            BackColor       =   &H00FFFFFF&
            Caption         =   "工单号:"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   1440
         End
      End
      Begin VB.CommandButton cmdGoon 
         Caption         =   "继续"
         Enabled         =   0   'False
         Height          =   615
         Left            =   9600
         TabIndex        =   5
         Top             =   5000
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "暂停"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7920
         TabIndex        =   4
         Top             =   5000
         Width           =   1095
      End
      Begin VB.CommandButton cmdReturn 
         Caption         =   "返回(Return)"
         Height          =   615
         Left            =   5640
         TabIndex        =   3
         Top             =   5000
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(Cancel)"
         Height          =   615
         Left            =   3240
         TabIndex        =   2
         Top             =   5000
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "打印(Print) &p"
         Height          =   615
         Left            =   840
         TabIndex        =   1
         Top             =   5000
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm0303_2D_H3C_8X8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' TBT

Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Dim bRun As Boolean
Dim first As String



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
  
   txtQty.Text = ""
    
    Me.chkY.Value = 0
    Me.chkY2.Value = 0
    Me.chkN.Value = 0
    Me.chkN4.Value = 0
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

    If Me.chkY.Value = 0 And Me.chkY2.Value = 0 And Me.chkN.Value = 0 And Me.chkN4.Value = 0 Then
        MsgBox "环保属性未输入,不能打印!", vbInformation + vbOKOnly, "未输入环保属性"
        txtSN.SetFocus
        Exit Sub
    End If
    If Me.chkY.Value + Me.chkY2.Value + Me.chkN.Value + Me.chkN4.Value > 1 Then
        MsgBox "环保属性输入多个,不能打印!", vbInformation + vbOKOnly, "输入环保属性多个"
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
   Dim leftstr, rightstr, str As String
   qty = CInt(txtQty.Text)
   qty1 = CInt(txtQty1.Text)
   leftstr = UCase(Left(txtSN.Text, 10))
   rightstr = tbFirst.Text + Right(Me.txtSN.Text, 5)
     
'' To double check the value of pb
    Dim Pb As String
    If chkY2.Value = 1 Then
         Pb = "Y2"
    ElseIf chkY.Value = 1 Then
         Pb = "Y*"
    ElseIf chkN.Value = 1 Then
         Pb = "N*"
    ElseIf chkN4.Value = 1 Then
         Pb = "N4"
    End If
'    If (MsgBox("环保属性为<" & pb & ">,是否继续打印", vbYesNo, "确认信息") <> vbYes) Then
'        cmdPrint.Caption = "打印(Print) &p"
'        cmdPrint.Enabled = True
'        Exit Sub
'    End If
     
     
    OpenLppx
     
    bRun = True
    Dim k As Integer
    k = 0
   
   
      '' build serial number from SN1 TO SN5
   Dim strSN1 As String, strSN2 As String, strSN3 As String, strSN4 As String, strSN5 As String
    '===============add by Carson 2015-12-15 start===============
    Dim normalSNFlag, normalSNFlag1, normalSNFlag2, normalSNFlag3, normalSNFlag4 As Boolean
    '===============add by Carson 2015-12-15 end===============

   For i = 0 To qty - 1 Step 6
   
    '===============add by Carson 2015-12-15 start===============
    normalSNFlag = False
    normalSNFlag1 = False
    normalSNFlag2 = False
    normalSNFlag3 = False
    normalSNFlag4 = False
    normalSNFlag5 = False
    '===============add by Carson 2015-12-15 end===============
   strSN1 = "XXXXXX"
   strSN2 = "XXXXXX"
   strSN3 = "XXXXXX"
   strSN4 = "XXXXXX"
   strSN5 = "XXXXXX"
   strSN6 = "XXXXXX"
'      str = leftstr & Right("000000" & CStr(CLng(rightstr) + i), 6)
    If (i / 6) * 6 < qty Then
        strSN1 = leftstr & "0" & Right("000000" & CStr(CLng(rightstr) + i + 0), 5)
        normalSNFlag = True
    End If
    If (i / 6) * 6 + 1 < qty Then
        strSN2 = leftstr & "0" & Right("000000" & CStr(CLng(rightstr) + i + 1), 5)
        normalSNFlag1 = True
    End If
    If (i / 6) * 6 + 2 < qty Then
        strSN3 = leftstr & "0" & Right("000000" & CStr(CLng(rightstr) + i + 2), 5)
        normalSNFlag2 = True
    End If
    If (i / 6) * 6 + 3 < qty Then
        strSN4 = leftstr & "0" & Right("000000" & CStr(CLng(rightstr) + i + 3), 5)
        normalSNFlag3 = True
    End If
    If (i / 6) * 6 + 4 < qty Then
        strSN5 = leftstr & "0" & Right("000000" & CStr(CLng(rightstr) + i + 4), 5)
        normalSNFlag4 = True
    End If
    If (i / 6) * 6 + 5 < qty Then
        strSN6 = leftstr & "0" & Right("000000" & CStr(CLng(rightstr) + i + 5), 5)
        normalSNFlag5 = True
    End If
    
     '======Add by mike 2015.3.24 for data upload to FTPC============
    If UploadH3CInfo(Pb, Trim(strSN1), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
         Or UploadH3CInfo(Pb, Trim(strSN2), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
         Or UploadH3CInfo(Pb, Trim(strSN3), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
         Or UploadH3CInfo(Pb, Trim(strSN4), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
         Or UploadH3CInfo(Pb, Trim(strSN5), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
         Or UploadH3CInfo(Pb, Trim(strSN6), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False Then
        MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    
    If UploadH3C_PB(Pb, Trim(strSN1), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
         Or UploadH3C_PB(Pb, Trim(strSN2), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
         Or UploadH3C_PB(Pb, Trim(strSN3), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
         Or UploadH3C_PB(Pb, Trim(strSN4), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
         Or UploadH3C_PB(Pb, Trim(strSN5), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
          Or UploadH3C_PB(Pb, Trim(strSN6), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False Then
        MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============

    '===============add by Carson 2015-12-15 start===============
    If reprint = False Then
        If Connect.isPrintedLabel(Trim(strSN1), Me.Name) = True _
            Or Connect.isPrintedLabel(Trim(strSN2), Me.Name) = True _
            Or Connect.isPrintedLabel(Trim(strSN3), Me.Name) = True _
            Or Connect.isPrintedLabel(Trim(strSN4), Me.Name) = True _
            Or Connect.isPrintedLabel(Trim(strSN5), Me.Name) = True _
             Or Connect.isPrintedLabel(Trim(strSN6), Me.Name) = True Then
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


    For j = 0 To qty1 - 1
 
       If bRun = True Then
            If k > 0 And k Mod 100 = 0 Then
                    Savetime = timeGetTime '记下开始时的时间
                    While timeGetTime < Savetime + 30000 '循环等待
                        DoEvents '转让控制权，以便让操作系统处理其它的事件。
                    Wend
            End If
keepprint:
'            myVars.Item("sn").Value = Trim(str)
            myVars.Item("SN1").Value = Trim(strSN1)
            myVars.Item("SN2").Value = Trim(strSN2)
            myVars.Item("SN3").Value = Trim(strSN3)
            myVars.Item("SN4").Value = Trim(strSN4)
            myVars.Item("SN5").Value = Trim(strSN5)
            myVars.Item("SN6").Value = Trim(strSN6)
            
            If txtVer.Text = "" Or txtVer.Text = "/" Then
                myVars.Item("Rev").Value = "N/A"
            ElseIf txtVer.Text = "00" Then
                myVars.Item("Rev").Value = ""
            Else
                myVars.Item("Rev").Value = Trim(UCase(txtVer.Text))
            End If
            myVars.Item("Type").Value = Trim(txtXH.Text)
   
 
            If Me.chkN.Value = 1 Then
                myVars.Item("Rohs").Value = "N*"
            ElseIf Me.chkN4.Value = 1 Then
                myVars.Item("Rohs").Value = "N4"
            ElseIf Me.chkY.Value = 1 Then
                myVars.Item("Rohs").Value = "Y*"
            ElseIf Me.chkY2.Value = 1 Then
                myVars.Item("Rohs").Value = "Y2"
            End If
 
  
            'myApp.Visible = True
            myDoc.PrintLabel 1
            myDoc.FormFeed
            
      '===============add by Carson 2015-12-15 start===============
            If normalSNFlag = True Then
                Call Connect.addPrintedLabel(strSN1, Me.Name)
            End If
            
            If normalSNFlag1 = True Then
                Call Connect.addPrintedLabel(strSN2, Me.Name)
            End If
            
            If normalSNFlag2 = True Then
                Call Connect.addPrintedLabel(strSN3, Me.Name)
            End If
            
            If normalSNFlag3 = True Then
                Call Connect.addPrintedLabel(strSN4, Me.Name)
            End If
            
            If normalSNFlag4 = True Then
                Call Connect.addPrintedLabel(strSN5, Me.Name)
            End If
            
          If normalSNFlag5 = True Then
                Call Connect.addPrintedLabel(strSN6, Me.Name)
            End If

        '===============add by Carson 2015-12-15 end=================
            'Call Connect.addPrintedLabel(Trim(strSN1), Me.Name)
            k = k + 1
            
            DoEvents
    Else
        While (bRun = False)
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

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdStop_Click()
    bRun = False
    cmdPrint.Enabled = False
    cmdCancel.Enabled = False
    cmdReturn.Enabled = False
    cmdStop.Enabled = False
    cmdGoon.Enabled = True
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
   bRun = False
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
        txtVer.Enabled = False
        If Len(txtSN.Text) <> 16 Then
           MsgBox "产品序号长度必须为16位!"
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
        
        End If
        rec.Close
        txtWO.SetFocus
    Else
        txtWO.Text = ""
        txtCPN.Text = ""
        txtVer.Text = ""
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
   Set myDoc = myApp.Documents.Open("\\10.11.1.25\Public\Manufacture\打印中心\" & "0303二维码标签8X8.Lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub


Private Sub txtWO_KeyPress(KeyAscii As Integer)
Dim tempWO As String
Dim partlist As String
    If (KeyAscii = 13) Then
        If Len(Trim(txtCPN.Text)) <> 8 Then
            MsgBox "产品编码长度必须为8位!"
            txtSN.SetFocus
            Exit Sub
        Else
            tempWO = txtWO.Text
            If tempWO = "" Or tempWO = Null Then Return
            If UCase(tempWO) = "TASK" Then
                txtVer.Enabled = True
                chkY.Enabled = True
                chkY2.Enabled = True
                chkN.Enabled = True
                chkN4.Enabled = True
                Exit Sub
            End If
            sql = "select part_revision,(select order_type_S from [10.11.1.130].afg_active_90.dbo.UDA_Order where object_key=A.order_key) order_type from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A,[10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B " & _
                "WHERE A.order_key = B.order_key AND A.order_number ='" & tempWO & "' and ( part_number like 'HWF" & txtCPN.Text & "%')"
            rec.Open sql, conn, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                MsgBox "SAP中此工单的编码号与此产品编码不一致或者该工单是HUV工单!"
                txtWO.Text = ""
                txtVer.Text = ""
                txtWO.SetFocus
                rec.Close
                Exit Sub
            Else
                txtVer.Text = Trim(rec.Fields(0))
                If rec.Fields(1) = "PP05" Then
                    txtVer.Enabled = True
                    chkY.Enabled = True
                    chkY2.Enabled = True
                    chkN.Enabled = True
                    chkN4.Enabled = True
                    rec.Close
                    Exit Sub
                End If
                rec.Close
                partlist = Connect.getPartListByOrder(tempWO)
                If partlist = "" Then
                    MsgBox "找不到改工单" & tempWO & "对应的0302阶或当前阶的信息,请确认SAP是否释放改工单并且传送MES系统中"
                    txtWO.Text = ""
                    txtVer.Text = ""
                    txtWO.SetFocus
                    Exit Sub
                End If
                Dim cmd As New ADODB.Command
                cmd.ActiveConnection = conn
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "[PbHandler]"
                cmd.Parameters.Append cmd.CreateParameter("partlist", adVarChar, adParamInput, 8000, partlist)
                cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 8)
                cmd.Parameters.Append cmd.CreateParameter("first", adVarChar, adParamOutput, 1)
                cmd.Execute
'                Me.tbFirst.Text = cmd("first")      ''Cancel the 9 principle
                Me.tbFirst.Text = "0"
                Select Case cmd("res")
                    Case "No"
                        Me.chkY2.Value = 1
                        Me.chkY2.Enabled = False
                        Me.chkY.Value = 0
                        Me.chkY.Enabled = False
                        Me.chkN4.Value = 0
                        Me.chkN4.Enabled = False
                        Me.chkN.Value = 0
                        Me.chkN.Enabled = False
                    Case "Non"
                        MsgBox "此工单未设定有铅无铅,请相关ME去设定!"
                        txtWO.Text = ""
                        txtVer.Text = ""
                        txtWO.SetFocus
                        Exit Sub
                    Case "Half"
                       chkY.Enabled = False
                       chkY.Value = 0
                       chkY2.Value = 0
                       chkY2.Enabled = False
                       chkN4.Value = 1
                       chkN4.Enabled = False
                       chkN.Value = 0
                       chkN.Enabled = False
                    Case "Yes"
                       chkY.Value = 0
                       chkY.Enabled = False
                       chkY2.Value = 0
                       chkY2.Enabled = False
                       chkN4.Value = 1
                       chkN4.Enabled = False
                       chkN.Value = 0
                       chkN.Enabled = False
                End Select
                OverridePb      ''check label history
            End If
        End If
    Else
        txtVer.Text = ""
    End If
End Sub


Private Sub OverridePb()
    Dim labelHistory As New Label_History
    Dim sn As String
    sn = txtSN.Text
    If labelHistory.Init(sn) Then
    
        chkY.Value = 0
        chkY2.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
        
        Select Case labelHistory.Pb
        Case "Y*"
            chkY.Value = 1
        Case "Y2"
            chkY2.Value = 1
        Case "N*"
            chkN.Value = 1
        Case "N4"
            chkN4.Value = 1
        End Select
        
    End If
End Sub
