VERSION 5.00
Begin VB.Form frm0302Print 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "0302类标签打印"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm0302Print.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10665
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdStop 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7920
      TabIndex        =   22
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdGoon 
      Caption         =   "继续"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9360
      TabIndex        =   21
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   480
      TabIndex        =   13
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   5640
      TabIndex        =   12
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   3120
      TabIndex        =   11
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   10455
      Begin VB.TextBox txtWO 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1680
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkY3 
         BackColor       =   &H80000009&
         Caption         =   "Y3"
         Height          =   330
         Left            =   3840
         TabIndex        =   20
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox chkY2 
         BackColor       =   &H80000009&
         Caption         =   "Y2"
         Height          =   330
         Left            =   2760
         TabIndex        =   19
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox chkY1 
         BackColor       =   &H80000009&
         Caption         =   "Y1"
         Height          =   330
         Left            =   1680
         TabIndex        =   18
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   9120
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtXH 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   15
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtQty1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   405
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
         TabIndex        =   23
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "打印数量:"
         Height          =   375
         Left            =   7800
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品型号:"
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "环保属性:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "起始条码:"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "一式几份:"
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本:"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      Picture         =   "frm0302Print.frx":13652
      ScaleHeight     =   1545
      ScaleWidth      =   10425
      TabIndex        =   4
      Top             =   480
      Width           =   10455
   End
End
Attribute VB_Name = "frm0302Print"
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

Private Sub chkY1_Click()
 If chkY1.Value = 1 Then
      chkY2.Value = 0
      chkY3.Value = 0
   Else
      chkY2.Value = 0
      chkY3.Value = 0
   End If
End Sub

Private Sub chkY2_Click()
   If chkY2.Value = 1 Then
      chkY1.Value = 0
      chkY3.Value = 0
   Else
      chkY3.Value = 0
      chkY1.Value = 0
   End If
End Sub

Private Sub chkY3_Click()
   If chkY3.Value = 1 Then
      chkY1.Value = 0
      chkY2.Value = 0
   Else
      chkY1.Value = 0
      chkY2.Value = 0
   End If
End Sub

Private Sub cmdCancel_Click()
   txtSN.Text = ""
   txtVer.Text = ""
   txtCPN.Text = ""
   txtXH.Text = ""
  
   txtQty.Text = ""
  ' txtQty1.Text = ""
'   chkCE.Value = 0
'   chkWEEE.Value = 0
'   chkRoHS.Value = 0
'   chkNonChinaRoHS.Value = 0
   chkY1.Value = 0
   chkY2.Value = 0
   chkY3.Value = 0
 
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
   
   If chkY1.Value = 0 And chkY2.Value = 0 And chkY3.Value = 0 Then
        MsgBox "环保属性未选择,不能打印!", vbInformation + vbOKOnly, "未输入型号"
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
   rightstr = Right(txtSN.Text, 6)
     OpenLppx
     
     bRun = True
     Dim k As Integer
     k = 0
     
   For i = 0 To qty - 1
      str = leftstr & Right("000000" & CStr(CInt(rightstr) + i), 6)

    For j = 0 To qty1 - 1
        If bRun = True Then
            If k > 0 And k Mod 100 = 0 Then
                Savetime = timeGetTime '记下开始时的时间
                While timeGetTime < Savetime + 30000 '循环等待
                    DoEvents '转让控制权，以便让操作系统处理其它的事件。
                Wend
            End If
keepprint:
            myVars.Item("sn").Value = Trim(str)
            'myVars.Item("Item").Value = "03" & UCase(Left(txtSN.Text, 6))
            If txtVer.Text = "" Or txtVer.Text = "/" Then
                'myObjs("Sver").Top = 5
                myVars.Item("ver").Value = "N/A"
            ElseIf txtVer.Text = "00" Then
                myVars.Item("ver").Value = ""
            Else
                'myObjs("Sver").Top = 5
                myVars.Item("ver").Value = Trim(UCase(txtVer.Text))
            End If
            myVars.Item("Type").Value = Trim(txtXH.Text)

            If chkY1.Value = 1 Then
                myVars.Item("Rohs").Value = "Y1"
            End If
  
            If chkY2.Value = 1 Then
                myVars.Item("Rohs").Value = "Y2"
            End If
 
            If chkY3.Value = 1 Then
                myVars.Item("Rohs").Value = "Y3"
            End If
  
            'myApp.Visible = True
'            myDoc.CopyToClipboard
            myDoc.PrintLabel 1
            myDoc.FormFeed
   
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
     txtVer.SetFocus
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
      If Len(txtSN.Text) <> 16 Then
         MsgBox "产品序号长度必须为16位!"
         txtSN.SetFocus
         Exit Sub
      End If

'If Left(txtSN.Text, 2) = "03" Or Left(txtSN.Text, 2) = "21" Then
 '  chkChinaRoHS.Caption = "Rohs"
  ' chkNonChinaRoHS.Caption = "Non-Rohs"
'ElseIf Left(txtSN.Text, 2) = "02" Then
 '  chkChinaRoHS.Caption = "无铅"
  ' chkNonChinaRoHS.Caption = "有铅"
'End If


If Left(txtSN.Text, 2) = "21" Then
   sql = "select * from SingleUnit where SN= '" & Mid(txtSN.Text, 3, 8) & "'"
Else
   sql = "select * from SingleUnit where SN='03" & Mid(txtSN.Text, 1, 6) & "'"
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

'        If UCase(Trim(rec.Fields(3))) = "ROHS" Then
'           chkChinaRoHS.Value = 1
'           chkNonChinaRoHS.Value = 0
'        ElseIf rec.Fields(3) = "/" Then
'           chkChinaRoHS.Value = 0
'           chkNonChinaRoHS.Value = 1
'        End If
       
    
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
   Set myDoc = myApp.Documents.Open("\\10.11.1.25\Public\Manufacture\标签模板\" & "0302_SN.lab")
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
'            While (Len(tempWO) < 12)
'                tempWO = "0" & tempWO
'            Wend
'            sql = "select MaterialRevision from [10.11.1.17].dsActive.dbo.SAP_WO " & _
'                "where WorkOrderNumber = '" & tempWO & "' and MaterialNumber = 'HWF" & txtCPN.Text & "' "
            sql = "select part_revision from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A ,[10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B " & _
                "WHERE A.order_key = B.order_key AND A.order_number ='" & tempWO & "' and ( part_number like 'HWF" & txtCPN.Text & "%' " & _
                "or part_number like 'HUV" & txtCPN.Text & "' or part_number like 'HWFC" & txtCPN.Text & "')"
            rec.Open sql, conn, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                MsgBox "SAP中此工单的编码号与此产品编码不一致!"
                txtWO.Text = ""
                txtVer.Text = ""
'                txtSN.Text = ""
'                txtCPN.Text = ""
'                txtXH.Text = ""
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

