VERSION 5.00
Begin VB.Form frmHPSelfPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HP本体标签"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHPSelfPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5160
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdGoon 
      Caption         =   "继续"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3960
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   6360
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtWO 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1680
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtPN 
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
         Caption         =   "工单号"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "打印数量"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblChinaRoHS 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "PCS"
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "机种号"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本:"
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      Picture         =   "frmHPSelfPrint.frx":13652
      ScaleHeight     =   1665
      ScaleWidth      =   2985
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmHPSelfPrint"
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
Private Sub cmdCancel_Click()
   txtVer.Text = ""
   txtPN.Text = ""
   txtQty.Text = ""
   txtPN.SetFocus
  ' txtQty1.Text = ""
'   chkCE.Value = 0
'   chkWEEE.Value = 0
'   chkRoHS.Value = 0
'   chkNonChinaRoHS.Value = 0
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
   If txtPN.Text = "" Then
      MsgBox "机种号未输入,不能打印!", vbInformation + vbOKOnly, "未输入产品条码"
      txtPN.SetFocus
      Exit Sub
   End If
   
   If txtVer.Text = "" Then
      MsgBox "版本未输入,不能打印!", vbInformation + vbOKOnly, "未输入版本"
      txtWO.SetFocus
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
   
   cmdPrint.Caption = "执行中..."
   cmdPrint.Enabled = False
   cmdStop.Enabled = True
    
   Dim i, j, qty As Integer
   Dim pn, rev As String
   qty = CInt(txtQty.Text)
   pn = UCase(Trim(txtPN.Text))
   rev = Trim(txtVer.Text)

   OpenLppx
     
   bRun = True
   Dim k As Integer
   k = 0
     
   For i = 0 To qty - 1
'   For j = 0 To qty1 - 1
        If bRun = True Then
            If k > 0 And k Mod 100 = 0 Then
                Savetime = timeGetTime '记下开始时的时间
                While timeGetTime < Savetime + 30000 '循环等待
                    DoEvents '转让控制权，以便让操作系统处理其它的事件。
                Wend
            End If
keepprint:
            myVars.Item("PN").Value = Trim(pn)
            If rev = "" Or rev = "/" Then
                'myObjs("Sver").Top = 5
                myVars.Item("Rev").Value = "N/A"
            ElseIf rev = "00" Then
                myVars.Item("Rev").Value = ""
            Else
                'myObjs("Sver").Top = 5
                myVars.Item("Rev").Value = Trim(rev)
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
'   Next
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
Private Sub txtPN_GotFocus()
    txtVer.Text = ""
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

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\10.11.1.25\Public\Manufacture\标签模板\" & "HP本体标签.Lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub
Private Sub txtWO_GotFocus()
    txtVer.Text = ""
End Sub

Private Sub txtWO_KeyPress(KeyAscii As Integer)
Dim tempWO As String
    If (KeyAscii = 13) Then
        If Len(Trim(txtPN.Text)) <> 8 Then
            MsgBox "产品编码长度必须为8位!"
            txtSN.SetFocus
            Exit Sub
        Else
            tempWO = txtWO.Text
            While (Len(tempWO) < 12)
                tempWO = "0" & tempWO
            Wend
'            sql = "select MaterialRevision from [10.11.1.17].dsActive.dbo.SAP_WO " & _
'                "where WorkOrderNumber = '" & tempWO & "' and MaterialNumber = 'HWF" & txtCPN.Text & "' "
            sql = "select MaterialRevision from [10.11.1.17].dsActive.dbo.SAP_WO " & _
                "where WorkOrderNumber = '" & tempWO & "' and ( MaterialNumber like 'HWF%" & txtPN.Text & "%' " & _
                "or MaterialNumber like 'HUV" & txtPN.Text & "%' ) "
'            sql = "select MaterialRevision from [10.11.1.17].dsActive.dbo.SAP_WO " & _
'                "where WorkOrderNumber = '" & tempWO & "' and MaterialNumber like 'HWF%" & txtPN.Text & "%' "
            rec.Open sql, conn, adOpenKeyset, adLockOptimistic
            If rec.EOF = True Then
                MsgBox "SAP中此工单的编码号与此产品编码不一致!"
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

