VERSION 5.00
Begin VB.Form frmHWBatchPrint 
   Caption         =   "非H3C批量打印"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   ScaleHeight     =   5730
   ScaleWidth      =   10680
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      Picture         =   "frmHWBatchPrint.frx":0000
      ScaleHeight     =   1545
      ScaleWidth      =   10425
      TabIndex        =   19
      Top             =   0
      Width           =   10455
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   10455
      Begin VB.TextBox txtVer 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1680
         TabIndex        =   11
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   1485
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtQty1 
         Enabled         =   0   'False
         Height          =   405
         Left            =   6720
         TabIndex        =   8
         Text            =   "1"
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Non-Rohs"
         Height          =   375
         Left            =   8160
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chkChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rohs"
         Height          =   375
         Left            =   6840
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtXH 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   9000
         TabIndex        =   4
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本:"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "一式几份:"
         Height          =   375
         Left            =   5040
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "起始条码:"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "环保属性:"
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品型号:"
         Height          =   375
         Left            =   5160
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "打印数量:"
         Height          =   375
         Left            =   7680
         TabIndex        =   12
         Top             =   1920
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   6840
      TabIndex        =   1
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   4920
      Width           =   1815
   End
End
Attribute VB_Name = "frmHWBatchPrint"
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
   txtQty1.Text = ""
'   chkCE.Value = 0
   
'   chkWEEE.Value = 0
 
'   chkRoHS.Value = 0
   chkNonChinaRoHS.Value = 0
  chkChinaRoHS.Value = 0
 
 
   txtSN.SetFocus
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
      txtVer.SetFocus
      Exit Sub
   End If
   
   If txtXH.Text = "" Then
      MsgBox "型号未输入,不能打印!", vbInformation + vbOKOnly, "未输入型号"
      txtXH.SetFocus
      Exit Sub
   End If
   Dim barcode() As String
   Dim str As String
   barcode = Split(Trim(txtSN.Text), vbCrLf, , vbTextCompare)
   OpenLppx
   For i = 0 To UBound(barcode)
      str = barcode(i)
      If Len(str) = 16 Then
            myVars.Item("sn").Value = Trim(str)
            If txtVer.Text = "" Or txtVer.Text = "/" Then
               myVars.Item("ver").Value = "N/A"
            Else
              myVars.Item("ver").Value = Trim(UCase(txtVer.Text))
            End If
            myVars.Item("Type").Value = Trim(txtXH.Text)
            If chkChinaRoHS.Value = 1 Then
           If Left(str, 2) = "03" Or Left(str, 2) = "21" Then
             myVars.Item("Rohs").Value = "Y*"
           End If
           If Left(str, 2) = "02" Then
             myVars.Item("Rohs").Value = "Y2"
           End If
        Else
            If Left(str, 2) = "03" Or Left(str, 2) = "21" Then
             myVars.Item("Rohs").Value = "N*"
           End If
           If Left(str, 2) = "02" Then
             myVars.Item("Rohs").Value = "Y1"
           End If
        End If
         myDoc.PrintLabel 1
         myDoc.FormFeed
    End If
   Next
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
     txtVer.SetFocus
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
        barcode = Split(Trim(txtSN.Text), vbCrLf, , vbTextCompare)
        If UBound(barcode) = 0 Then
           txtSN.Text = ""
           txtSN.SetFocus
           Exit Sub
        End If
        str = barcode(0)

        Dim rcd As New ADODB.Recordset
        sql = "select * from tblCustomType where PartNumber='" & Mid(str, 3, 8) & "'"
        rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rcd.EOF = True Then
           MsgBox "品牌未维护!"
           txtSN.Text = ""
           txtSN.SetFocus
           rcd.Close
           Exit Sub
        Else
            If rcd.Fields(1) = "H3C" Then
                MsgBox "请使用[H3C整机模块类标签打印]打印!"
                txtSN.Text = ""
                txtSN.SetFocus
                rcd.Close
                Exit Sub
            End If
        End If
        rcd.Close
      
        
      sql = "select * from SingleUnit where SN='" & Mid(str, 3, 8) & "'"
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
        
        chkChinaRoHS.Value = 0
        chkNonChinaRoHS.Value = 0
        

'        If UCase(Trim(rec.Fields(3))) = "ROHS" Then
'           chkChinaRoHS.Value = 1
'           chkNonChinaRoHS.Value = 0
'        ElseIf rec.Fields(3) = "/" Then
'           chkChinaRoHS.Value = 0
'           chkNonChinaRoHS.Value = 1
'        End If
       
    
      End If
      rec.Close
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
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "21HUAWEI.lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub txtVer_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtHV.SetFocus
  End If
End Sub

