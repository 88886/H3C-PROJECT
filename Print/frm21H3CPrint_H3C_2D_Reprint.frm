VERSION 5.00
Begin VB.Form frm21H3CPrint_H3C_2D_Reprint 
   Caption         =   "H3C SN&MAC地址补印"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   11445
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   7680
      TabIndex        =   9
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   4560
      TabIndex        =   8
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   1560
      Picture         =   "frm21H3CPrint_H3C_2D_Reprint.frx":0000
      TabIndex        =   7
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   10815
      Begin VB.TextBox txtAutoTest 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6000
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtXH 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   960
         TabIndex        =   17
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtPB 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   960
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtVer 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3000
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtSN 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtMac 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5520
         TabIndex        =   10
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtQty1 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7560
         TabIndex        =   3
         Text            =   "1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5520
         TabIndex        =   2
         Text            =   "1"
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "自动测试机柜:"
         Height          =   255
         Left            =   4680
         TabIndex        =   20
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品型号:"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本:"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "环保属性:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品条码:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "一式几份:"
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "数量:"
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MAC地址:"
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   240
      Picture         =   "frm21H3CPrint_H3C_2D_Reprint.frx":13652
      ScaleHeight     =   2235
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "frm21H3CPrint_H3C_2D_Reprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim bRun As Boolean

Private Sub cmdReturn_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   bRun = False
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\10.11.1.25\Public\Manufacture\打印中心\" & "整机SN标签40x15x15.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub


Private Sub cmdPrint_Click()
    If txtSN.Text = "" Then
       MsgBox "产品条码未输入,不能打印!", vbInformation + vbOKOnly, "未输入产品条码"
       txtSN.SetFocus
       Exit Sub
    End If

'    If txtMac.Text = "" Or Len(txtMac.Text) < 12 Then
'          MsgBox "未输入正确的MAC！", vbInformation + vbOKOnly, "未输入正确的MAC"
'          txtMac.SetFocus
'          Exit Sub
'    End If
    
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
   
   If txtPB.Text = "" Then
      MsgBox "环保属性未输入,不能打印!", vbInformation + vbOKOnly, "未输入环保属性"
      txtPB.SetFocus
      Exit Sub
   End If
    
    If txtQty.Text = "" Then
       MsgBox "未输入数量！", vbInformation + vbOKOnly, "未输入数量"
       Exit Sub
    End If
    If txtQty1.Text = "" Then
       MsgBox "未输入一式几份数量！", vbInformation + vbOKOnly, "未输入一式几份数量"
       Exit Sub
    End If
    
    If CInt(txtQty.Text) = 0 Then
       MsgBox "请输入正确的数量！", vbInformation + vbOKOnly, "数量不对"
       Exit Sub
    End If
    If CInt(txtQty1.Text) = 0 Then
       MsgBox "请输入正确的一式几份数量！", vbInformation + vbOKOnly, "一式几份数量不对"
       Exit Sub
    End If
    
    
    If rec.State = 1 Then
        rec.Close
    End If
    sql = "select a.serial_number,a.MAC,b.software_version,b.pb,a.SingleUnit_Type from C_MACAndSN_PrintRecord a left join H3C_PB b on a.serial_number=b.serial_number where EFFE_FLAG='1' and a.MAC= '" & Trim(Me.txtMac.Text) & "' AND a.serial_number='" & Trim(Me.txtSN.Text) & "'  "
    If connFTPC.State = 0 Then
       connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
       connFTPC.Open
    End If
    rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
    If rec.EOF = True Then
        MsgBox "该MAC和SN打印记录不匹配，请与ME确认!"
        txtSN.Text = ""
        txtVer.Text = ""
        txtPB.Text = ""
        txtXH.Text = ""
        txtAutoTest.Text = ""
        Exit Sub
    End If
    rec.Close
    connFTPC.Close
   qty = CInt(txtQty.Text)
   qty1 = CInt(txtQty1.Text)
    
    OpenLppx
''''''''''qty 为打印数量
    bRun = True
    For i = 0 To qty - 1

''''''''''qty1 为一式几份
        For j = 0 To qty1 - 1
     
            If bRun = True Then
                If k > 0 And k Mod 100 = 0 Then
                    Savetime = timeGetTime '记下开始时的时间
                    While timeGetTime < Savetime + 30000 '循环等待
                        DoEvents '转让控制权，以便让操作系统处理其它的事件。
                    Wend
                End If
keepprint:
                myVars.Item("sn").Value = Trim(txtSN.Text)
                If txtVer.Text = "" Or txtVer.Text = "/" Then
                    myVars.Item("rev").Value = "N/A"
                ElseIf Me.txtVer.Text <> "" Then
                    myVars.Item("rev").Value = Trim(txtVer.Text)
                Else
                    myVars.Item("rev").Value = UCase(txtVer.Text)
                End If
                myVars.Item("PID").Value = txtXH.Text
                myVars.Item("Rohs").Value = txtPB.Text
     
                'myApp.Visible = True
                If Trim(txtMac.Text) = "" Or Len(Trim(txtMac.Text)) <> 12 Then
                    mac = ""
                    myObjs("text3").Top = 10000
                    myObjs("MAC").Top = 10000
                    myObjs("MAC1").Top = 10000
                    
                    If txtAutoTest.Text <> "Y" Then
                        myObjs("SN&MAC").Top = 10000
                        myObjs("SN2").Top = 10000
                        myObjs("MAC(2)").Top = 10000
                    End If
                Else
                    mac = Trim(txtMac.Text)
                    myVars.Item("MAC").Value = mac
                    If txtAutoTest.Text <> "Y" Then
                        myObjs("SN&MAC").Top = 10000
                        myObjs("SN2").Top = 10000
                        myObjs("MAC(2)").Top = 10000
                    End If
                End If
            
                myDoc.PrintLabel 1
                myDoc.FormFeed
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
Private Sub cmdCancel_Click()
    Dim Ctr As Control
    For Each Ctr In Me.Controls
        If TypeOf Ctr Is TextBox Then
            Ctr.Text = ""
        End If
        If TypeOf Ctr Is CheckBox Then
            Ctr.Value = 0
        End If
    Next
   txtQty.Text = "1"
   txtQty1.Text = "1"
   txtSN.SetFocus
End Sub

Private Sub txtMac_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If rec.State = 1 Then
            rec.Close
        End If
        sql = "select a.serial_number,a.MAC,b.software_version,b.pb,a.SingleUnit_Type,isnull(a.AutoTest,'') from C_MACAndSN_PrintRecord a left join H3C_PB b on a.serial_number=b.serial_number where EFFE_FLAG='1' and a.MAC= '" & Trim(Me.txtMac.Text) & "'  "
        If connFTPC.State = 0 Then
           connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
           connFTPC.Open
        End If
        rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
        If rec.EOF = True Then
            MsgBox "该MAC没有打印过，请确认!"
            txtSN.Text = ""
            txtVer.Text = ""
            txtPB.Text = ""
            txtXH.Text = ""
            txtAutoTest.Text = ""
            Exit Sub
        Else
          txtMac.Text = Trim(rec.Fields(1))
          txtVer.Text = Trim(rec.Fields(2))
          txtPB.Text = Trim(rec.Fields(3))
          txtXH.Text = Trim(rec.Fields(4))
          txtAutoTest.Text = Trim(rec.Fields(5))
        End If
        rec.Close
        connFTPC.Close
        txtSN.SetFocus
    Else
        txtMac.Text = ""
        txtVer.Text = ""
        txtPB.Text = ""
        txtSN.Text = ""
        txtXH.Text = ""
        txtAutoTest.Text = ""
        txtMac.SetFocus
    End If
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If rec.State = 1 Then
            rec.Close
        End If
          
       'sql = "select a.serial_number,a.MAC,b.software_version,b.pb,a.SingleUnit_Type,isnull(a.AutoTest,'') from C_MACAndSN_PrintRecord a left join H3C_PB b on a.serial_number=b.serial_number  where EFFE_FLAG='1' and a.serial_number= '" & Trim(Me.txtSN.Text) & "'  "
        
       'Update By Robin 2018.4.19 硬件版本修改
       sql = "select a.serial_number,a.MAC,u.version,b.pb,a.SingleUnit_Type,isnull(a.AutoTest,'') from C_MACAndSN_PrintRecord a left join H3C_PB b on a.serial_number=b.serial_number  left join H3C_PB_Version u on u.serial_number = a.serial_number  where EFFE_FLAG='1' and a.serial_number= '" & Trim(Me.txtSN.Text) & "'  "
       
        
        If connFTPC.State = 0 Then
           connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
           connFTPC.Open
        End If
        rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
        If rec.EOF = True Then
            MsgBox "该流水号没有打印过，请确认!"
            txtVer.Text = ""
            txtSN.SetFocus
            Exit Sub
        Else
          txtMac.Text = Trim(rec.Fields(1))
          txtVer.Text = Trim(rec.Fields(2))
          txtPB.Text = Trim(rec.Fields(3))
          txtXH.Text = Trim(rec.Fields(4))
          txtAutoTest.Text = Trim(rec.Fields(5))
        End If
        rec.Close
        connFTPC.Close
        txtSN.SetFocus
    Else
        txtMac.Text = ""
        txtVer.Text = ""
        txtXH.Text = ""
        txtPB.Text = ""
        txtSN.Text = ""
        txtAutoTest.Text = ""
        txtSN.SetFocus
    End If
End Sub
