VERSION 5.00
Begin VB.Form frmH3CPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H3C Label Print"
   ClientHeight    =   8910
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
   ScaleHeight     =   8910
   ScaleWidth      =   14130
   StartUpPosition =   2  '屏幕中心
   Begin VB.ListBox lstPrinter 
      Height          =   390
      Left            =   10440
      TabIndex        =   36
      Top             =   3000
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   1440
      Picture         =   "frmH3CPrint.frx":13652
      ScaleHeight     =   3435
      ScaleWidth      =   8595
      TabIndex        =   32
      Top             =   120
      Width           =   8655
   End
   Begin VB.TextBox lblNALday 
      Height          =   450
      Left            =   480
      TabIndex        =   31
      Top             =   8160
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox lblMSday 
      Height          =   450
      Left            =   120
      TabIndex        =   30
      Top             =   8160
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   4080
      TabIndex        =   24
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   9000
      TabIndex        =   23
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   6480
      TabIndex        =   22
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   0
      TabIndex        =   12
      Top             =   3840
      Width           =   13935
      Begin VB.TextBox txtPowerCode 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   40
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtEC 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   10920
         TabIndex        =   38
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8325
         TabIndex        =   29
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox chkRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7605
         TabIndex        =   28
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox chkVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本信息:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   27
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "外尺寸(MM):"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   25
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   10
         Top             =   3480
         Width           =   3135
      End
      Begin VB.TextBox txtNAL 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10560
         TabIndex        =   8
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtMS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   7
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtHV 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   9
         Top             =   3480
         Width           =   3135
      End
      Begin VB.TextBox txtGW 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10920
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtOS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   5895
      End
      Begin VB.TextBox txtDes 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   4
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtEPN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7560
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtVer 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7560
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "电源代码:"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "电流(A):"
         Height          =   375
         Left            =   9720
         TabIndex        =   37
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbPath 
         Height          =   375
         Left            =   2400
         TabIndex        =   34
         Top             =   2400
         Width           =   11295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Logo 图片路径:"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息RoHS:"
         Height          =   375
         Left            =   5400
         TabIndex        =   26
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
         Height          =   375
         Left            =   8400
         TabIndex        =   21
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "进网许可号:"
         Height          =   375
         Left            =   8400
         TabIndex        =   20
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label lblMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "执行标准:"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label lblHV 
         BackColor       =   &H00FFFFFF&
         Caption         =   "硬件版本:"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblGW 
         BackColor       =   &H00FFFFFF&
         Caption         =   "毛重(kg):"
         Height          =   375
         Left            =   9600
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(英文):"
         Height          =   375
         Left            =   5400
         TabIndex        =   16
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品名称(中文):"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品描述:"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.Label Label2 
      Caption         =   "打印机设定:"
      Height          =   375
      Left            =   10440
      TabIndex        =   35
      Top             =   2400
      Width           =   1935
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
      TabIndex        =   11
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
   chkRoHS.Value = 0
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
'        If Connect.isPrintedLabel(Me.txtSN.Text, Me.Name) Then
'            MsgBox ("此序列号已打印！")
'            txtSN.SetFocus
'            Exit Sub
'        End If
    End If
'===============add by ben 2012-02-05 end=================
   
   OpenLppx
   
   myVars.Item("SN").Value = UCase(txtSN.Text)
   myVars.Item("Item").Value = Mid(UCase(txtSN.Text), 3, 8) + "*1PCS"
   
   If chkVer.Value = 0 Then
      myObjs("BSver").Top = 10000
      myVars.Item("SVer").Value = "N/A"
   Else
       If txtVer.Text = "" Or txtVer.Text = "/" Or txtVer.Text = "N/A" Then
         myVars.Item("SW").Value = "N/A"
       Else
         myVars.Item("SW").Value = UCase(Trim(Replace(txtVer.Text, vbCrLf, "")))
       End If
   End If
   
   myVars.Item("CPD").Value = UCase(txtCPN.Text)
   myVars.Item("EPD").Value = txtEPN.Text
   myVars.Item("PID").Value = txtDes.Text
   myVars.Item("2D").Value = "SN:210235A43B0126000001" & vbCrLf & "PlD: NS -Navigator2 - 2" & vbCrLf & "SW:5116P20" & "HW: E0" & "VENDOR: POW001"
   myVars.Item("OD").Value = Me.txtOS.Text
'   myVars.Item("SW").Value =


   
   myVars.Item("MS").Value = txtMS.Text
   myVars.Item("NAL").Value = UCase(txtNAL.Text)
   
   If txtHV.Text = "" Or txtHV.Text = "/" Or txtHV.Text = "N/A" Then
      myObjs("BHver").Top = 10000
      myVars.Item("HVer").Value = "N/A"
   Else
'      myObjs("THver").Top = 10000
'      myVars.Item("HVer").Value = UCase(Trim(Replace(txtHV.Text, vbCrLf, "")))
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
    FormHPFahuo.lbPrinter.Caption = lstPrinter.Text
    
    FormHPFahuo.Show
    Call FormHPFahuo.cmdMPrint_Click
   End If

End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Dim PrintData As Printer
    Dim defprinterpos%
    For Each PrintData In Printers
        lstPrinter.AddItem PrintData.DeviceName
    ' Check for default printer7.
        If PrintData.DeviceName = Printer.DeviceName Then
            defprinterpos = lstPrinter.NewIndex
        End If
    
    Next
    lstPrinter.ListIndex = defprinterpos%
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

        If conn11.State = 1 Then
             conn11.Close
        End If
        
        strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
            'con13.ConnectionTimeout = 50
        conn11.Open ConnectionString:=strConn
'        conn11.Open
        sql = " SELECT component_SN from dc_component_sn where unit_key IN (SELECT unit_key from UNIT WHERE SERIAL_NUMBER = '" & Trim(txtSN.Text) & "')" & " AND Remark = 'HP'"
        checkhp.Open sql, conn11, adOpenKeyset, adLockOptimistic
        If checkhp.EOF = True Then
            MsgBox ("没有对应的HP条码！")
            txtSN.Text = ""
            txtSN.SetFocus
            checkhp.Close
            Exit Sub
        Else
            hpsn = checkhp.Fields(0)
            checkhp.Close
        End If
    

        If conn11.State = 1 Then
             conn11.Close
        End If
        
      End If
      
    'add by allen
    'for add a power vendor code check function
    
      
      Dim powerset As New ADODB.Recordset
      strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
      conn11.Open ConnectionString:=strConn
      sql = "select d.component_SN,b.Rule1_char,Power_Length,ISNULL(B.Power_Vendor_Code,'') from Power_Supply a,Power_Supply_Regulation b,UNIT c, DC_Component_SN d " & _
"where a.Material = b.Material and a.is_valid = '1' and b.is_valid = '1' and a.Part_Number = c.part_number AND c.unit_key = d.object_key and d.Remark = 'PowerSupply' and c.serial_number = '" & Trim(txtSN.Text) & "'"
      powerset.Open sql, conn11, adOpenKeyset, adLockReadOnly
      If powerset.EOF = True Then
        MsgBox ("没有电源供应商对应的代码！")
            txtSN.Text = ""
            txtSN.SetFocus
            powerset.Close
            Exit Sub
      Else
        If Len(powerset.Fields(0)) = CInt(powerset.Fields(2)) Then
            Me.txtPowerCode.Text = Trim(powerset.Fields(3))
        Else
            MsgBox ("没有电源供应商对应的代码长度与系统设定不一致！")
            txtSN.Text = ""
            txtSN.SetFocus
            powerset.Close
            Exit Sub
        End If
        
      End If
      If conn11.State = 1 Then
             conn11.Close
      End If
      
      
      '=========================================================================
            Dim con13 As ADODB.Connection
            Dim rs13 As ADODB.Recordset
            Dim com As ADODB.Command

            Set con13 = New ADODB.Connection
            Set rs13 = New ADODB.Recordset
            strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
            'con13.ConnectionTimeout = 50
            con13.Open ConnectionString:=strConn
            Set com = New ADODB.Command
            com.ActiveConnection = con13
            'str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtSN.Text) & "'"
            'str = " select top 1 a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "'"
            str = " select top 1 part_number,part_revision,creation_time,order_number from (" & _
            "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
            "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
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
'                Dim str As String
                Dim part_number As String
                Dim part_revision As String
                Dim order_number As String
                
                Set rs3.ActiveConnection = con3
                rs3.CursorType = adOpenDynamic
                
                str = " select top 1 part_number,part_revision,creation_time,order_number from (" & _
                " select a.part_number,a.part_revision,a.creation_time,(select order_number from work_order with(NOLOCK) where order_key=a.order_key) as order_number from dsactive.dbo.unit a with(nolock) " & _
                " where a.serial_number='" & Trim(txtSN.Text) & "'" & _
                " union" & _
                " select part_number,part_rev as part_revision,creation_time,order_number from dsactive.dbo.dc_task_order NOLOCK  " & _
                " where order_number=(select order_number from dsactive.dbo.taskorder_unit NOLOCK" & _
                " where serial_number='" & Trim(txtSN.Text) & "')" & _
                " ) as t " & _
                " order by t.creation_time desc"
                
                rs3.Open str
                If rs3.EOF = True Then
                    MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
                    rs3.Close
                    rs13.Close
                    cmdCancel_Click
                    Exit Sub
                Else
                    txtHV.Text = rs3.Fields(1)
                    part_number = rs3.Fields(0)
                    part_revision = rs3.Fields(1)
                    order_number = rs3.Fields(3)
                End If
                rs3.Close
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
            stringSQL = "select A.bom_name from BOM as A where A.bom_name in ('" & order_number & "_" & part_number & "')  or A.bom_name in (" & _
            "'_DEL_" & order_number & "_" & part_number & "') "
            If rs5.State = 1 Then rs5.Close
            rs5.Open stringSQL
            If rs5.EOF = False Then
                flagHasBOM = True
                stringBOM = rs5.Fields(0)
            Else
                flagHasBOM = False
            End If
            
            If flagHasBOM = True Then
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
      sql = "select PrintSV from tblH3C where  SN='" & Mid(txtSN.Text, 3, 8) & "'  and HV = '" & txtHV.Text & "'"
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
                        
                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (rtrim(remark) <> '' and remark is not null AND testtime >= dateadd(month,-3,getdate())) ORDER BY testtime DESC "
'                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (ISNULL(remark, '') <> '') ORDER BY testtime DESC "
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
      
      
      sql = "select ID, HV, SN, CPN, EPN, Des, OS, GW, RoHS,[Combination], PictureID ,Power, MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark, EC from tblH3C_2D where SN='" & Mid(txtSN.Text, 3, 8) & "' and HV='" & txtHV.Text & "'"
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
        txtEC.Text = rec.Fields(18)
        
        Dim psv As String
        psv = rec.Fields(16)
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
        If UCase(Trim(rec.Fields(8))) = "N" Then
           Me.chkRoHS.Value = 0
           Me.chkNonRoHS.Value = 1
        ElseIf UCase(Trim(rec.Fields(8))) = "Y" Then
           Me.chkRoHS.Value = 0
           Me.chkNonRoHS.Value = 1
        End If
        
        If Trim(rec.Fields(8)) = "N" Then
            Me.lbPath.Caption = "\\sz-fs01\Public\Manufacture\标签模板\H3C发货标签认证组合\H3C发货标签认证组合图库-Rohs\" & rec.Fields(9) & ".bmp"
        Else
            Me.lbPath.Caption = "\\sz-fs01\Public\Manufacture\标签模板\H3C发货标签认证组合\H3C发货标签认证组合图库-Non-Rohs\" & rec.Fields(9) & ".bmp"
        End If
        

        txtMS.Text = rec.Fields(12)
        Me.lblMSday.Text = rec.Fields(13)
        txtNAL.Text = rec.Fields(14)
        Me.lblNALday.Text = rec.Fields(15)
'        dtpValidFrom.Value = rec.Fields(14)
'        dtpValidTo.Value = rec.Fields(15)
'        txtChangNAL.Text = rec.Fields(16)
'        txtHV.Text = rec.Fields(17)
        txtRemark.Text = rec.Fields(17)
        
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
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "H3C发货模版(含2D)变量.lab")
   Me.MousePointer = vbDefault
   myDoc.Printer.SwitchTo (lstPrinter.Text)
   
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

'Private Sub OpenLppx_hp()
'   Me.MousePointer = vbHourglass
'   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP发货标签.lab")
'
'   Me.MousePointer = vbDefault
'   Set myVars = myDoc.Variables
'   Set myObjs = myDoc.DocObjects
'End Sub

