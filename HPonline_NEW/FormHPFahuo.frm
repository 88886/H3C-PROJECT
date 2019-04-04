VERSION 5.00
Begin VB.Form FormHPFahuo 
   BackColor       =   &H80000004&
   Caption         =   "HP发货标签"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   7155
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtHPSN 
      Height          =   495
      Left            =   6480
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtMac 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6360
      Width           =   2895
   End
   Begin VB.TextBox txtImei1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6960
      Width           =   2895
   End
   Begin VB.TextBox txtImei2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7560
      Width           =   2895
   End
   Begin VB.TextBox txtModel_hid 
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Top             =   6600
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   6945
      TabIndex        =   14
      Top             =   120
      Width           =   6975
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   0
         Picture         =   "FormHPFahuo.frx":0000
         ScaleHeight     =   2985
         ScaleWidth      =   6945
         TabIndex        =   15
         Top             =   0
         Width           =   6975
      End
   End
   Begin VB.CommandButton cmdMPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "批量打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "打 印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "取 消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdReturn_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "返 回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   10
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtUPC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5640
      Width           =   2895
   End
   Begin VB.TextBox txtProduct 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox txtPN 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox txtSN 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "MAC："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000004&
      Caption         =   "IMEI："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   21
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000004&
      Caption         =   "IMEI："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "产品描述："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      Caption         =   "产品UPC："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "产品编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "产品机种："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "产品序列号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
End
Attribute VB_Name = "FormHPFahuo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim str As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects

Public Sub cmdMPrint_Click()

    'begin
            If Len(txtSN.Text) < 10 Then
            MsgBox "产品序号长度不能小于10!"
            txtSN.SetFocus
            Exit Sub
        End If
        sql = "select * from hp where charindex(hp_sn_iii,'" & Trim(txtSN.Text) & "')=5 and pack_label='Y' and h3c_bom_code='" & txtModel_hid.Text & "'"
        'MsgBox (sql)
        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
        If rec.EOF = True Then
            MsgBox "此序列号未维护信息!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
        Else
        
           If IsNull(rec.Fields("hp_pn")) Then                     'modified by Jimmy Sun 2010.06.17
           txtPN = ""
            '    MsgBox ("此序列号未维护机种!")
             '   rec.Close
              '  Exit Sub
            Else
                txtPN = rec.Fields("hp_pn")
            End If
             
             
            If IsNull(rec.Fields("hp_gtin_number")) Then           'modified by David Xu 2011.04.20
                'MsgBox ("此序列号未维护UPC!")
                'rec.Close
                'Exit Sub
                txtUPC = ""
            Else
                txtUPC = rec.Fields("hp_gtin_number")
            End If
        
            If IsNull(rec.Fields("hp_product")) Then
                MsgBox ("此序列号未维护产品编码!")
                rec.Close
                Exit Sub
            Else
                txtProduct = rec.Fields("hp_product")
            End If
      
            If IsNull(rec.Fields("hp_desc1")) Then
                MsgBox ("此序列号未维护描述信息!")
                rec.Close
                Exit Sub
            Else
                txtDesc = rec.Fields("hp_desc1")
            End If
            
            If Not IsNull(rec.Fields("hp_desc2")) Then
                txtDesc = txtDesc & " " & rec.Fields("hp_desc2")
            End If
            rec.Close
            
'=========================== add by carson 20161114 for guodong ma required
            'get MAC address from test records
            txtMac.Text = ""
            lMAC = ""
            Dim conSQL01 As ADODB.Connection
            Dim rsSQL01 As ADODB.Recordset
            Dim comSQL01 As ADODB.Command
            Set conSQL01 = New ADODB.Connection
            Set rsSQL01 = New ADODB.Recordset
            strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.31;Initial Catalog=dataT; User ID=sa; PWD=Itadmin1"
            
            Dim conSQL01_1 As ADODB.Connection
            Dim rsSQL01_1 As ADODB.Recordset
            Dim comSQL01_1 As ADODB.Command
            Set conSQL01_1 = New ADODB.Connection
            Set rsSQL01_1 = New ADODB.Recordset

            MAC1 = GetMacFromTestRecord(Me.txtHPSN.Text, "MTP-Flash")
            MAC2 = GetMacFromTestRecord(Me.txtHPSN.Text, "HANDWORK_TAAERASE")
            MAC3 = GetMacFromTestRecord(Me.txtHPSN.Text, "HANDWORK_JTAG")
            MAC4 = GetMacFromTestRecord(Me.txtHPSN.Text, "")
            If MAC1 <> "" Then 'equipment LIKE 'MTP-Flash%'  exist
                If MAC1 <> "1" Then
                    lMAC = MAC1
                Else
                    lMAC = ""
                End If
            ElseIf MAC2 <> "" Then 'equipment LIKE 'HANDWORK_TAAERASE%' NOT exist
                If MAC2 <> "1" Then
                    lMAC = MAC2
                Else
                    lMAC = ""
                End If
            ElseIf MAC3 <> "1" Then 'equipment LIKE 'HANDWORK_JTAG%' NOT exist
                If MAC3 <> "1" Then
                    lMAC = MAC3
                Else
                    lMAC = ""
                End If
            ElseIf MAC4 <> "" Then 'equipment ="" NOT exist
                lMAC = ""
            End If
            txtMac.Text = lMAC
                    
            'get IMEI  from FTPC
            txtImei1.Text = ""
            txtImei2.Text = ""
            lIMEI1 = ""
            lIMEI2 = ""
            Dim conFTPC1 As ADODB.Connection
            Dim rsFTPC1 As ADODB.Recordset
            Dim comFTPC1 As ADODB.Command
            Set conFTPC1 = New ADODB.Connection
            Set rsFTPC1 = New ADODB.Recordset
            strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
            'con13.ConnectionTimeout = 50
            conFTPC1.Open ConnectionString:=strConn
            Set comFTPC1 = New ADODB.Command
            comFTPC1.ActiveConnection = conFTPC1
            str = " select a.serial_number,b.part_serial,d.component_SN,d.seqno from UNIT a join CONSUMED_PART b  on a.unit_key=b.tobj_key " & _
                " join UNIT c on c.serial_number=b.part_serial" & _
                " join DC_Component_SN d on d.unit_key=c.unit_key " & _
                " where b.status='Consumed' and d.Remark='GPSSupply' " & _
                " and a.serial_number= '" & Trim(Me.txtHPSN.Text) & "' "
            comFTPC1.CommandText = str
            rsFTPC1.Open Source:=comFTPC1
            While rsFTPC1.EOF = False
                If lIMEI1 = "" Then
                    lIMEI1 = rsFTPC1.Fields("component_SN")
                    txtImei1.Text = lIMEI1
                Else
                    lIMEI2 = rsFTPC1.Fields("component_SN")
                    txtImei2.Text = lIMEI2
                    
                End If
                rsFTPC1.MoveNext
            Wend
            rsFTPC1.Close
        '========================================
            
            cmdPrint_HPSN_Click
        End If
                
        'FormHPFahuo.Hide
        frmNewH3CPrint.Show
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
Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub
Private Sub OpenLppx()
    Me.MousePointer = vbHourglass
    'Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HUAWEI-生产.lab")
    'If Trim(txtPN.Text) <> "" Then
    '    Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP发货标签.lab")
    'Else
    '    Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP发货标签_NO_PN.lab")
    'End If
    Set myDoc = myApp.Documents.Open("\\10.11.1.25\Public\Manufacture\标签模板\" & "HP发货标签NEW.lab")
    
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub

Private Sub cmdCancel_HPSN_Click()
txtSN.Text = ""
txtProduct.Text = ""
txtDesc.Text = ""
txtUPC.Text = ""
txtPN.Text = ""
'txtSZ.Text = ""
txtMac.Text = ""
txtImei1.Text = ""
txtImei2.Text = ""
txtHPSN.Text = ""

txtSN.SetFocus

End Sub

Private Sub cmdPrint_HPSN_Click()

    If txtSN.Text = "" Then
        MsgBox ("序列号未输入，不能打印！")
        txtSN.SetFocus
        Exit Sub
    End If
'    If txtPN.Text = "" Then
'        MsgBox ("机种未带出，不能打印！")
'        txtSN.SetFocus
'        Exit Sub
'    End If
    If txtProduct.Text = "" Then
        MsgBox ("产品编码未带出，不能打印！")
        Exit Sub
    End If
    If txtDesc.Text = "" Then
        MsgBox ("产品描述未带出，不能打印！")
        Exit Sub
    End If
    'If txtUPC.Text = "" Then
    '    MsgBox ("产品UPC未带出，不能打印！")
    '    Exit Sub
    'End If
    OpenLppx

    'myVars.Item("SN1").Value = UCase(txtSN.Text)
    'myVars.Item("SN2").Value = "S" & UCase(txtSN.Text)
    'myVars.Item("PN1").Value = UCase(txtPN.Text)
    'myVars.Item("Product1").Value = UCase(txtProduct.Text)
         
    myVars.Item("ID").Value = txtDesc.Text
    myVars.Item("SN2").Value = UCase(txtSN.Text)
   
    If Trim(txtPN.Text) <> "" Then
        myVars.Item("PN2").Value = UCase(txtPN.Text)
    Else
        myObjs("bcPN").Top = 10000
        myObjs("Text1(16)").Top = 10000
        
    End If

    myVars.Item("Product2").Value = UCase(txtProduct.Text)
    
    If Trim(txtUPC.Text) <> "" Then
        myVars.Item("UPC").Value = Left(Trim(txtUPC.Text), 11)
    Else
        myObjs("Barcode26(6)").Top = 10000
        myObjs("Text1(21)").Top = 10000
    End If
        
    
    If Trim(txtMac.Text) <> "" Then
        myVars.Item("MAC1").Value = UCase(Trim(txtMac.Text))
    Else
        myObjs("MAC(1)").Top = 10000
        myObjs("MAC").Top = 10000
        myObjs("MAC1").Top = 10000
    End If

    If Trim(txtImei1.Text) <> "" Then
        myVars.Item("IMEI1").Value = UCase(Trim(txtImei1.Text))
    Else
        myObjs("IMEI1").Top = 10000
        myObjs("Barcode6").Top = 10000
        myObjs("IMEI2").Top = 10000
        myObjs("Barcode7").Top = 10000
    End If

    If Trim(txtImei2.Text) <> "" Then
        myVars.Item("IMEI2").Value = UCase(Trim(txtImei2.Text))
    Else
        myObjs("IMEI2").Top = 10000
        myObjs("Barcode7").Top = 10000
    End If
    
    myDoc.PrintLabel 1
    myDoc.FormFeed
    UnloadLppx
    cmdCancel_HPSN_Click
    
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If Len(txtSN.Text) < 10 Then
            MsgBox "产品序号长度不能小于10!"
            txtSN.SetFocus
            Exit Sub
        End If
        sql = "select * from hp where charindex(hp_sn_iii,'" & Trim(txtSN.Text) & "')<>0 and h3c_bom_code='" & txtModel_hid.Text & "'"
        'MsgBox (sql)
        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
        If rec.EOF = True Then
            MsgBox "此序列号未维护信息!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
        Else
        
            If IsNull(rec.Fields("hp_pn")) Then
                MsgBox ("此序列号未维护机种!")
                rec.Close
                Exit Sub
            Else
                txtPN = rec.Fields("hp_pn")
            End If
             
            If IsNull(rec.Fields("hp_gtin_number")) Then
                'MsgBox ("此序列号未维护UPC!")
                'rec.Close
                'Exit Sub
                txtUPC = ""
            Else
                txtUPC = rec.Fields("hp_gtin_number")
            End If
        
            If IsNull(rec.Fields("hp_product")) Then
                MsgBox ("此序列号未维护产品编码!")
                rec.Close
                Exit Sub
            Else
                txtProduct = rec.Fields("hp_product")
            End If
      
            If IsNull(rec.Fields("hp_desc1")) Then
                MsgBox ("此序列号未维护描述信息!")
                rec.Close
                Exit Sub
            Else
                txtDesc = rec.Fields("hp_desc1")
            End If
            
            If Not IsNull(rec.Fields("hp_desc2")) Then
                txtDesc = txtDesc & " " & rec.Fields("hp_desc2")
            End If
        
        End If
        
    End If
End Sub

Private Sub cmdReturn_HPSN_Click()
Unload Me
End Sub
