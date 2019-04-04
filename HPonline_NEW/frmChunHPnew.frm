VERSION 5.00
Begin VB.Form frmChunHPnew 
   Caption         =   "纯HP发货在线打印"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9210
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox TextStatus 
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TextPowerCode 
      Height          =   375
      Left            =   6000
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TextVersion 
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TextType 
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TextModel 
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CommandCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton CommandPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox chkY 
      Caption         =   "Y*"
      Height          =   495
      Left            =   3840
      TabIndex        =   12
      Top             =   6120
      Width           =   975
   End
   Begin VB.CheckBox chkY2 
      Caption         =   "Y2"
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txtHPSN 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2760
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4080
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4800
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   5520
      Width           =   2895
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9015
      Begin VB.Image Image1 
         Height          =   1860
         Left            =   4560
         Picture         =   "frmChunHPnew.frx":0000
         Top             =   240
         Width           =   4320
      End
      Begin VB.Image Image3 
         Height          =   1815
         Left            =   120
         Picture         =   "frmChunHPnew.frx":27D5
         Top             =   240
         Width           =   4305
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "环保属性："
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
      Left            =   720
      TabIndex        =   13
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "HP SN："
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
      Left            =   1200
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
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
      Left            =   720
      TabIndex        =   7
      Top             =   4200
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
      Left            =   720
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
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
      Left            =   720
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
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
      Left            =   720
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmChunHPnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim rs As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim hpsn As String
'Dim myApp2 As New LabelManager2.Application
'Dim myDoc2 As LabelManager2.Document
'Dim myVars2 As LabelManager2.Variables
'Dim myObjs2 As LabelManager2.DocObjects
Dim strDir As String
Dim rec As New ADODB.Recordset
Dim res2 As New ADODB.Recordset
Dim rec13 As New ADODB.Recordset

Private Sub chkY_Click()
    If chkY.Value = 1 Then
        chkY2.Value = 0
    End If
End Sub

Private Sub chkY2_Click()
    If chkY2.Value = 1 Then
        chkY.Value = 0
    End If
End Sub

Private Sub CommandCancel_Click()
    cmdCancel_HPSN_Click
End Sub

Private Sub CommandPrint_Click()

    Dim strModel As String, strType As String, strVersion As String, PB As String, uploadPowerCode As Boolean
    Dim status As String
    strModel = Trim(TextModel.Text)
    strType = Trim(TextType.Text)
    strVersion = Trim(TextVersion.Text)
    uploadPowerCode = CBool(TextPowerCode.Text)
    status = Trim(TextStatus.Text)
    
    If verifyPB() = False Then
        Exit Sub
    End If
    
    If (chkY2.Value = 1) Then
        PB = CommonValue.PB_Y2
    ElseIf (chkY.Value = 1) Then
        PB = CommonValue.PB_Y
    End If


    If UploadH3C_PB(PB, Trim(UCase(txtHPSN.Text)), strVersion, "NA", "N/A", "CHINA", "frmChunHP") = False Then
        MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        Exit Sub
    End If

    If UploadH3CInfo2(uploadPowerCode, Trim(Me.txtHPSN.Text), "N/A", status, "", "CHINA", golUSERNAME, PB) = False Then
         MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
         txtSN.SetFocus
         Exit Sub
    End If
        cmdPrint_HPSN_Click
        
        Sleep 2000

        cmdPrint_Model_Click strModel, strType, strVersion
        Call Connect.addPrintedLabel(Me.txtHPSN.Text, Me.Name)

        cmdCancel_HPSN_Click
End Sub

Private Sub Form_Load()
    Me.Show
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    
   txtHPSN.SetFocus
   
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
    myApp.EnableEvents = True
    Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "新HP发货标签NEW.lab")
    
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub


Private Sub cmdPrint_HPSN_Click()

    If txtHPSN.Text = "" Then
        MsgBox ("序列号未输入，不能打印！")
        txtSN.SetFocus
        Exit Sub
    End If

    If txtProduct.Text = "" Then
        MsgBox ("产品编码未带出，不能打印！")
        Exit Sub
    End If
    If txtDesc.Text = "" Then
        MsgBox ("产品描述未带出，不能打印！")
        Exit Sub
    End If

    OpenLppx

         
    myVars.Item("ID").Value = txtDesc.Text
    myVars.Item("SN2").Value = UCase(txtHPSN.Text)
   
    If Trim(txtPN.Text) <> "" Then
        myVars.Item("PN2").Value = UCase(txtPN.Text)
    Else
        myObjs("bcPN").Top = 10000
    End If

    myVars.Item("Product2").Value = UCase(txtProduct.Text)
    myVars.Item("UPC").Value = Left(Trim(txtUPC.Text), 11)
    

    myDoc.PrintLabel 1
    myDoc.FormFeed
    UnloadLppx
    
End Sub


Private Sub cmdCancel_HPSN_Click()
txtHPSN.Text = ""
txtProduct.Text = ""
txtDesc.Text = ""
txtUPC.Text = ""
txtPN.Text = ""
txtHPSN.SetFocus

chkY.Value = 0
chkY2.Value = 0
chkY.Enabled = True
chkY.Enabled = True

End Sub


Private Sub txtHPSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
    
        Dim strModel As String
        Dim strVer As String
        Dim strIII As String
        Dim str2 As String
        Dim strPartNumber As String, status As String, uploadPowerCode As Boolean
        
        Me.txtHPSN.Text = Trim(Me.txtHPSN.Text)
        
        '==================
        Dim con As ADODB.Connection
        Dim rs3 As ADODB.Recordset
        Dim rs4 As ADODB.Recordset
        

        Set con = New ADODB.Connection

        Set rs4 = New ADODB.Recordset
        
        con.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
        con.ConnectionTimeout = 50
        con.Open
        Dim str As String

                
        Set rs4.ActiveConnection = con
        rs4.CursorType = adOpenForwardOnly
        

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
'            str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtHPSN.Text) & "'"
             str = " select top 1 t.part_number,t.part_revision,t.creation_time,t.order_number from (" & _
            "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtHPSN.Text) & "' union " & _
            "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
            "where b.original_sn_S = '" & Trim(Me.txtHPSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
            com.CommandText = str
            rs13.Open Source:=com
            If rs13.EOF = True Then
                MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
                cmdCancel_HPSN_Click
                rs13.Close
                Exit Sub
            Else
                strPartNumber = rs13.Fields(0)
                strModel = Mid(Trim(rs13.Fields(0)), 4, 8)
                strVer = rs13.Fields(1)
                strIII = Mid(Trim(txtHPSN.Text), 5, 3)
                Select Case strModel
                Case "0231A1DS", "0231A1DT", "0231A1DU", "0231A1DV", "0231A1WQ", "0231A1WR ", "0231A1WS"
                    MsgBox "当前机种应对应'HP双SN发货程序'打印", vbOKOnly + vbExclamation, "打印程序使用错误"
                    cmdCancel_HPSN_Click
                    rs13.Close
                    Exit Sub
                Case Else
                    
                End Select


            If IsValidECOVersion(strPartNumber, strVer) = False Then
                cmdCancel_HPSN_Click
                Exit Sub
            End If
              
                
            Dim con14 As ADODB.Connection
            Dim rs14 As ADODB.Recordset
            Dim com14 As ADODB.Command

            Set con14 = New ADODB.Connection
            Set rs14 = New ADODB.Recordset
            strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
            'con13.ConnectionTimeout = 50
            con14.Open ConnectionString:=strConn
            Set com14 = New ADODB.Command
            com14.ActiveConnection = con14
            'str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtHPSN.Text) & "'"
              str = "select 1 from [H3C_HPWeight] " & _
                " where ((Part_Number = '" & strPartNumber & "' and part_revision = '" & strVer & "') or " & _
                " (Part_Number = '" & strPartNumber & "' and part_revision = 'ALL')) " & _
                " and GrossWeight is not null and NetWeight is not null " & _
                " and GrossWeight <> '' and NetWeight <> '' " & _
                " and is_Valid = 1 "
            com14.CommandText = str
               
               'rs14.Open str
                rs14.Open Source:=com14
                If rs14.EOF = True Then
                    MsgBox "没有维护重量", vbOKOnly + vbExclamation, "警告"
                    cmdCancel_HPSN_Click
                    rs14.Close
                    Exit Sub
                End If

                Set fs = CreateObject("Scripting.FileSystemObject")
                'Dim fs As New FileSystemObject
    
                strDir = "\\sz-fs01\Public\Manufacture\标签模板\Phase3 HP认证发货标签\" & strModel & ".lab"
                If Not fs.FileExists(strDir) Then
                    MsgBox "没有对应机种打印模板", vbOKOnly + vbExclamation, "警告"
                    cmdCancel_HPSN_Click
                    rs3.Close
                    Exit Sub
                End If
                       

                sql = "select * from HP where h3c_bom_code='" & strModel & "' and hp_sn_iii='" & strIII & "'"
                If conn.State = 0 Then
                    conn.ConnectionString = Connect.getConnectionstring
                    conn.Open
                End If
                rec.Open sql, conn, adOpenKeyset, adLockReadOnly
                If rec.EOF = False Then
                    
                    
                    '====================
                     If IsNull(rec.Fields("hp_desc1")) Then
                        MsgBox ("此序列号未维护描述信息!")
                        cmdCancel_HPSN_Click
                        rec.Close
                        Exit Sub
                    Else
                        txtDesc = rec.Fields("hp_desc1")
                    End If
                    
                    If Not IsNull(rec.Fields("hp_desc2")) Then
                        txtDesc = txtDesc & " " & rec.Fields("hp_desc2")
                    End If
                
                    If IsNull(rec.Fields("hp_product")) Then
                        MsgBox ("此序列号未维护产品编码!")
                        cmdCancel_HPSN_Click
                        rs.Close
                        rec.Close
                        Exit Sub
                    Else
                        txtProduct = rec.Fields("hp_product")
                    End If
                    If IsNull(rec.Fields("upload_power_code")) = True Or IsNull(rec.Fields("5000_status")) Then
                        MsgBox ("此机种未维护是否上传电源代码和5000米状态!")
                        cmdCancel_HPSN_Click
                        rs.Close
                        rec.Close
                        Exit Sub
                    End If
                    
                    uploadPowerCode = rec.Fields("upload_power_code")
                    status = rec.Fields("5000_status")
                    
                    
                    Dim res2 As New ADODB.Recordset
                    sql = "select * from singleunit where sn='" & strModel & "'"
                    res2.Open sql, conn, adOpenKeyset, adLockReadOnly
                    If res2.EOF = True Then
                        MsgBox ("此序列号未维护产品型号!")
                        cmdCancel_HPSN_Click
                        res2.Close
                        rec.Close
                        Exit Sub
                    Else
                        str2 = res2.Fields("type")
                    End If
                    res2.Close
                    
                    If IsNull(rec.Fields("hp_pn")) Then
                        txtPN = ""
                    Else
                        txtPN = rec.Fields("hp_pn")
                    End If
                    
                    If IsNull(rec.Fields("hp_gtin_number")) Then
                        txtUPC = ""
                    Else
                        txtUPC = rec.Fields("hp_gtin_number")
                    End If
    '===============add by ben 2012-02-05 start===============
                    If Connect.reprint = False Then
    '                If reprint = False Then
                        If Connect.isPrintedLabel(Me.txtHPSN.Text, Me.Name) Then
                            MsgBox ("此序列号已打印！")
                            cmdCancel_HPSN_Click
                            If rec.State = 1 Then
                                rec.Close
                            End If
                            Exit Sub
                        End If
                    End If
                    rec.Close
    '======================================================================
    
    Dim lh As New Label_History, PB As String
    Dim sn As String
    sn = txtHPSN.Text
    If (lh.Init(sn)) Then
        If lh.PB = "Y*" Then
            chkY.Value = 1
            chkY2.Value = 0
            chkY.Enabled = False
            chkY2.Enabled = False
            PB = CommonValue.PB_Y
        End If
        If lh.PB = "Y2" Then
            chkY.Value = 0
            chkY2.Value = 1
            chkY.Enabled = False
            chkY2.Enabled = False
            PB = CommonValue.PB_Y2
        End If
    Else
        chkY.Enabled = True
        chkY2.Enabled = True
        TextModel.Text = strModel
        TextType.Text = str2
        TextVersion.Text = strVer
        TextPowerCode.Text = CStr(uploadPowerCode)
        TextStatus.Text = status
        MsgBox "请确认环保属性值"
        Exit Sub
    End If
    

    
        If UploadH3CInfo2(uploadPowerCode, Trim(Me.txtHPSN.Text), "N/A", status, "", "CHINA", golUSERNAME, PB) = False Then
             MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
             txtSN.SetFocus
             Exit Sub
        End If
    '===============add by ben 2012-02-05 end=================
                    cmdPrint_HPSN_Click
                    
                    Sleep (2000)
    
                    cmdPrint_Model_Click strModel, str2, strVer
    '===============add by ben 2012-02-05 start===============
                    Call Connect.addPrintedLabel(Me.txtHPSN.Text, Me.Name)
                End If
            End If
            If rs13.State = 1 Then
                rs13.Close
            End If
            If con13.State = 1 Then
                con13.Close
            End If
            cmdCancel_HPSN_Click
        End If
        
End Sub

Private Sub OpenLppx2(model As String)
    Me.MousePointer = vbHourglass

'    Set myDoc2 = myApp2.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\Phase3 HP认证发货标签\" & model & ".lab")
'
'    Me.MousePointer = vbDefault
'    Set myVars2 = myDoc2.Variables
'    Set myObjs2 = myDoc2.DocObjects
    Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\Phase3 HP认证新发货标签\" & model & ".lab")
    
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub

Private Function verifyPB() As Boolean
    If (chkY2.Value + chkY.Value = 1) = False Then
        MsgBox "请确认环保属性"
        verifyPB = False
        Exit Function
    End If
    verifyPB = True
End Function

Private Sub cmdPrint_Model_Click(strModel As String, strXingHao As String, strVer As String)
    
    Dim PB As String
    If (chkY2.Value = 1) Then
        PB = CommonValue.PB_Y2
    ElseIf (chkY.Value = 1) Then
        PB = CommonValue.PB_Y
    End If

    OpenLppx2 strModel

    myVars.Item("Model").Value = strXingHao
    myVars.Item("PN").Value = UCase(strModel)
    myVars.Item("Rev").Value = UCase(strVer)
    
    myVars.Item("Y2").Value = UCase(PB)
   
    myDoc.PrintLabel 1
    myDoc.FormFeed
    UnloadLppx2
    
End Sub

Private Sub UnloadLppx2()
    myApp.Documents.CloseAll False
    myApp.Quit
    Set myApp = Nothing
End Sub

