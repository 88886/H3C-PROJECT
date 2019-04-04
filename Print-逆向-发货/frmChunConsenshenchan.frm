VERSION 5.00
Begin VB.Form frmChunConsenshenchan 
   Caption         =   "海康生产"
   ClientHeight    =   9225
   ClientLeft      =   5970
   ClientTop       =   1425
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   9210
   Begin VB.CheckBox chkN4 
      Caption         =   "N4"
      Height          =   495
      Left            =   4800
      TabIndex        =   31
      Top             =   8400
      Width           =   495
   End
   Begin VB.CheckBox chkN 
      Caption         =   "N*"
      Height          =   495
      Left            =   4080
      TabIndex        =   30
      Top             =   8400
      Width           =   495
   End
   Begin VB.TextBox txtVer 
      BackColor       =   &H00E0E0E0&
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
      Height          =   405
      Left            =   6720
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtSZ 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   450
      Left            =   6720
      TabIndex        =   27
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox txtHKSN 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6720
      TabIndex        =   25
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox txtHKPart 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4920
      Width           =   2895
   End
   Begin VB.PictureBox Picture15 
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   8400
      Width           =   615
   End
   Begin VB.CheckBox chkNonCCC 
      BackColor       =   &H0000C000&
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
      Left            =   7560
      TabIndex        =   20
      Top             =   8520
      Width           =   735
   End
   Begin VB.CheckBox chkCCC 
      Caption         =   "有"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   8520
      Width           =   735
   End
   Begin VB.TextBox txtOrder 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7800
      Width           =   2895
   End
   Begin VB.TextBox txtRev 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6720
      TabIndex        =   16
      Top             =   4200
      Width           =   2175
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
      Left            =   7440
      TabIndex        =   15
      Top             =   3360
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
      Left            =   5760
      TabIndex        =   14
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkY 
      Caption         =   "Y*"
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   8400
      Width           =   495
   End
   Begin VB.CheckBox chkY2 
      Caption         =   "Y2"
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   8400
      Width           =   615
   End
   Begin VB.TextBox txtSN 
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
      Left            =   2400
      TabIndex        =   10
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox txtPart 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox txtType 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   6360
      Width           =   2895
   End
   Begin VB.TextBox txtMaterial 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   7080
      Width           =   2895
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   9015
      Begin VB.Image Image3 
         Height          =   6045
         Left            =   -120
         Picture         =   "frmChunConsenshenchan.frx":0000
         Top             =   -720
         Width           =   14805
      End
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SZ:"
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
      Left            =   5760
      TabIndex        =   28
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000004&
      Caption         =   "海康SN："
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
      Left            =   5400
      TabIndex        =   26
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000004&
      Caption         =   "H3C版本:"
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
      Left            =   5400
      TabIndex        =   24
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000004&
      Caption         =   "海康机种："
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
      Left            =   600
      TabIndex        =   23
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "销售订单："
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
      Left            =   600
      TabIndex        =   18
      Top             =   7920
      Width           =   1575
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
      Left            =   600
      TabIndex        =   13
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9000
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "H3C SN："
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
      Left            =   840
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "H3C机种："
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
      Left            =   600
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "产品型号："
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
      Left            =   600
      TabIndex        =   6
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      Caption         =   "物料代码："
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
      Left            =   600
      TabIndex        =   5
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "产品类别："
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
      Left            =   600
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
   End
End
Attribute VB_Name = "frmChunConsenshenchan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim rs As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim mydoc As LabelManager2.Document
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
Dim newLableFlag As Boolean

Private Sub chkY_Click()
    If chkY.Value = 1 Then
        chkY2.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub chkY2_Click()
    If chkY2.Value = 1 Then
        chkY.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub chkN_Click()
    If chkN.Value = 1 Then
        chkY.Value = 0
        chkY2.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub chkN4_Click()
    If chkN4.Value = 1 Then
        chkY.Value = 0
        chkY2.Value = 0
        chkN.Value = 0
    End If
End Sub

Private Sub CommandCancel_Click()
    cmdCancel_HPSN_Click
End Sub

Private Sub CommandPrint_Click()

    Dim Pb As String

    
    If verifyPB() = False Then
        Exit Sub
    End If
    
   If (chkY2.Value = 1) Then
        Pb = CommonValue.PB_Y2
    ElseIf (chkY.Value = 1) Then
        Pb = CommonValue.PB_Y
    ElseIf (chkN.Value = 1) Then
        Pb = CommonValue.PB_N
    ElseIf (chkN4.Value = 1) Then
        Pb = CommonValue.PB_N4
    End If
                   
    If UploadH3CInfo2(False, Trim(Me.txtSN.Text), Trim(Me.txtVer.Text), "NA", "", "CHINA", golUSERNAME, Pb) = False Then
           MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
           txtSN.SetFocus
           Exit Sub
    End If
                    
'     If UploadHKShip_Rec(Trim(UCase(txtSN.Text)), Trim(UCase(txtPart.Text)), Trim(UCase(txtRev.Text)), "", Trim(UCase(txtHKSN.Text)), Trim(UCase(txtDesc.Text)), Trim(UCase(txtType.Text)), Trim(UCase(txtMaterial.Text)), Trim(UCase(txtOrder.Text)), IIf(chkY2.Value, "1", "0"), IIf(chkCCC.Value, "1", "0"), golUSERNAME) = False Then
'        MsgBox "海康资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
'        '        txtSN.SetFocus
'        UnloadLppx
'        Exit Sub
'    End If
    cmdPrint_HPSN_Click
    
    Sleep (2000)
    Call Connect.addPrintedLabel(Me.txtSN.Text, Me.Name)

'    If con13.State = 1 Then
'        con13.Close
'    End If
    cmdCancel_HPSN_Click
End Sub

Private Sub Form_Load()
    Me.Show
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    
   txtSN.SetFocus
   
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
    

    Set mydoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\H3C发货标签NEW\逆向H3CNEW\新H3C地址模板\" & "海康-逆向生产.lab")
    
    Me.MousePointer = vbDefault
    Set myVars = mydoc.Variables
    Set myObjs = mydoc.DocObjects
End Sub


Private Sub cmdPrint_HPSN_Click()

    If txtSN.Text = "" Then
        MsgBox ("H3C序列号未输入，不能打印！")
        txtSN.SetFocus
        Exit Sub
    End If
    
    If txtHKSN.Text = "" Then
        MsgBox ("海康序列号未带出，不能打印！")
        Exit Sub
    End If
    
    If txtDesc.Text = "" Then
        MsgBox ("产品类别未带出，不能打印！")
        Exit Sub
    End If
   
    If txtType.Text = "" Then
        MsgBox ("产品型号未带出，不能打印！")
        Exit Sub
    End If
    
    If Trim(txtMaterial.Text) = "" Then
         MsgBox ("物料代码未带出，不能打印！")
         Exit Sub
    End If
    
   If Trim(txtOrder.Text) = "" Then
'        MsgBox ("销售订单未带出，不能打印！")
'        Exit Sub
        SalesOrder = ""
   Else
        SalesOrder = Trim(txtOrder.Text)
   End If
   
   If chkCCC.Value = 0 And chkNonCCC.Value = 0 Then
        MsgBox ("CCC未带出，不能打印！")
        Exit Sub
   Else
       If chkCCC.Value = 1 Then
            CCC = "1"
        Else
            CCC = "0"
        End If
   End If
   If chkY2.Value = 0 And chkY.Value = 0 And chkN.Value = 0 And chkN4.Value = 0 Then
        MsgBox ("H3C序列号的Pb未带出，不能打印！")
        Exit Sub
   Else
        If chkY2.Value = 1 Then
            Pb = "Y2"
        ElseIf chkY.Value = 1 Then
            Pb = "Y*"
        ElseIf chkN4.Value = 1 Then
            Pb = "N4"
        ElseIf chkN.Value = 1 Then
            Pb = "N*"
        End If
   End If

   If Trim(txtRev.Text) = "" Then
        MsgBox ("H3C序列号的版本未带出，不能打印！")
        Exit Sub
   End If
'   If Trim(txtPower.Text) = "" Then
'      MsgBox "电源维护不能为空!!", vbExclamation + vbOKOnly, "电源维护空"
'      txtPower.SetFocus
'      Exit Sub
'   End If

    OpenLppx

         
    myVars.Item("Product ID").Value = UCase(Trim(txtType.Text))
    myVars.Item("Product Name1").Value = UCase(Trim(txtDesc.Text))
    myVars.Item("SN").Value = UCase(Trim(txtHKSN.Text))
    myVars.Item("WLDM").Value = UCase(Trim(txtMaterial.Text))
    If SalesOrder <> "" Then
        myVars.Item("XSDD").Value = UCase(Trim(txtOrder.Text))
    Else
        myObjs("MS Title(2)").Top = 10000
        myObjs("order").Top = 10000
    End If
    myVars.Item("SCPC").Value = UCase(Trim(txtSN.Text)) + " " + UCase(Trim(txtRev.Text)) + " " + Pb
    
    If CCC = 1 Then
        myObjs("addr").Top = 10000
    Else
        myObjs("addr3c").Top = 10000
    End If

   If txtSZ.Text <> "SZ" Then
        myObjs("SZ").Top = 10000
   End If
   
    mydoc.PrintLabel 1
    mydoc.FormFeed
    UnloadLppx
    
End Sub


Private Sub cmdCancel_HPSN_Click()
txtSN.Text = ""
txtHKSN.Text = ""
txtPart.Text = ""
txtRev.Text = ""
txtHKPart.Text = ""
txtDesc.Text = ""
txtType.Text = ""
txtMaterial.Text = ""
txtOrder.Text = ""
txtSN.SetFocus

chkY.Value = 0
chkY2.Value = 0
chkN.Value = 0
chkN4.Value = 0
chkCCC.Value = 0
chkNonCCC.Value = 0
'chkY.Enabled = True
'chkY2.Enabled = True

End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
    
        Dim strModel As String
        Dim strVer As String
        Dim strIII As String
        Dim str2 As String
        Dim strPartNumber As String, Status As String, uploadPowerCode As Boolean
        
        txtHKSN.Text = ""
        txtPart.Text = ""
        txtRev.Text = ""
        txtHKPart.Text = ""
        txtDesc.Text = ""
        txtType.Text = ""
        txtMaterial.Text = ""
        txtOrder.Text = ""
        txtSN.SetFocus
        
        chkY.Value = 0
        chkY2.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
        chkCCC.Value = 0
        chkNonCCC.Value = 0
        
        Me.txtSN.Text = Trim(Me.txtSN.Text)
        
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
'            str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtSN.Text) & "'"
'             str = " select top 1 t.part_number,t.part_revision,t.creation_time,t.order_number from (" & _
'            "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
'            "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
'            "where b.original_sn_S = '" & Trim(Me.txtSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
'            com.CommandText = str
'            rs13.Open Source:=com
'            If rs13.EOF = True Then
'                MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
'                cmdCancel_HPSN_Click
'                rs13.Close
'                Exit Sub
'            Else
''                strPartNumber = rs13.Fields(0)
''                strVer = rs13.Fields(1)
'                txtPart.Text = rs13.Fields(0)
'                txtRev.Text = rs13.Fields(1)
                
                txtPart.Text = "HWF" & Mid(Trim(Me.txtSN.Text), 3, 8)
                txtHKPart.Text = "HWF" & Mid(Trim(Me.txtSN.Text), 3, 8)

'            If IsValidECOVersion(strPartNumber, strVer) = False Then
'                cmdCancel_HPSN_Click
'                Exit Sub
'            End If
              
                
'                Dim con14 As ADODB.Connection
'                Dim rs14 As ADODB.Recordset
'                Dim com14 As ADODB.Command
'
'                Set con14 = New ADODB.Connection
'                Set rs14 = New ADODB.Recordset
'                strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
'
'                con14.Open ConnectionString:=strConn
'                Set com14 = New ADODB.Command
'                com14.ActiveConnection = con14
'
'                str = "select top 1 c.part_number,a.component_SN,b.serial_number,c.serial_number from [DC_Component_SN] a join unit b on a.unit_key=b.unit_key  join C_Consen_Print_Rec c on a.component_SN=c.serial_number " & _
'                    " WHERE a.Remark='HP'and c.EFFE_FLAG='1' and b.serial_number like '" & Trim(txtSN.Text) & "%'"
'                com14.CommandText = str
'
'                rs14.Open Source:=com14
'                If rs14.EOF = True Then
'                    MsgBox "没有Link海康SN", vbOKOnly + vbExclamation, "警告"
'                    cmdCancel_HPSN_Click
'                    rs14.Close
'                    Exit Sub
'                Else
'                    If IsNull(rs14.Fields("part_number")) Then
'                        MsgBox ("此序列号Link的海康SN机种为空，请确认!")
'                        cmdCancel_HPSN_Click
'                        rs14.Close
'                        Exit Sub
'                    Else
'                        txtHKPart.Text = rs14.Fields("part_number")
'                        txtHKSN.Text = rs14.Fields("component_SN")
'                    End If
'                End If
'                txtHKPart.Text = "HWF0231A1MW"
'                txtHKSN.Text = "CN63FD601"
                 If getConsenInformation = False Then
                    If con13.State = 1 Then
                        con13.Close
                    End If
                    
                    If con14.State = 1 Then
                        con14.Close
                    End If
                    Exit Sub
                End If
       '============add by carson start for TR5=============
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
      '============add by carson end  =============
      
          'add by carson 20160810 for Roy required  software
    '''''''''''''''''''''''''''''''''''''''''''''''
        Dim conSQL01_1 As ADODB.Connection
        Dim rsSQL01_1 As ADODB.Recordset
        Dim comSQL01_1 As ADODB.Command
        Set conSQL01_1 = New ADODB.Connection
        Set rsSQL01_1 = New ADODB.Recordset
        strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.31;Initial Catalog=dataT; User ID=sa; PWD=Itadmin1"
        
        lRemark = "N/A"
        conSQL01_1.Open ConnectionString:=strConn
        Set comSQL01_1 = New ADODB.Command
        comSQL01_1.ActiveConnection = conSQL01_1
        
        str = "select top 1 remark from [test_equ_ATE] " & _
        " where barcode = '" & Trim(Me.txtSN.Text) & "' and pass = N'通过' and remark<>'' order by id DESC "
        comSQL01_1.CommandText = str
        rsSQL01_1.Open Source:=comSQL01_1
        If rsSQL01_1.EOF = False Then
            lRemark = UCase(Trim(rsSQL01_1.Fields("remark")))
        End If
        rsSQL01_1.Close
        txtVer.Text = lRemark
    '''''''''''''''''''''''''''''''''''''''''''''''
                
'                sql = "select * from HP where h3c_bom_code='" & strModel & "' and hp_sn_iii='" & strIII & "'"
'                If conn.State = 0 Then
'                    conn.ConnectionString = Connect.getConnectionstring
'                    conn.Open
'                End If
'                rec.Open sql, conn, adOpenKeyset, adLockReadOnly
'                If rec.EOF = False Then
                    
                    
                    '====================
'                     If IsNull(rec.Fields("hp_desc1")) Then
'                        MsgBox ("此序列号未维护描述信息!")
'                        cmdCancel_HPSN_Click
'                        rec.Close
'                        Exit Sub
'                    Else
'                        txtDesc = rec.Fields("hp_desc1")
'                    End If
                    
'                    If Not IsNull(rec.Fields("hp_desc2")) Then
'                        txtDesc = txtDesc & " " & rec.Fields("hp_desc2")
'                    End If
'
'                    If IsNull(rec.Fields("hp_product")) Then
'                        MsgBox ("此序列号未维护产品编码!")
'                        cmdCancel_HPSN_Click
'                        rs.Close
'                        rec.Close
'                        Exit Sub
'                    Else
'                        txtProduct = rec.Fields("hp_product")
'                    End If
'                    If IsNull(rec.Fields("upload_power_code")) = True Or IsNull(rec.Fields("5000_status")) Then
'                        MsgBox ("此机种未维护是否上传电源代码和5000米状态!")
'                        cmdCancel_HPSN_Click
'                        rs.Close
'                        rec.Close
'                        Exit Sub
'                    End If
                    
'                    If rec.Fields("new_label") = "Y" Then
'                        newLableFlag = True
'                    Else
'                        newLableFlag = False
'                    End If
'
'                    uploadPowerCode = rec.Fields("upload_power_code")
'                    Status = rec.Fields("5000_status")
'
'
'                    Dim res2 As New ADODB.Recordset
'                    sql = "select * from singleunit where sn='" & strModel & "'"
'                    res2.Open sql, conn, adOpenKeyset, adLockReadOnly
'                    If res2.EOF = True Then
'                        MsgBox ("此序列号未维护产品型号!")
'                        cmdCancel_HPSN_Click
'                        res2.Close
'                        rec.Close
'                        Exit Sub
'                    Else
'                        str2 = res2.Fields("type")
'                    End If
'                    res2.Close
'
'                    If IsNull(rec.Fields("hp_pn")) Then
'                        txtPN = ""
'                    Else
'                        txtPN = rec.Fields("hp_pn")
'                    End If
'
'                    If IsNull(rec.Fields("hp_gtin_number")) Then
'                        txtUPC = ""
'                    Else
'                        txtUPC = rec.Fields("hp_gtin_number")
'                    End If
    '===============add by ben 2012-02-05 start===============
'                    If Connect.reprint = False Then
'    '                If reprint = False Then
'                        If Connect.isPrintedHKLabel(Me.txtSN.Text) Then
'                            MsgBox ("此序列号已打印！")
'                            cmdCancel_HPSN_Click
'                            If rec.State = 1 Then
'                                rec.Close
'                            End If
'                            Exit Sub
'                        End If
'                    End If
                    'rs14.Close
    '======================================================================
    
'                    Dim lh As New Label_History, Pb As String
'                    Dim sn As String
'                    sn = txtSN.Text
'                    If (lh.Init(sn)) Then
'                        If lh.Pb = "Y*" Then
'                            chkY.Value = 1
'                            chkY2.Value = 0
'                            chkN.Value = 0
'                            chkN4.Value = 0
'                            chkY.Enabled = False
'                            chkY2.Enabled = False
'                            chkN.Enabled = False
'                            chkN4.Enabled = False
'                            Pb = CommonValue.PB_Y
'                        ElseIf lh.Pb = "Y2" Then
'                            chkY.Value = 0
'                            chkY2.Value = 1
'                            chkN.Value = 0
'                            chkN4.Value = 0
'                            chkY.Enabled = False
'                            chkY2.Enabled = False
'                            chkN.Enabled = False
'                            chkN4.Enabled = False
'                            Pb = CommonValue.PB_Y2
'                        ElseIf lh.Pb = "N*" Then
'                            chkY.Value = 0
'                            chkY2.Value = 0
'                            chkN.Value = 1
'                            chkN4.Value = 0
'                            chkY.Enabled = False
'                            chkY2.Enabled = False
'                            chkN.Enabled = False
'                            chkN4.Enabled = False
'                            Pb = CommonValue.PB_N
'                        ElseIf lh.Pb = "N4" Then
'                            chkY.Value = 0
'                            chkY2.Value = 0
'                            chkN.Value = 0
'                            chkN4.Value = 1
'                            chkY.Enabled = False
'                            chkY2.Enabled = False
'                            chkN.Enabled = False
'                            chkN4.Enabled = False
'                            Pb = CommonValue.PB_N4
'                        End If
'                    Else
'                        chkY.Enabled = True
'                        chkY2.Enabled = True
'                        chkN.Enabled = True
'                        chkN4.Enabled = True
'                        TextModel.Text = strModel
'                        TextType.Text = str2
'                        TextVersion.Text = strVer
'                        TextPowerCode.Text = CStr(uploadPowerCode)
'                        TextStatus.Text = Status
'                        MsgBox "请确认环保属性值"
'                        Exit Sub
'                    End If
'

        End If
        
End Sub

Private Sub OpenLppx2(model As String)
    Me.MousePointer = vbHourglass

'    Set myDoc2 = myApp2.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\Phase3 HP认证发货标签\" & model & ".lab")
'
'    Me.MousePointer = vbDefault
'    Set myVars2 = myDoc2.Variables
'    Set myObjs2 = myDoc2.DocObjects
    Set mydoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\Phase3 HP认证新发货标签\" & model & ".lab")
    
    Me.MousePointer = vbDefault
    Set myVars = mydoc.Variables
    Set myObjs = mydoc.DocObjects
End Sub

Private Function verifyPB() As Boolean
    If (chkY2.Value + chkY.Value + chkN.Value + chkN4.Value = 1) = False Then
        MsgBox "请确认环保属性"
        verifyPB = False
        Exit Function
    End If
    verifyPB = True
End Function

Private Sub cmdPrint_Model_Click(strModel As String, strXingHao As String, strVer As String)
    
    Dim Pb As String
    If (chkY2.Value = 1) Then
        Pb = CommonValue.PB_Y2
    ElseIf (chkY.Value = 1) Then
        Pb = CommonValue.PB_Y
    ElseIf (chkN.Value = 1) Then
        Pb = CommonValue.PB_N
    ElseIf (chkN4.Value = 1) Then
        Pb = CommonValue.PB_N4
    End If

    OpenLppx2 strModel

    myVars.Item("Model").Value = strXingHao
    myVars.Item("PN").Value = UCase(strModel)
    myVars.Item("Rev").Value = UCase(strVer)
    
    myVars.Item("Y2").Value = UCase(Pb)
   
    mydoc.PrintLabel 1
    mydoc.FormFeed
    UnloadLppx2
    
End Sub

Private Sub UnloadLppx2()
    myApp.Documents.CloseAll False
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Function getConsenInformation() As Boolean

'      sql = "select * from hp where hp_sn_iii=substring('" & Trim(txtSN.Text) & "',5,3) and h3c_bom_code = '" + txtModel.Text + "'"
      sql = "select * from tblConsenNew where Part_Number= '" + txtHKPart.Text + "'"
      If rec.State = 1 Then
        rec.Close
      End If
      
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      
      If rec.EOF = True Then
          MsgBox "此序列号未维护信息!"
          txtSN.Text = ""
'          txtSN.SetFocus
          rec.Close
          getConsenInformation = False
          Exit Function
      Else
          If IsNull(rec.Fields("Part_Number")) Then
              MsgBox ("此序列号未维护产品编码!")
              rec.Close
              getConsenInformation = False
              Exit Function
          Else
              txtHKPart.Text = rec.Fields("Part_Number")
          End If
'          hpsnproduct
    
          If IsNull(rec.Fields("Part_Type")) Then
              MsgBox ("此序列号未维护产品类别!")
              rec.Close
              getConsenInformation = False
              Exit Function
          Else
              txtDesc.Text = rec.Fields("Part_Type")
          End If
          
          If IsNull(rec.Fields("Part_Model")) Then
              MsgBox ("此序列号未维护产品型号!")
              rec.Close
              getConsenInformation = False
              Exit Function
          Else
              txtType.Text = rec.Fields("Part_Model")
          End If
    
          If IsNull(rec.Fields("Material")) Then
              MsgBox ("此序列号未维护物料代码")
              rec.Close
              getConsenInformation = False
              Exit Function
          Else
              txtMaterial.Text = rec.Fields("Material")
          End If
          
          If IsNull(rec.Fields("Power")) Then
'              MsgBox ("此序列号未维护电源!")
'              rec.Close
'              Exit Sub
'                txtPower.Text = ""
          Else
'              txtPower.Text = rec.Fields("Power")
          End If
          
          If IsNull(rec.Fields("SalesOrder")) Then
'              MsgBox ("此序列号未维护电源!")
'              rec.Close
'              Exit Sub
                txtOrder.Text = ""
          Else
              txtOrder.Text = rec.Fields("SalesOrder")
          End If
'
'        If UCase(Trim(rec.Fields("ChinaRoHS"))) = True Then
'           chkChinaRoHS.Value = 1
'           chkNonChinaRoHS.Value = 0
'        Else
'           chkChinaRoHS.Value = 0
'           chkNonChinaRoHS.Value = 1
'        End If
'
'        If UCase(Trim(rec.Fields("WEEE"))) = True Then
'           chkWEEE.Value = 1
'           chkNonWEEE.Value = 0
'        Else
'           chkWEEE.Value = 0
'           chkNonWEEE.Value = 1
'        End If
        
        If UCase(Trim(rec.Fields("CCC"))) = True Then
           chkCCC.Value = 1
           chkNonCCC.Value = 0
        Else
           chkCCC.Value = 0
           chkNonCCC.Value = 1
        End If
'        If UCase(Trim(rec.Fields("Laser"))) = True Then
'           chkLaser.Value = 1
'           chkNonLaser.Value = 0
'        Else
'           chkLaser.Value = 0
'           chkNonLaser.Value = 1
'        End If
'        If GetPbProperty(txtSN.Text) = "Y2" Then
'            chkY2.Value = 1
'            chkY.Value = 0
'            chkN.Value = 0
'            chkN4.Value = 0
'        ElseIf GetPbProperty(txtSN.Text) = "Y*" Then
'            chkY2.Value = 0
'            chkY.Value = 1
'            chkN.Value = 0
'            chkN4.Value = 0
'        ElseIf GetPbProperty(txtSN.Text) = "N*" Then
'            chkY2.Value = 0
'            chkY.Value = 0
'            chkN.Value = 1
'            chkN4.Value = 0
'        ElseIf GetPbProperty(txtSN.Text) = "N4" Then
'            chkY2.Value = 0
'            chkY.Value = 0
'            chkN.Value = 0
'            chkN4.Value = 1
'        Else
'            chkY2.Value = 0
'            chkY.Value = 0
'            chkN.Value = 0
'            chkN4.Value = 0
'            MsgBox ("此序列号Pb属性未设置")
'            rec.Close
'            getConsenInformation = False
'            Exit Function
'        End If
        
      End If
      getConsenInformation = True
End Function
