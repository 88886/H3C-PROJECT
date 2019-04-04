VERSION 5.00
Begin VB.Form frmHW3COM 
   Caption         =   "3COM标签打印"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8820
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox TextFTState 
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TextHPSN 
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox CheckHpPackLabel 
      Caption         =   "Y2"
      Height          =   495
      Left            =   5640
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TextPowerCode 
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TextVersion 
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TextModel 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkY 
      Caption         =   "Y*"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   3960
      Width           =   975
   End
   Begin VB.CheckBox chkY2 
      Caption         =   "Y2"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   6600
      Width           =   1815
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
      Left            =   3000
      TabIndex        =   1
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.Image Image3 
         Height          =   2640
         Left            =   120
         Picture         =   "frmHW3COM.frx":0000
         Top             =   120
         Width           =   8475
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
      Left            =   1200
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "SN："
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
      Left            =   2040
      TabIndex        =   2
      Top             =   3240
      Width           =   735
   End
End
Attribute VB_Name = "frmHW3COM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim rec As New ADODB.Recordset
Dim myApp2 As New LabelManager2.Application
Dim myDoc2 As LabelManager2.Document
Dim myVars2 As LabelManager2.Variables
Dim myObjs2 As LabelManager2.DocObjects
Public unit_key As Long
Public HP_pack_label As Boolean

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
        chkY2.Value = 0
        chkY.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub chkN4_Click()
    If chkN4.Value = 1 Then
        chkY.Value = 0
        chkY2.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub cmdPrint_Click()

    If verifyPB() = False Then
        Exit Sub
    End If

    Dim strModel As String, strType As String, strVersion As String, PB As String, uploadPowerCode As Boolean
    Dim hpsn As String, FTState As String
    strModel = Trim(TextModel.Text)
    strVersion = Trim(TextVersion.Text)
    uploadPowerCode = CBool(TextPowerCode.Text)
    FTState = TextFTState.Text
    hpsn = Trim(TextHPSN.Text)
    
    
    
    If (chkY2.Value = 1) Then
        PB = CommonValue.PB_Y2
    ElseIf (chkY.Value = 1) Then
        PB = CommonValue.PB_Y
    ElseIf (chkN.Value = 1) Then
        PB = CommonValue.PB_N
    ElseIf (chkN4.Value = 1) Then
        PB = CommonValue.PB_N4
    End If


    If UploadH3C_PB(PB, Trim(UCase(txtSN.Text)), strVersion, "NA", "N/A", "CHINA", "frmHW3COM") = False Then
        MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        Exit Sub
    End If


        If UploadH3CInfo2(uploadPowerCode, Me.txtSN.Text, strVersion, FTState, "", "CHINA", golUSERNAME, PB) = False Then
            MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
            txtSN.SetFocus
            Exit Sub
        End If
        

        OpenLppx2 strModel

        myVars2.Item("SN").Value = UCase(txtSN.Text)
        myVars2.Item("Rev").Value = UCase(TextVersion.Text)
        myDoc2.PrintLabel 1
        myDoc2.FormFeed
        
'===============add by ben 2012-02-05 start===============
        Call Connect.addPrintedLabel(Me.txtSN.Text, Me.Name)
'===============add by ben 2012-02-05 end=================

        UnloadLppx2
        
        cmdCancel_SN_Click
        
        
        If CheckHpPackLabel.Value = 1 Then
            frmHW3COM.Hide
    
            FormHPFahuo3COM.txtSN = hpsn
            FormHPFahuo3COM.txtModel_hid = strModel
    
    
            FormHPFahuo3COM.Show
            Call FormHPFahuo3COM.cmdMPrint_Click
        End If
        
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

Private Sub cmdCancel_SN_Click()
txtSN.Text = ""
txtSN.SetFocus


chkY.Value = 0
chkY.Enabled = False
chkY2.Value = 0
chkY2.Enabled = False
chkN.Value = 0
chkN.Enabled = False
chkN4.Value = 0
chkN4.Enabled = False

End Sub


Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If Len(txtSN.Text) < 10 Then
            MsgBox "产品序号长度不能小于10!"
            txtSN.SetFocus
            Exit Sub
        End If
    
        Dim snstring As String
        Dim verstring As String
        snstring = Trim(Me.txtSN.Text)
    
        Dim strModel As String
        Dim strVer As String
        Dim strIII As String
        Dim str2 As String
        Dim print_sv As String, print_power As String, FTState As String
               
        
        HP_pack_label = False
        
        '==================
       
        Dim con13 As ADODB.Connection
        Dim rs13 As ADODB.Recordset
        Dim com As ADODB.Command
        Dim str As String, part_number As String
        Set con13 = New ADODB.Connection
        Set rs13 = New ADODB.Recordset
        strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
        'con13.ConnectionTimeout = 50
        con13.Open ConnectionString:=strConn
        Set com = New ADODB.Command
        com.ActiveConnection = con13
        'str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtSN.Text) & "'"
        'str = " select top 1 a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "'"
        str = " select top 1 part_number,part_revision,creation_time,order_number,unit_key from (" & _
        "select a.part_number,a.part_revision,a.creation_time,a.unit_key,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
        "select top 1 a.part_number,a.part_revision,a.creation_time,a.unit_key,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
        "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
        com.CommandText = str
        rs13.Open Source:=com
        'rs13.Open str
        If rs13.EOF = True Then
            MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
'                Me.cmdCancel.Click
            rs13.Close
            Exit Sub
        Else
            strModel = Mid(Trim(rs13.Fields(0)), 4, 8)
            strVer = rs13.Fields(1)
            unit_key = rs13.Fields(4)
            part_number = rs13.Fields(0)
            
        End If
        If rs13.State = 1 Then
            rs13.Close
        End If
        If con13.State = 1 Then
            con13.Close
        End If
        
        If checkWeighInformation(part_number, strVer) = False Then
            Exit Sub
        End If
        'add by allen yan 2014/05/20
        'the main purpose of this function is to block the ECO versions that are disabled.
        If IsValidECOVersion(part_number, strVer) = False Then
            Exit Sub
        End If
        
        If strModel <> "" And strVer <> "" Then
        '926FEDSDAE704
            '+++++++++++++++++++++
            hpsn = ""
            strIII = ""
            Dim checkhp As New ADODB.Recordset
            Dim con As New ADODB.Connection
            strConn = "Provider=SQLOLEDB.1; Data Source=server08;Initial Catalog=Print; User ID=sa; PWD=sa"
            con.Open ConnectionString:=strConn
            If con.State = 0 Then
                con.Open
            End If
            
            sql = "SELECT Label,hp_sn FROM H3C_HP with(NOLOCK) where part_number='" & strModel & "'"
            rec.Open sql, con, adOpenForwardOnly, adLockReadOnly
            If Not rec.EOF Then
                If rec("Label") = "Yes" Then
                    HP_pack_label = True
                    strIII = rec("hp_sn")
                End If
            End If
            If rec.State = 1 Then rec.Close
            sql = "SELECT Print_SV,Print_Power,[5000_State] FROM tblOthers with(NOLOCK) where part_number='" & strModel & "' and Part_revision = '" & strVer & "'"
            rec.Open sql, con, adOpenForwardOnly, adLockReadOnly
            If Not rec.EOF Then
                print_sv = rec("print_sv")
                print_power = rec("print_power")
                FTState = rec.Fields(2)
            Else
                MsgBox ("当前机种" & strModel & "版本:" & strVer & "未设定5000状态")
                txtSN.Text = ""
                txtSN.SetFocus
                rec.Close
                Exit Sub
            End If
            If rec.State = 1 Then rec.Close
            
            If HP_pack_label = True Then
      
                
                sql = "select top 1 component_SN from DC_Component_SN where unit_key = " & CStr(unit_key) & " and remark = 'HP' order by last_modified_time DESC"
                con13.Open
                checkhp.Open sql, con13, adOpenKeyset, adLockReadOnly
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
                con13.Close
        
            End If
            '+++++++++++++++++++++
            
            
            
            'strIII = Mid(Trim(txtSN.Text), 5, 3)
            verstring = strVer

            
            Set fs = CreateObject("Scripting.FileSystemObject")

            Dim strDir As String
            strDir = "\\10.11.1.25\Public\Manufacture\标签模板\3Com发货标签\" & strModel & ".lab"
            If Not fs.FileExists(strDir) Then
                MsgBox "没有对应机种打印模板", vbOKOnly + vbExclamation, "警告"
                cmdCancel_SN_Click
                rs3.Close
                Exit Sub
            End If
            
            '==============================


            If verstring = "" Then
                MsgBox ("DS版本未带出，不能打印！")
                Exit Sub
            End If

'===============add by ben 2012-02-05 start===============
                If reprint = False Then
                    If Connect.isPrintedLabel(Me.txtSN.Text, Me.Name) Then
                        MsgBox ("此序列号已打印！")
                        cmdCancel_SN_Click
                        Exit Sub
                    End If
                End If
'===============add by ben 2012-02-05 end=================


            Dim pc As Boolean
            If print_power = "True" Then
                pc = True
            Else
                pc = False
            End If


    Dim lh As New Label_History, PB As String
    Dim sn As String
    sn = txtSN.Text
    If (lh.Init(sn)) Then
        If lh.PB = "Y*" Then
            chkY.Value = 1
            chkY2.Value = 0
            chkN.Value = 0
            chkN4.Value = 0
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
            PB = CommonValue.PB_Y
        ElseIf lh.PB = "Y2" Then
            chkY.Value = 0
            chkY2.Value = 1
            chkN.Value = 0
            chkN4.Value = 0
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
            PB = CommonValue.PB_Y2
        ElseIf lh.PB = "N*" Then
            chkY.Value = 0
            chkY2.Value = 0
            chkN.Value = 1
            chkN4.Value = 0
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
            PB = CommonValue.PB_N
        ElseIf lh.PB = "N4" Then
            chkY.Value = 0
            chkY2.Value = 0
            chkN.Value = 0
            chkN4.Value = 1
            chkY.Enabled = False
            chkY2.Enabled = False
            chkN.Enabled = False
            chkN4.Enabled = False
            PB = CommonValue.PB_N4
        End If
    Else
        chkY.Enabled = True
        chkY2.Enabled = True
        chkN.Enabled = True
        chkN4.Enabled = True
        TextModel.Text = strModel
        TextVersion.Text = verstring
        TextPowerCode.Text = CStr(pc)
        TextFTState.Text = FTState
        If HP_pack_label Then
            CheckHpPackLabel.Value = 1
        Else
            CheckHpPackLabel.Value = 0
        End If
        TextHPSN.Text = hpsn
        MsgBox "请确认环保属性值"
        Exit Sub
    End If
         
            If UploadH3CInfo2(pc, Me.txtSN.Text, strVer, FTState, "", "CHINA", golUSERNAME, PB) = False Then
                MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
                txtSN.SetFocus
                Exit Sub
            End If
            

            OpenLppx2 strModel

            myVars2.Item("SN").Value = UCase(snstring)
            myVars2.Item("Rev").Value = UCase(verstring)
            myDoc2.PrintLabel 1
            myDoc2.FormFeed
            
'===============add by ben 2012-02-05 start===============
            Call Connect.addPrintedLabel(Me.txtSN.Text, Me.Name)
'===============add by ben 2012-02-05 end=================

            UnloadLppx2
            '======================================
            
        End If
        If (rs.State = 1) Then
            rs3.Close
        End If
        
        
        
        
        cmdCancel_SN_Click
        
        
        If HP_pack_label = True Then
            frmHW3COM.Hide
    
            FormHPFahuo3COM.txtSN = hpsn
            FormHPFahuo3COM.txtModel_hid = strModel
    
    
            FormHPFahuo3COM.Show
            Call FormHPFahuo3COM.cmdMPrint_Click
        End If
   End If
End Sub

Private Sub OpenLppx2(model As String)
    Me.MousePointer = vbHourglass

    Set myDoc2 = myApp2.Documents.Open("\\10.11.1.25\Public\Manufacture\标签模板\3Com发货标签\" & model & ".lab")
    
    Me.MousePointer = vbDefault
    Set myVars2 = myDoc2.Variables
    Set myObjs2 = myDoc2.DocObjects
End Sub

Private Sub UnloadLppx2()
    myApp2.Documents.CloseAll False
    myApp2.Quit
    Set myApp2 = Nothing
End Sub
Private Function verifyPB() As Boolean
    If (chkY2.Value + chkY.Value + chkN.Value + chkN4.Value = 1) = False Then
        MsgBox "请确认环保属性"
        verifyPB = False
        Exit Function
    End If
    verifyPB = True
End Function
