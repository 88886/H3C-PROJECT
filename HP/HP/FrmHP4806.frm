VERSION 5.00
Begin VB.Form FrmHP4806 
   Caption         =   "HP SN标签48*06"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8415
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtModel 
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Top             =   5880
      Width           =   735
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
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
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
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   10
      Top             =   4560
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
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "打 印"
      Enabled         =   0   'False
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
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txtProduct 
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
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox txtSN 
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
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtWorkOrder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtPart 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox txtRevision 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CheckBox chkY 
      Caption         =   "Y*"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   3840
      Width           =   615
   End
   Begin VB.CheckBox chkY2 
      Caption         =   "Y2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   1440
      Picture         =   "FrmHP4806.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "产品编号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "产品序列号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "工单:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "机种:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "版本:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000004&
      Caption         =   "环保属性:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "FrmHP4806"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New Recordset
Dim bom_code As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim rs As New Recordset
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdCancel_HPSN_Click()
    Me.txtSN.Text = ""
    Me.txtModel.Text = ""
    Me.txtProduct.Text = ""
    Me.txtWorkOrder.Text = ""
    Me.txtPart.Text = ""
    Me.txtRevision.Text = ""
    Me.chkY.Value = 0
    Me.chkY2.Value = 0
'    UnloadLppx
    
End Sub

Private Sub cmdMPrint_Click()
 Dim model As String
 If Me.txtPart.Text = "" And txtModel.Text = "" Then
    MsgBox "该机种信息不能打印!"
    Exit Sub
 ElseIf Me.txtPart.Text <> "" Then
    model = Mid(txtPart.Text, 4, 8)
 ElseIf txtModel.Text <> "" Then
    model = Trim(txtModel.Text)
 End If
 If Connect.checkPrintPreCondition(model, 4) = False Then
    MsgBox "该机种没有在HP序列号类型维护为[48*06]打印选项!"
'    txtSN.SetFocus
    Exit Sub
 End If
cmdReturn_HPSN.Enabled = False
'cmdPrint_HPSN.Enabled = False
cmdCancel_HPSN.Enabled = False
sql = "select ITEM_CODE,BARCODE from tblHP_Print where isnull(BARCODE,'')<>'' and isnull(ITEM_CODE,'')<>'' order by BARCODE"
If conn1.State = 0 Then
    conn1.Open
End If

rs.Open sql, conn1, adOpenStatic, adLockReadOnly
If rs.EOF = True Then
    MsgBox ("序列号未导入！")
    rs.Close
    cmdReturn_HPSN.Enabled = True
    'cmdPrint_HPSN.Enabled = True
    cmdCancel_HPSN.Enabled = True
    Exit Sub
Else
    For i = 1 To rs.RecordCount
    
        txtSN.Text = rs("BARCODE")
        txtModel.Text = rs("ITEM_CODE")
        'begin
        If Len(txtSN.Text) < 10 Then
            MsgBox "产品序号长度不能小于10!"
            txtSN.SetFocus
            Exit Sub
        End If
        If InStr(1, txtPart.Text, txtModel.Text) <= 0 Then
            MsgBox ("该工单料号和条码对应的料号不一致，请确认输入工单是否正确!")
            rec.Close
            Exit Sub
        End If
        updateHPInformation
        cmdPrint_HPSN_Click
        rs.MoveNext
        If i Mod 100 = 0 Then
            Sleep (1000 * 10)
       End If
    Next
    UnloadLppx
    cmdCancel_HPSN_Click
    rs.Close
End If
del_sql
cmdReturn_HPSN.Enabled = True
'cmdPrint_HPSN.Enabled = True
cmdCancel_HPSN.Enabled = True
End Sub

Private Sub cmdPrint_HPSN_Click()
    Dim Pb As String
    
    If txtSN.Text = "" Then
        MsgBox ("序列号未输入，不能打印！")
        txtSN.SetFocus
        Exit Sub
    End If
    If txtProduct.Text = "" Then
        MsgBox ("产品编码未带出，不能打印！")
        Exit Sub
    End If
    OpenLppx
    myVars.Item("SN").Value = UCase(txtSN.Text)
    myVars.Item("PN").Value = UCase(txtProduct.Text)
    myVars.Item("Model").Value = UCase(txtModel.Text)
    myVars.Item("Rev").Value = UCase(txtRevision.Text)
    If (Me.chkY.Value = 1) Then
        Pb = "Y*"
        OverridePb Pb
        myVars.Item("Rohs").Value = Pb
    ElseIf (Me.chkY2.Value = 1) Then
        Pb = "Y2"
        OverridePb Pb
        myVars.Item("Rohs").Value = Pb
    Else
        MsgBox "环保属性未选择，不能打印!"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    
    '======Add by mike 2015.3.24 for data upload to FTPC============
    If UploadH3C_PB(Pb, Trim(UCase(txtSN.Text)), Trim(UCase(txtRevision.Text)), "NA", "N/A", "CHINA", "frmHP4806") = False Then
        MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
    
'
'    If txtDesc2.Text <> "" Then
'        myVars.Item("ID-1").Value = txtDesc1.Text
'        myVars.Item("ID-2").Value = txtDesc2.Text
'    Else
'        myVars.Item("ID-1").Value = txtDesc1.Text
'        myVars.Item("ID-2").Value = ""
'    End If
    'OpenLppx
    myDoc.PrintLabel 1
    myDoc.FormFeed
End Sub

Private Sub OverridePb(ByRef Pb As String)
    Dim labelHistory As New Label_History
    Dim sn As String
    sn = txtSN.Text
    If labelHistory.Init(sn) Then
        Pb = labelHistory.Pb
    End If
End Sub

Private Sub cmdReturn_HPSN_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    If conn1.State = 0 Then
      conn1.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
      conn1.Open
   End If
   If connFTPC.State = 0 Then
        connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
        connFTPC.Open
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
   
   If connFTPC.State = 1 Then
        connFTPC.Close
        Set connFTPC = Nothing
   End If
   UnloadLppx
End Sub


Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub
Private Sub OpenLppx()
    Me.MousePointer = vbHourglass
    Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\HP本体标签正向\" & "HP SN标签486.lab")
'    If txtDesc2.Text = "" Then
'        Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP序列号标签小于30位描述.lab")
'    Else
'        Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP序列号标签大于30位描述.lab")
'    End If
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub

Sub del_sql()
    Dim delsql As String
    delsql = "delete from tblHP_Print"
    conn1.Execute delsql
End Sub

Private Sub updateHPInformation()
      sql = "select * from hp where hp_sn_iii=substring('" & Trim(txtSN.Text) & "',5,3) and h3c_bom_code = '" + txtModel.Text + "'"
      If rec.State = 1 Then
        rec.Close
      End If
      
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      
      If rec.EOF = True Then
          MsgBox "此序列号未维护信息!"
          txtSN.Text = ""
          txtSN.SetFocus
          rec.Close
          Exit Sub
      Else
          If IsNull(rec.Fields("hpsnproduct")) Then
              MsgBox ("此序列号未维护产品编码!")
              rec.Close
              Exit Sub
          Else
              txtProduct = rec.Fields("hpsnproduct")
          End If
'          hpsnproduct
    
'          If IsNull(rec.Fields("hp_desc1")) Then
'              MsgBox ("此序列号未维护描述信息!")
'              rec.Close
'              Exit Sub
'          Else
'              txtDesc1 = rec.Fields("hp_desc1")
'          End If
'
'          If Not IsNull(rec.Fields("hp_desc2")) Then
'              txtDesc2 = rec.Fields("hp_desc2")
'          End If
      End If
    
End Sub


Private Sub txtWorkOrder_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Me.txtPart.Text = ""
        Me.txtRevision.Text = ""
        Me.chkY.Value = 0
        Me.chkY2.Value = 0
        If Trim(Me.txtWorkOrder.Text) <> "" Then
            If rec.State = 1 Then
                rec.Close
            End If
           
            If connFTPC.State = 0 Then
                connFTPC.Open
            End If
            Dim tempWO As String
         
            tempWO = Trim(Me.txtWorkOrder.Text)
            sql = "select b.part_number,b.part_revision from WORK_ORDER a,WORK_ORDER_ITEMS b where a.order_key = b.order_key and a.order_number = '" & tempWO & "'"
            rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                MsgBox ("该工单不存在，请确认输入工单是否正确!")
                rec.Close
                Exit Sub
            Else

                Me.txtPart.Text = rec.Fields("part_number")
                Me.txtRevision.Text = rec.Fields("part_revision")
                
                
                If Connect.IsCheckRohs(Mid(Me.txtPart.Text, 4, 8)) = True Then
                    If Connect.getPartList(Trim(Me.txtWorkOrder.Text)) <> "" Then
                        lPB = Connect.GetPBState(Connect.getPartList(Trim(Me.txtWorkOrder.Text)))
                        If (lPB = "NPb") Then
                            Me.chkY2.Value = 1
                            Me.chkY.Value = 0
                        ElseIf (lPB = "Pb") Then
                            Me.chkY2.Value = 0
                            Me.chkY.Value = 1
                        Else
                            Me.chkY.Value = 0
                            Me.chkY2.Value = 0
                        End If
                    Else
                        MsgBox ("该工单ROHS的属性不存在，请确认输入工单是否正确!")
                        rec.Close
                        Exit Sub
                    End If
                Else
                    Me.chkY.Value = 1
                    Me.chkY2.Value = 0
                End If
            End If
        End If
    End If
End Sub
