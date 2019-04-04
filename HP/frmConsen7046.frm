VERSION 5.00
Begin VB.Form frmConsen7046 
   Caption         =   "海康SN标签"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11340
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtOrder 
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
      Left            =   6240
      TabIndex        =   32
      Top             =   5880
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   360
      Picture         =   "frmConsen7046.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   31
      Top             =   3720
      Width           =   615
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
      Left            =   3480
      TabIndex        =   29
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtPower 
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
      Left            =   2040
      TabIndex        =   27
      Top             =   3120
      Width           =   3735
   End
   Begin VB.PictureBox Picture3 
      Height          =   615
      Left            =   360
      Picture         =   "frmConsen7046.frx":071D
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   26
      Top             =   4440
      Width           =   615
   End
   Begin VB.CheckBox chkNonChinaRoHS 
      BackColor       =   &H0000C000&
      Caption         =   "无"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Top             =   3840
      Width           =   735
   End
   Begin VB.CheckBox chkNonWEEE 
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
      Left            =   2400
      TabIndex        =   24
      Top             =   4440
      Width           =   735
   End
   Begin VB.CheckBox chkWEEE 
      Caption         =   "有"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   23
      Top             =   4440
      Width           =   735
   End
   Begin VB.CheckBox chkChinaRoHS 
      Caption         =   "有"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   22
      Top             =   3840
      Width           =   615
   End
   Begin VB.CheckBox chkLaser 
      Caption         =   "有"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   21
      Top             =   5280
      Width           =   735
   End
   Begin VB.CheckBox chkNonLaser 
      BackColor       =   &H0000C000&
      Caption         =   "无"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   5280
      Width           =   735
   End
   Begin VB.PictureBox Picture14 
      Height          =   975
      Left            =   360
      Picture         =   "frmConsen7046.frx":20AF
      ScaleHeight     =   915
      ScaleWidth      =   2355
      TabIndex        =   19
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CheckBox chkCCC 
      Caption         =   "有"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   3840
      Width           =   735
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
      Left            =   5160
      TabIndex        =   17
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox Picture15 
      Height          =   495
      Left            =   3360
      Picture         =   "frmConsen7046.frx":3784
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txtModel 
      Height          =   285
      Left            =   7320
      TabIndex        =   15
      Top             =   7440
      Width           =   495
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
      Left            =   1200
      TabIndex        =   13
      Top             =   240
      Width           =   1335
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
      Left            =   2040
      TabIndex        =   8
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox txtDesc 
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
      Left            =   2040
      TabIndex        =   7
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox txtType 
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
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtMaterial 
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
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   3735
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
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   6720
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
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   6720
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
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   6720
      Width           =   1095
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
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   5220
      Left            =   6000
      Picture         =   "frmConsen7046.frx":3E8A
      ScaleHeight     =   5160
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label7 
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
      Left            =   2640
      TabIndex        =   30
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "电源:"
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
      Left            =   360
      TabIndex        =   28
      Top             =   3240
      Width           =   1815
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
      Left            =   360
      TabIndex        =   14
      Top             =   240
      Width           =   735
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
      Left            =   360
      TabIndex        =   12
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "产品类别:"
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
      Left            =   360
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "产品型号:"
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
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      Caption         =   "物料代码:"
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
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "frmConsen7046"
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
Dim newLableFlag As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Sub cmdCancel_HPSN_Click()
    Me.txtWorkOrder.Text = ""
    Me.txtPart.Text = ""
    Me.txtDesc.Text = ""
    Me.txtType.Text = ""
    Me.txtMaterial.Text = ""
    Me.txtPower.Text = ""
    Me.chkChinaRoHS.Value = 0
    Me.chkNonChinaRoHS.Value = 0
    Me.chkCCC.Value = 0
    Me.chkNonCCC.Value = 0
    Me.chkWEEE.Value = 0
    Me.chkNonWEEE.Value = 0
    Me.chkLaser.Value = 0
    Me.chkNonLaser.Value = 0
    Me.txtModel.Text = ""
    Me.txtSN.Text = ""
    Me.txtOrder.Text = ""

    
    
End Sub

Private Sub cmdMPrint_Click()
    Dim model As String
    If Me.txtPart.Text = "" And txtModel.Text = "" Then
        MsgBox "该机种信息不能打印!"
        Exit Sub
    ElseIf Me.txtPart.Text <> "" Then
        model = Trim(txtPart.Text)
    ElseIf txtModel.Text <> "" Then
        model = Trim(txtModel.Text)
    End If
    

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
        If Len(txtSN.Text) <> 9 Then
            MsgBox "产品序号长度不等于9!"
'            txtSN.SetFocus
            Exit Sub
        End If
'        If InStr(1, txtPart.Text, txtModel.Text) <= 0 Then
        If Trim(txtPart.Text) <> Trim(txtModel.Text) Then
            MsgBox ("该工单料号和条码对应的料号不一致，请确认输入工单是否正确!")
            rs.Close
            Exit Sub
        End If
'        updateHPInformation
        
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
'del_excel
del_sql
cmdReturn_HPSN.Enabled = True
'cmdPrint_HPSN.Enabled = True
cmdCancel_HPSN.Enabled = True
'MsgBox ("批量打印成功！")
End Sub

Private Sub cmdPrint_HPSN_Click()
    Dim Pb As String
    
    If txtSN.Text = "" Then
        MsgBox ("序列号未输入，不能打印！")
'        txtSN.SetFocus
        Exit Sub
    End If

    If txtDesc.Text = "" Then
        MsgBox ("产品类别未带出，不能打印！")
        Exit Sub
    End If
    If txtModel.Text = "" Then
        MsgBox ("导入资料中ITEM_CODE栏不能为空！")
        Exit Sub
    End If
    OpenLppx
    myVars.Item("SN").Value = UCase(Trim(txtSN.Text))
    myVars.Item("Product Name1").Value = UCase(Trim(txtDesc.Text))
    myVars.Item("Product ID").Value = UCase(Trim(txtType.Text))
    myVars.Item("WLDM").Value = UCase(Trim(txtMaterial.Text))
    If Trim(txtPower.Text) = "" Then
        myObjs("DY").Top = 10000
        myObjs("Text1(3)").Top = 1000
    Else
        myVars.Item("DY").Value = UCase(Trim(txtPower.Text))
    End If

    If chkNonChinaRoHS.Value = 1 Then
      myObjs("rohs").Top = 10000
    End If
    
    If chkNonCCC.Value = 1 Then
      myObjs("CCC").Top = 10000
    End If
    
    If chkNonWEEE.Value = 1 Then
      myObjs("WEEE").Top = 10000
    End If
    
    If chkNonWEEE.Value = 1 Then
      myObjs("Image3").Top = 10000
    End If
'    If (Me.chkY.Value = 1) Then
'        Pb = "N*"
'        OverridePb Pb
'        myVars.Item("Rohs").Value = Pb
'    ElseIf (Me.chkY2.Value = 1) Then
'        Pb = "Y2"
'        OverridePb Pb
'        myVars.Item("Rohs").Value = Pb
'    Else
'        MsgBox "环保属性未选择，不能打印!"
'        txtSN.SetFocus
'        UnloadLppx
'        Exit Sub
'    End If
    
'    If UploadH3CInfo(Pb, Trim(UCase(txtSN.Text)), Trim(UCase(txtRevision.Text)), "NA", "N/A", "CHINA", "frmConsen7046") = False Then
'        MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
'        txtSN.SetFocus
'        UnloadLppx
'        Exit Sub
'    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
'    If UploadH3C_PB(Pb, Trim(UCase(txtSN.Text)), Trim(UCase(txtRevision.Text)), "NA", "N/A", "CHINA", "frmConsen7046") = False Then
'        MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
'        txtSN.SetFocus
'        UnloadLppx
'        Exit Sub
'    End If

    If UploadConsen_Rec(Trim(UCase(txtSN.Text)), Trim(UCase(txtWorkOrder.Text)), Trim(UCase(txtPart.Text)), Trim(UCase(txtDesc.Text)), Trim(UCase(txtType.Text)), Trim(UCase(txtMaterial.Text)), Trim(UCase(txtPower.Text)), Trim(UCase(txtOrder.Text)), IIf(chkCCC.Value, "1", "0"), IIf(chkChinaRoHS.Value, "1", "0"), IIf(chkWEEE.Value, "1", "0"), IIf(chkLaser.Value, "1", "0"), "frmConsen7046") = False Then
        MsgBox "Consen资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
'        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
    
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
    
          If IsNull(rec.Fields("hp_desc1")) Then
              MsgBox ("此序列号未维护描述信息!")
              rec.Close
              Exit Sub
          Else
              txtDesc1 = rec.Fields("hp_desc1")
          End If
    
          If Not IsNull(rec.Fields("hp_desc2")) Then
              txtDesc2 = rec.Fields("hp_desc2")
          End If
          
          If IsNull(rec.Fields("new_label")) Or Trim(rec.Fields("new_label")) = "" Then
              newLableFlag = False
              MsgBox ("此机种序列号没有维护new_label选项，请联系ME!")
              rec.Close
              Exit Sub
          Else
              newLableFlag = True
          End If
      End If
    
End Sub

Private Sub getConsenInformation()

'      sql = "select * from hp where hp_sn_iii=substring('" & Trim(txtSN.Text) & "',5,3) and h3c_bom_code = '" + txtModel.Text + "'"
      sql = "select * from tblConsenNew where Part_Number= '" + txtPart.Text + "'"
      If rec.State = 1 Then
        rec.Close
      End If
      
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      
      If rec.EOF = True Then
          MsgBox "此序列号未维护信息!"
          txtSN.Text = ""
'          txtSN.SetFocus
          rec.Close
          Exit Sub
      Else
          If IsNull(rec.Fields("Part_Number")) Then
              MsgBox ("此序列号未维护产品编码!")
              rec.Close
              Exit Sub
          Else
              txtPart.Text = rec.Fields("Part_Number")
          End If
'          hpsnproduct
    
          If IsNull(rec.Fields("Part_Type")) Then
              MsgBox ("此序列号未维护产品类别!")
              rec.Close
              Exit Sub
          Else
              txtDesc.Text = rec.Fields("Part_Type")
          End If
          
          If IsNull(rec.Fields("Part_Model")) Then
              MsgBox ("此序列号未维护产品型号!")
              rec.Close
              Exit Sub
          Else
              txtType.Text = rec.Fields("Part_Model")
          End If
    
          If IsNull(rec.Fields("Material")) Then
              MsgBox ("此序列号未维护物料代码")
              rec.Close
              Exit Sub
          Else
              txtMaterial.Text = rec.Fields("Material")
          End If
          
          If IsNull(rec.Fields("Power")) Then
'              MsgBox ("此序列号未维护电源!")
'              rec.Close
'              Exit Sub
                txtPower.Text = ""
          Else
              txtPower.Text = rec.Fields("Power")
          End If
          
          If IsNull(rec.Fields("SalesOrder")) Then
'              MsgBox ("此序列号未维护电源!")
'              rec.Close
'              Exit Sub
                txtOrder.Text = ""
          Else
              txtOrder.Text = rec.Fields("SalesOrder")
          End If

        If UCase(Trim(rec.Fields("ChinaRoHS"))) = True Then
           chkChinaRoHS.Value = 1
           chkNonChinaRoHS.Value = 0
        Else
           chkChinaRoHS.Value = 0
           chkNonChinaRoHS.Value = 1
        End If
        
        If UCase(Trim(rec.Fields("WEEE"))) = True Then
           chkWEEE.Value = 1
           chkNonWEEE.Value = 0
        Else
           chkWEEE.Value = 0
           chkNonWEEE.Value = 1
        End If
        
        If UCase(Trim(rec.Fields("CCC"))) = True Then
           chkCCC.Value = 1
           chkNonCCC.Value = 0
        Else
           chkCCC.Value = 0
           chkNonCCC.Value = 1
        End If
        If UCase(Trim(rec.Fields("Laser"))) = True Then
           chkLaser.Value = 1
           chkNonLaser.Value = 0
        Else
           chkLaser.Value = 0
           chkNonLaser.Value = 1
        End If
      End If
    
End Sub

Private Sub txtWorkOrder_KeyPress(KeyAscii As Integer)
    Me.txtPart.Text = ""
    Me.txtDesc.Text = ""
    Me.txtType.Text = ""
    Me.txtMaterial.Text = ""
    Me.txtPower.Text = ""
    Me.chkChinaRoHS.Value = 0
    Me.chkNonChinaRoHS.Value = 0
    Me.chkCCC.Value = 0
    Me.chkNonCCC.Value = 0
    Me.chkWEEE.Value = 0
    Me.chkNonWEEE.Value = 0
    Me.chkLaser.Value = 0
    Me.chkNonLaser.Value = 0
    Me.txtModel.Text = ""
    Me.txtSN.Text = ""
    Me.txtOrder.Text = ""
    
    If KeyAscii = 13 Then
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
'
                Me.txtPart.Text = rec.Fields("part_number")
                getConsenInformation
'                Me.txtRevision.Text = rec.Fields("part_revision")
                
                
'                If Connect.getPartList(Trim(Me.txtWorkOrder.Text)) <> "" Then
'                    If (Connect.GetPBState(Connect.getPartList(Trim(Me.txtWorkOrder.Text))) = "NPb") Then
'                        Me.chkY2.Value = 1
'                        Me.chkY.Value = 0
'                    Else
'                        Me.chkY.Value = 1
'                        Me.chkY2.Value = 0
'                    End If
'                End If
            End If
        End If
    End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub
Private Sub OpenLppx()
'
'        If newLableFlag = True Then
            Me.MousePointer = vbHourglass
            Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\海康模板\" & "海康本体标签.lab")
            
'        Else
'            Me.MousePointer = vbHourglass
'            Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\HP本体标签正向\" & "HP SN标签5020.lab")
'        End If
    
'    Me.MousePointer = vbHourglass
'    Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\HP本体标签正向\" & "HP SN标签5020.lab")
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

