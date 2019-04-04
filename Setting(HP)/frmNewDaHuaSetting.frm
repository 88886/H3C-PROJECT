VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmNewDaHuaSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Da Hua Setting"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   17325
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   17325
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fmH3C 
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   16935
      Begin VB.TextBox txtEN 
         Height          =   495
         Left            =   9840
         TabIndex        =   42
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H0000C000&
         Caption         =   "无"
         Height          =   375
         Left            =   2760
         TabIndex        =   27
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H0000C000&
         Caption         =   "无"
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
         Left            =   6120
         TabIndex        =   26
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtSN 
         Height          =   450
         Left            =   1800
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkWEEE 
         Caption         =   "有"
         Height          =   375
         Left            =   5160
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txXXXXXX 
         Height          =   495
         Left            =   25200
         TabIndex        =   23
         Text            =   "sdgfdsfadsfadsf"
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox chkRoHS 
         Caption         =   "有"
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtDesc 
         Height          =   450
         Left            =   5760
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtRev 
         Height          =   450
         Left            =   14520
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtModel 
         Height          =   450
         Left            =   9840
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkCCC 
         Caption         =   "有"
         Height          =   375
         Left            =   9240
         TabIndex        =   18
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox chkNonCCC 
         BackColor       =   &H0000C000&
         Caption         =   "无"
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
         Left            =   10200
         TabIndex        =   17
         Top             =   2160
         Width           =   735
      End
      Begin VB.PictureBox Picture16 
         Height          =   615
         Left            =   840
         Picture         =   "frmNewDaHuaSetting.frx":0000
         ScaleHeight     =   555
         ScaleWidth      =   675
         TabIndex        =   16
         Top             =   2040
         Width           =   735
      End
      Begin VB.PictureBox Picture3 
         Height          =   615
         Left            =   4320
         Picture         =   "frmNewDaHuaSetting.frx":1566
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   15
         Top             =   2040
         Width           =   615
      End
      Begin VB.PictureBox Picture15 
         Height          =   495
         Left            =   8400
         Picture         =   "frmNewDaHuaSetting.frx":2EF8
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   14
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtDCSign 
         Height          =   450
         Left            =   9840
         TabIndex        =   13
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtWeight 
         Height          =   450
         Left            =   14520
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtACSign 
         Height          =   450
         Left            =   5760
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtExeStandard 
         Height          =   450
         Left            =   1800
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtSize 
         Height          =   450
         Left            =   5760
         TabIndex        =   9
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtCustomerCode 
         Height          =   450
         Left            =   1800
         TabIndex        =   8
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label 英文描述 
         Caption         =   "英文描述:"
         Height          =   375
         Left            =   8160
         TabIndex        =   41
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblSN 
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblHV 
         Caption         =   "硬件版本:"
         Height          =   495
         Left            =   -1920
         TabIndex        =   38
         Top             =   7560
         Width           =   1455
      End
      Begin VB.Label lblRemark 
         Caption         =   "备注:"
         Height          =   495
         Left            =   24120
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "产品类别:"
         Height          =   375
         Left            =   4200
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "整机版本:"
         Height          =   375
         Left            =   12840
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "产品型号:"
         Height          =   375
         Left            =   8160
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "直流符号:"
         Height          =   375
         Left            =   8160
         TabIndex        =   33
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "重量Kg:"
         Height          =   375
         Left            =   12840
         TabIndex        =   32
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "交流符号:"
         Height          =   375
         Left            =   4200
         TabIndex        =   31
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "执行标准:"
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "尺寸:"
         Height          =   375
         Left            =   4200
         TabIndex        =   29
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "客户编码:"
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14400
      TabIndex        =   6
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   12840
      TabIndex        =   5
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "确定(Confirm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11160
      TabIndex        =   4
      Top             =   9840
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(Delete)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14400
      TabIndex        =   3
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修改(Update)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   12840
      TabIndex        =   2
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "新增(Insert)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11160
      TabIndex        =   1
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查询(Query)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9480
      TabIndex        =   0
      Top             =   9480
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgH3C 
      Height          =   5655
      Left            =   0
      TabIndex        =   40
      Top             =   3360
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   9975
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmNewDaHuaSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim op As String
Dim xlApp As New Excel.Application
Dim xlBook As New Excel.Workbook
Dim xlSheet As New Excel.Worksheet
Dim query As Boolean

Private Sub Reset()
    Dim ctr As Control
    Me.chkRoHS.Value = 0
    Me.chkNonRoHS.Value = 0
    Me.chkCCC.Value = 0
    Me.chkNonCCC.Value = 0
    Me.chkWEEE.Value = 0
    Me.chkNonWEEE.Value = 0
    If op = "Insert" Then
        cmdQuery.Enabled = True
        cmdInsert.Enabled = False
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
        cmdConfirm.Enabled = True
        cmdCancel.Enabled = True
        For Each ctr In Me.Controls
            If TypeOf ctr Is TextBox Then
                    ctr.Text = ""
                    ctr.Enabled = True
                    ctr.BackColor = &HFFFFFF
                 ElseIf TypeOf ctr Is ComboBox Then
                    If ctr.Style = 2 Then
                       ctr.ListIndex = -1
                    Else
                       ctr.Text = ""
                    End If
                 ElseIf TypeOf ctr Is CheckBox Then
                    ctr.Value = 0
            End If
        Next

        Me.txtSN.Enabled = True
        Me.txtSN.SetFocus
    ElseIf op = "Cancel" Then
        For Each ctr In Me.Controls
        If TypeOf ctr Is TextBox Then
                ctr.Enabled = False
                ctr.BackColor = &HFFFFFF
             ElseIf TypeOf ctr Is ComboBox Then
                ctr.Enabled = False
                ctr.BackColor = &HFFFFFF
             ElseIf TypeOf ctr Is CheckBox Then
                ctr.Enabled = True
        End If
    Next
        cmdQuery.Enabled = True
        cmdInsert.Enabled = True
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
        cmdConfirm.Enabled = True
        cmdCancel.Enabled = False
    ElseIf op = "Update" Then
        For Each ctr In Me.Controls
        If TypeOf ctr Is TextBox Then
                ctr.Enabled = True
             ElseIf TypeOf ctr Is ComboBox Then
                ctr.Enabled = True
                End If
    Next
        txtSN.Enabled = False
        txtSN.BackColor = &HE0E0E0
    End If
End Sub
Private Sub enable()
   txtSN.Enabled = True
   txtSN.BackColor = &HFFFFFF
   txtDesc.Enabled = True
   txtDesc.BackColor = &HFFFFFF
   txtModel.Enabled = True
   txtModel.BackColor = &HFFFFFF
   txtRev.Enabled = True
   txtRev.BackColor = &HFFFFFF
   txtExeStandard.Enabled = True
   txtExeStandard.BackColor = &HFFFFFF
   txtAC.Enabled = True
   txtAC.BackColor = &HFFFFFF
   txtDC.Enabled = True
   txtDC.BackColor = &HFFFFFF
   txtWeight.Enabled = True
   txtWeight.BackColor = &HFFFFFF
   txtCustomerCode.Enabled = True
   txtCustomerCode.BackColor = &HFFFFFF
   txtSize.Enabled = True
   txtSize.BackColor = &HFFFFFF
   
   chkRoHS.Enabled = True
   chkNonRoHS.Enabled = True
   
   chkCCC.Enabled = True
   chkNonCCC.Enabled = True
   chkWEEE.Enabled = True
   chkNonWEEE.Enabled = True
   
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = False
   cmdUpdate.Enabled = False
   cmdDelete.Enabled = False
   cmdConfirm.Enabled = True
   cmdCancel.Enabled = True
End Sub

Private Sub chkCCC_Click()
    If chkCCC.Value = 1 Then
        chkNonCCC.Value = 0
    Else
        chkNonCCC.Value = 1
    End If
End Sub


Private Sub chkNonCCC_Click()
    If chkNonCCC.Value = 1 Then
        chkCCC.Value = 0
    Else
        chkCCC.Value = 1
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


Private Sub chkRoHS_Click()
    If chkRoHS.Value = 1 Then
        chkNonRoHS.Value = 0
    Else
        chkNonRoHS.Value = 1
    End If
End Sub

Private Sub chkNonRoHS_Click()
    If chkNonRoHS.Value = 1 Then
        chkRoHS.Value = 0
    Else
        chkRoHS.Value = 1
    End If
End Sub


Private Sub cmdCancel_Click()
   op = "Cancel"
   Reset
End Sub

Private Sub cmdConfirm_Click()
   If Trim(txtSN.Text) = "" Then
      MsgBox "产品编码不能为空!!", vbExclamation + vbOKOnly, "产品编码空"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If Trim(txtDesc.Text) = "" Then
      MsgBox "产品类别不能为空!!", vbExclamation + vbOKOnly, "产品类别空"
      txtDesc.SetFocus
      Exit Sub
   End If
   
   If Trim(txtModel.Text) = "" Then
      MsgBox "产品型号不能为空!!", vbExclamation + vbOKOnly, "产品型号空"
      txtModel.SetFocus
      Exit Sub
   End If
   
   If Trim(txtRev.Text) = "" Then
      MsgBox "整机版本不能为空!!", vbExclamation + vbOKOnly, "整机版本空"
      txtRev.SetFocus
      Exit Sub
   End If
   
'   If Trim(txtExeStandard.Text) = "" Then
'      MsgBox "执行标准不能为空!!", vbExclamation + vbOKOnly, "执行标准空"
'      txtExeStandard.SetFocus
'      Exit Sub
'   End If
   If Trim(txtWeight.Text) = "" Then
      MsgBox "重量不能为空!!", vbExclamation + vbOKOnly, "重量单空"
      txtWeight.SetFocus
      Exit Sub
   End If
   If Trim(txtCustomerCode.Text) = "" Then
       MsgBox "客户编码不能为空!", vbExclamation + vbOKOnly, "客户编码空"
       txtCustomerCode.SetFocus
       Exit Sub
   End If

   If Trim(txtSize.Text) = "" Then
      MsgBox "尺寸不能为空!", vbExclamation + vbOKOnly, "尺寸空"
      txtSize.SetFocus
      Exit Sub
   End If
 
  If Trim(txtEN.Text) = "" Then
      MsgBox "英文描述不能为空!", vbExclamation + vbOKOnly, "英文描述空"
      txtSize.SetFocus
      Exit Sub
   End If
   
   
   Dim CE, WEEE, ChinaRoHS, RoHS, TurkeyRoHS, SVPrint, ATick, CTick, ICT, RCM, Gost, KC, ATick_ID, CTick_ID, ICT_ID, RCM_ID, Gost_ID, KC_ID, PWPrint, EAC, CCC, Laser As String

   If chkCCC.Value = 1 Then
      CCC = "1"
   ElseIf chkNonCCC.Value = 1 Then
      CCC = "0"
   Else
      MsgBox "CCC未设定!", vbExclamation + vbOKOnly, "CCC未设定"
      Exit Sub
   End If
   
   If chkRoHS.Value = 1 Then
      RoHS = "1"
   ElseIf chkNonRoHS.Value = 1 Then
      RoHS = "0"
   Else
      MsgBox "RoHS未设定!", vbExclamation + vbOKOnly, "RoHS未设定"
      Exit Sub
   End If

   If chkWEEE.Value = 1 Then
      WEEE = "1"
   ElseIf chkNonWEEE.Value = 1 Then
      WEEE = "0"
   Else
      MsgBox "WEEE未设定!", vbExclamation + vbOKOnly, "RoHS未设定"
      Exit Sub
   End If

   
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from tblDaHuaNew where Part_Number ='" & Trim(txtSN.Text) & "' "
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "产品编码已存在!", vbExclamation + vbOKOnly, "产品编号重复"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close

      sql = "insert [tblDaHuaNew]([Part_Number],[Part_Desc],[Part_Model],[Rev],[ExeStandard],[ACSign],DCSign,Weight,CustomerCode,Size,[RoHS],[WEEE],[CCC],Part_ENDesc) " & _
            " Values('" & Trim(txtSN.Text) & "','" & Trim(txtDesc.Text) & "','" & Trim(txtModel.Text) & "','" & Trim(txtRev.Text) & "','" & Trim(txtExeStandard.Text) & "','" & Trim(txtACSign.Text) & "','" & Trim(txtDCSign.Text) & "','" & Trim(txtWeight.Text) & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtSize.Text) & "'," & RoHS & "," & WEEE & "," & CCC & ",'" & Trim(txtEN.Text) & "')"
             
      sql = sql & " insert tblDaHuaNew_log(CREATE_USER,[Part_Number],[Part_Desc],[Part_Model],[Rev],[ExeStandard],[ACSign],DCSign,Weight,CustomerCode,Size,[RoHS],[WEEE],[CCC],COMMENT,Part_ENDesc) " & _
            " Values('" & golUSERNAME & "','" & Trim(txtSN.Text) & "','" & Trim(txtDesc.Text) & "','" & Trim(txtModel.Text) & "','" & Trim(txtRev.Text) & "','" & Trim(txtExeStandard.Text) & "','" & Trim(txtACSign.Text) & "','" & Trim(txtDCSign.Text) & "','" & Trim(txtWeight.Text) & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtSize.Text) & "'," & RoHS & "," & WEEE & "," & CCC & ",'Insert','" & Trim(txtEN.Text) & "')"
             
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "新增大华设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "新增失败"
      Else
        MsgBox "新增大华设定资料成功!", vbInformation + vbOKOnly, "新增成功"
      End If
      renovate ("")
      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update tblDaHuaNew set Part_Desc='" & Trim(txtDesc.Text) & "',Part_Model='" & Trim(txtModel.Text) & "',Rev ='" & Trim(txtRev.Text) & "',ExeStandard='" & Trim(txtExeStandard.Text) & "',ACSign='" & Trim(txtACSign.Text) & "',DCSign='" & Trim(txtDCSign.Text) & "',Weight='" & Trim(txtWeight.Text) & "',CustomerCode='" & Trim(txtCustomerCode.Text) & "',Size='" & Trim(txtSize.Text) & "',RoHS=" & RoHS & ",WEEE=" & WEEE & ",Part_ENDesc='" & Trim(txtEN.Text) & "' ,CCC=" & CCC & _
            " where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and Part_Number ='" & Trim(txtSN.Text) & "'"
      
      
      sql = sql & " insert tblDaHuaNew_log(CREATE_USER,[Part_Number],[Part_Desc],[Part_Model],[Rev],[ExeStandard],[ACSign],DCSign,Weight,CustomerCode,Size,[RoHS],[WEEE],[CCC],COMMENT,Part_ENDesc) " & _
            " Values('" & golUSERNAME & "','" & Trim(txtSN.Text) & "','" & Trim(txtDesc.Text) & "','" & Trim(txtModel.Text) & "','" & Trim(txtRev.Text) & "','" & Trim(txtExeStandard.Text) & "','" & Trim(txtACSign.Text) & "','" & Trim(txtDCSign.Text) & "','" & Trim(txtWeight.Text) & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtSize.Text) & "'," & RoHS & "," & WEEE & "," & CCC & ",'Update','" & Trim(txtEN.Text) & "')"
     
     Debug.Print
     
     
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "修改大华设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "修改失败"
      Else
         MsgBox "修改大华设定资料成功!", vbInformation + vbOKOnly, "修改成功"
      End If
      renovate ("")
      cmdCancel_Click
   End If
   renovate ("")
End Sub

Private Sub cmdDelete_Click()
   If mfgH3C.RowSel <= 0 Then
      MsgBox "请选择要删除的行!", vbInformation + vbOKOnly, "未选择行"
      Exit Sub
   End If
   
   sql = " insert tblDaHuaNew_log(CREATE_USER,[Part_Number],[Part_Desc],[Part_Model],[Rev],[ExeStandard],[ACSign],DCSign,Weight,CustomerCode,Size,[RoHS],[WEEE],[CCC],COMMENT,Part_ENDesc) " & _
             " select '" & golUSERNAME & "',[Part_Number],[Part_Desc],[Part_Model],[Rev],[ExeStandard],[ACSign],DCSign,Weight,CustomerCode,Size,[RoHS],[WEEE],[CCC],'Delete',Part_ENDesc from tblDaHuaNew " & _
             " where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and Part_Number ='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 2) & "'"

   sql = sql & " delete from tblDaHuaNew where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and Part_Number ='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 2) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "删除大华设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "删除失败"
   End If
      MsgBox "删除大华设定资料成功!", vbInformation + vbOKOnly, "删除成功"
   renovate ("")
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If mfgH3C.Rows = 0 Then
      MsgBox "无资料可汇出", vbExclamation + vbOKOnly, "无资料"
      Exit Sub
   End If
   If txtPath.Text <> "" Then
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets.Item(1)
       For i = 0 To mfgH3C.Rows - 1
         For j = 1 To mfgH3C.Cols - 1
          xlSheet.Cells(i + 1, j) = mfgH3C.TextMatrix(i, j)
       Next j
      Next i
      xlBook.SaveAs (txtPath.Text)
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "汇出到EXCEL资料成功!!", vbInformation + vbOKOnly, "汇出成功"
    End If
End Sub

Private Sub cmdImport_Click()
   If txtPath.Text = "" Then
      MsgBox "导入路径不能为空!", vbExclamation + vbOKOnly, "导入路径空"
      Exit Sub
   End If
   Dim action As Integer
   Dim info As Boolean
   info = True
   Set xlBook = xlApp.Workbooks.Open(txtPath.Text)
      For i = 1 To xlBook.Sheets.Count
       Set xlSheet = xlBook.Sheets.Item(i)
       For j = 2 To xlSheet.Rows.Count
        r = xlSheet.Cells(j, 1)
        If r = "" Then
           Exit For
        Else
          Dim cellValue As String
          Dim cellhvValue As String
          
          Dim isexist As Boolean
          If xlSheet.Cells(j, 19) = "" Then
             MsgBox "导入资料格式不正确!", vbExclamation + vbOKOnly, "格式错误"
             Exit Sub
          End If
          If Not ((xlSheet.Cells(j, 18) = "N") Or (xlSheet.Cells(j, 18) = "Y")) Then
             MsgBox "导入资料格式不正确!", vbExclamation + vbOKOnly, "格式错误"
             Exit Sub
          End If
          isexist = False
          For K = 1 To 19
          '======================================================
           If K = 3 Then
             cellValue = xlSheet.Cells(j, K)
             cellhvValue = xlSheet.Cells(j, 2)
             
             If cellValue = "" Or cellhvValue = "" Then
                MsgBox "导入资料格式不正确!", vbExclamation + vbOKOnly, "格式错误"
                Exit Sub
             End If
             
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from tblH3C where SN='" & cellValue & "' and HV='" & cellhvValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("产品编码&版本已存在!", vbAbortRetryIgnore + vbExclamation, "资料重复")
                End If
                
                If action = vbAbort Then
                   MsgBox "资料导入已终止!!", vbInformation + vbOKOnly, "导入终止"
                   rcd.Close
                   Exit Sub
                ElseIf action = vbIgnore And info = True Then
                   MsgBox "重复产品编号资料不会导入,请稍等..!!", vbInformation + vbOKOnly, "重复不会导入"
                   rcd.Close
                   info = False
                   Exit For
                ElseIf action = vbRetry And info = True Then
                   MsgBox "重复产品编号资料会自动更新,请稍等..!!", vbInformation + vbOKOnly, "重复会自动更新"
                   info = False
                End If
                isexist = True
             Else
                isexist = False
             End If
             rcd.Close
            End If
            '==================================================
            
            If K = 19 Then
               If action = vbRetry Then
                   sql = "Update tblH3C set CPN='" & xlSheet.Cells(j, 4) & "',EPN='" & xlSheet.Cells(j, 5) & "',Des='" & xlSheet.Cells(j, 6) & "',OS='" & xlSheet.Cells(j, 7) & "',GW='" & xlSheet.Cells(j, 8) & "',CE='" & xlSheet.Cells(j, 9) & "',WEEE='" & xlSheet.Cells(j, 10) & "',ChinaRoHS='" & xlSheet.Cells(j, 11) & "'," & _
                        "RoHS='" & xlSheet.Cells(j, 12) & "',TurkeyRohs='" & xlSheet.Cells(j, 13) & "',MS='" & xlSheet.Cells(j, 14) & "',MSValidFrom='" & xlSheet.Cells(j, 15) & "',NAL='" & xlSheet.Cells(j, 16) & "',ValidFrom='" & xlSheet.Cells(j, 17) & "',PrintSV='" & xlSheet.Cells(j, 18) & "',Remark='" & xlSheet.Cells(j, 19) & "'" & _
                        " where SN='" & xlSheet.Cells(j, 3) & "' and HV='" & xlSheet.Cells(j, 2) & "' "
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "修改H3C设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "修改失败"
                   End If
'                   MsgBox "修改H3C设定资料成功!"
               ElseIf isexist = False Then
                   sql = "Insert into tblH3C(ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark) " & _
                        " Values(" & getmaxID("tblH3C") & ",'" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "','" & xlSheet.Cells(j, 4) & "','" & xlSheet.Cells(j, 5) & "','" & xlSheet.Cells(j, 6) & "','" & xlSheet.Cells(j, 7) & "','" & xlSheet.Cells(j, 8) & "','" & xlSheet.Cells(j, 9) & "','" & xlSheet.Cells(j, 10) & "','" & xlSheet.Cells(j, 11) & "'," & _
                        "'" & xlSheet.Cells(j, 12) & "','" & xlSheet.Cells(j, 13) & "','" & xlSheet.Cells(j, 14) & "','" & xlSheet.Cells(j, 15) & "','" & xlSheet.Cells(j, 16) & "','" & xlSheet.Cells(j, 17) & "','" & xlSheet.Cells(j, 18) & "','" & xlSheet.Cells(j, 19) & "')"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "新增H3C设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "修改失败"
                   End If
'                   MsgBox "新增H3C设定资料成功!"
               End If
           End If
         Next K
         
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "H3C设定资料导入成功!"
      renovate ("")
End Sub

Private Sub cmdInsert_Click()
    op = "Insert"
    Reset

'    enable
'    txtSN.Text = ""
'    txtCPN.Text = ""
'    txtEPN.Text = ""
'    txtDes.Text = ""
'    txtSize.Text = ""
'    txtGW.Text = ""
'
'    chkCE.Value = 1
'    chkWEEE.Value = 1
'    chkChinaRoHS.Value = 1
'    chkRoHS.Value = 1
'    chkTurkeyRohs.Value = 1
'    chkSVPrint.Value = 1
'
'    txtNAL2.Text = "/"
'    txtNAL2Title = "/"
'    txtMS.Text = "N/A"
'
'
'    txtHV.Text = "N/A"
'    txtRemark.Text = "N/A"

End Sub

Private Sub cmdQuery_Click()
    If txtSN.Enabled = False Then
      MsgBox "请按新增按钮清空就可输入查询内容!", vbOKOnly + vbInformation, "输入查询内容"
    End If
    If rec.State = 1 Then
        rec.Close
     End If
       sql = " SELECT [ID],[Part_Number],[Part_Desc],Part_ENDesc,[Part_Model],[Rev],[ExeStandard],[ACSign],[DCSign],[Weight],[CustomerCode],[Size]  " & _
        " ,case when RoHS is null then 'N/A' when RoHS = 0 then 'No' when RoHS = 1 then 'Yes' end as 'RoHS'" & _
        " ,case when WEEE is null then 'N/A' when WEEE = 0 then 'No' when WEEE = 1 then 'Yes' end as 'WEEE' " & _
        " ,case when CCC is null then 'N/A' when CCC = 0 then 'No' when CCC = 1 then 'Yes' end as 'CCC' " & _
        " FROM [Print].[dbo].[tblDaHuaNew] where 1 = 1"
     
     If txtSN.Text <> "" Then
        sql = sql & " and Part_Number = '" & txtSN.Text & "'"
     End If

     sql = sql & " order by Part_Number "
    renovate (sql)
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

'Private Sub cmdSelect_Click()
'   On Error Resume Next
''   cdSelect.CancelError = True
'   cdSelect.Filter = "*.xls|*.xls"
'   cdSelect.action = 1
'   If cdSelect.FileName <> "" Then txtPath.Text = cdSelect.FileName
'End Sub

Private Sub cmdUpdate_Click()
   If mfgH3C.RowSel <= 0 Then
      MsgBox "请选择要修改的行!", vbInformation + vbOKOnly, "未选择行"
      Exit Sub
   End If
   mfgH3C_Click
   op = "Update"
   Reset
End Sub

Private Sub Form_Load()
'   unable
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   
   renovate ("")
End Sub

Private Sub renovate(sql As String)
    Set mfgH3C.DataSource = Nothing
    If sql = "" Then
       sql = " SELECT [ID],[Part_Number] ,[Part_Desc],Part_ENDesc ,[Part_Model],[Rev],[ExeStandard],[ACSign],[DCSign],[Weight],[CustomerCode],[Size]  " & _
        " ,case when RoHS is null then 'N/A' when RoHS = 0 then 'No' when RoHS = 1 then 'Yes' end as 'RoHS'" & _
        " ,case when WEEE is null then 'N/A' when WEEE = 0 then 'No' when WEEE = 1 then 'Yes' end as 'WEEE' " & _
        " ,case when CCC is null then 'N/A' when CCC = 0 then 'No' when CCC = 1 then 'Yes' end as 'CCC' " & _
        " FROM [Print].[dbo].[tblDaHuaNew] where 1 = 1"
    End If
    Debug.Print sql
    
    If rec.State = 1 Then
    rec.Close
    End If
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    Set mfgH3C.DataSource = rec
    With mfgH3C
      .Cols = rec.Fields.Count + 1
      .ColWidth(0) = 400
      .ColWidth(1) = 400
      .ColWidth(2) = 2000
      .ColWidth(3) = 2000
      .ColWidth(4) = 2000
      .ColWidth(5) = 1000
      .ColWidth(6) = 3000
      .ColWidth(7) = 3000
      .ColWidth(8) = 3000
      .ColWidth(9) = 1500
      .ColWidth(10) = 2000
      .ColWidth(11) = 1000
      .ColWidth(12) = 1000
      .ColWidth(13) = 1000
      .ColWidth(14) = 400

      
      .TextMatrix(0, 1) = "ID"
      .TextMatrix(0, 2) = "产品编码"
      .TextMatrix(0, 3) = "产品类别"
      .TextMatrix(0, 4) = "英文描述"
      .TextMatrix(0, 5) = "产品型号"
      .TextMatrix(0, 6) = "整机版本"
      .TextMatrix(0, 7) = "执行标准"
      .TextMatrix(0, 8) = "交流符号"
      .TextMatrix(0, 9) = "直流符号"
      .TextMatrix(0, 10) = "重量"
      .TextMatrix(0, 11) = "客户编码"
      .TextMatrix(0, 12) = "尺寸"
      .TextMatrix(0, 13) = "RoHS"
      .TextMatrix(0, 14) = "WEEE"
      .TextMatrix(0, 15) = "CCC"

    End With
    rec.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rec.State = 1 Then
      rec.Close
      Set conn = Nothing
   End If
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub

Private Sub mfgH3C_Click()

   If mfgH3C.RowSel > 0 Then
'      txtHV.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 3)
      txtSN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 2)
      txtDesc.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 3)
      txtEN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 4)
      txtModel.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 5)
      txtRev.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 6)
      txtExeStandard.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 7)
      txtACSign.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 8)
      txtDCSign.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 9)
      txtWeight.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 10)
      txtCustomerCode.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 11)
      txtSize.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 12)
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 13))) = "YES" Then
        chkRoHS.Value = 1
      ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 13)) = "NO" Then
        chkNonRoHS.Value = 1
      End If
     
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 14))) = "YES" Then
        chkWEEE.Value = 1
      ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 14)) = "NO" Then
        chkNonWEEE.Value = 1
      End If
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 15))) = "YES" Then
        chkCCC.Value = 1
      ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 15)) = "NO" Then
        chkNonCCC.Value = 1
      End If
      
   End If
End Sub

Private Sub mfgH3C_SelChange()
   mfgH3C_Click
End Sub


