VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmOthers 
   Caption         =   "NEC 3COM 电源设定"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16380
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   16380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRemark 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   21
      Top             =   1920
      Width           =   2055
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
      Left            =   9720
      TabIndex        =   19
      Top             =   6360
      Width           =   1455
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
      Left            =   11400
      TabIndex        =   18
      Top             =   5880
      Width           =   1335
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
      Left            =   13080
      TabIndex        =   17
      Top             =   5880
      Width           =   1215
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
      Left            =   14520
      TabIndex        =   16
      Top             =   5880
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
      Left            =   11400
      TabIndex        =   15
      Top             =   6720
      Width           =   1335
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
      Left            =   13080
      TabIndex        =   14
      Top             =   6720
      Width           =   1215
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
      Left            =   14520
      TabIndex        =   13
      Top             =   6720
      Width           =   1215
   End
   Begin VB.ComboBox cb5000 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmOthers.frx":0000
      Left            =   2160
      List            =   "frmOthers.frx":0010
      TabIndex        =   11
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtHV 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtSN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
   Begin VB.CheckBox chkSVPrint 
      Caption         =   "是"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.CheckBox chkNonPrintSV 
      Caption         =   "否"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CheckBox chkNonPrintPC 
      Caption         =   "否"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.CheckBox chkPrintPC 
      Caption         =   "是"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgOther 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   4895
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
      OLEDropMode     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "Remark"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblHV 
      Caption         =   "硬件版本:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblSN 
      Caption         =   "产品编码:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblPrintSV 
      Caption         =   "打印软件版本:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "5000状态:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "打印电源代码:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim op, sql As String


Private Sub renovate(sql As String)
    Set mfgOther.DataSource = Nothing
    If sql = "" Then
        sql = "SELECT [ID],[Part_Number],[Part_Revision]" & _
        ",case when Print_SV is null then 'N/A' when Print_SV = 0 then 'No' when Print_SV = 1 then 'Yes' end as 'Print_SV'" & _
        ",case when Print_Power is null then 'N/A' when Print_Power = 0 then 'No' when Print_Power = 1 then 'Yes' end as 'Print_Power'" & _
        ",[5000_State],[Remark] FROM [Print].[dbo].[tblOthers] order by Part_Number,Part_Revision"
    End If
    If rec.State = 1 Then
    rec.Close
    End If
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    Set mfgOther.DataSource = rec
    With mfgOther
      .Cols = rec.Fields.Count + 1
      .ColWidth(0) = 400
      .ColWidth(1) = 650
      .ColWidth(2) = 1300
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .TextMatrix(0, 1) = "ID"
      .TextMatrix(0, 2) = "产品编码"
      .TextMatrix(0, 3) = "硬件版本"
      .TextMatrix(0, 4) = "打印软件版本"
      .TextMatrix(0, 5) = "打印电源代码"
      .TextMatrix(0, 6) = "5000状态"
      .TextMatrix(0, 7) = "备注"
    End With
    rec.Close
End Sub
Public Sub Reset()
      For Each ctr In Me.Controls
        If TypeOf ctr Is TextBox Then
'                ctr.Enabled = False
                ctr.BackColor = &HFFFFFF
             ElseIf TypeOf ctr Is ComboBox Then
'                ctr.Enabled = False
                ctr.BackColor = &HFFFFFF
             ElseIf TypeOf ctr Is CheckBox Then
                ctr.Enabled = True
        End If
    Next
End Sub

Private Sub cmdCancel_Click()
    Reset
    cmdQuery.Enabled = True
    cmdInsert.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdConfirm.Enabled = True
    cmdCancel.Enabled = False
End Sub

Private Sub cmdConfirm_Click()
    If Trim(txtSN.Text) = "" Then
      MsgBox "产品编码不能为空!!", vbExclamation + vbOKOnly, "产品编码空"
      txtSN.SetFocus
      Exit Sub
   End If

   Dim PWPrint As String, ftStatus, status, sql As String


   If chkSVPrint.Value = 1 Then
      SVPrint = "1"
   Else
      SVPrint = "0"
   End If

    If Me.chkPrintPC.Value = 1 Then
        PWPrint = "1"
    Else
        PWPrint = "0"
    End If

    If Me.cb5000.ListIndex <= -1 Then
        MsgBox "5000状态没有选择!", vbExclamation + vbOKOnly, "请选择5000状态的一个选项"
         Me.cb5000.SetFocus
         Exit Sub
    End If

'    Y，N，NA,TBD
    If Me.cb5000.ListIndex = 0 Then
        ftStatus = "Y"
    ElseIf Me.cb5000.ListIndex = 1 Then
        ftStatus = "N"
    ElseIf Me.cb5000.ListIndex = 2 Then
        ftStatus = "NA"
    ElseIf Me.cb5000.ListIndex = 3 Then
        ftStatus = "TBD"
    End If
'
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from tblOthers where Part_Number ='" & Trim(txtSN.Text) & "' and Part_Revision ='" & Trim(txtHV.Text) & "' "
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "产品编码&版本已存在!", vbExclamation + vbOKOnly, "产品编号重复"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close

      sql = "insert [tblOthers]([Part_Number],[Part_Revision],[Print_SV],[Print_Power],[5000_State],[Remark]) " & _
            "Values('" & Trim(txtSN.Text) & "','" & Trim(txtHV.Text) & "'," & SVPrint & "," & PWPrint & ",'" & ftStatus & "','" & txtRemark.Text & "')"


      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "新增H3C设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "新增失败"
      Else
        MsgBox "新增H3C设定资料成功!", vbInformation + vbOKOnly, "新增成功"
      End If
      renovate ("")
'      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update tblOthers set Print_Power = " & PWPrint & ",Part_Revision ='" & txtHV.Text & "',Print_SV='" & SVPrint & "',[5000_State] = '" & ftStatus & "',Remark='" & txtRemark.Text & "'" & _
            " where ID=" & mfgOther.TextMatrix(mfgOther.RowSel, 1) & " and Part_Number ='" & Trim(txtSN.Text) & "'"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "修改H3C设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "修改失败"
      Else
         MsgBox "修改H3C设定资料成功!", vbInformation + vbOKOnly, "修改成功"
      End If
      renovate ("")
'      cmdCancel_Click
   End If
   renovate ("")
End Sub

Private Sub cmdDelete_Click()
    If mfgOther.RowSel <= 0 Then
      MsgBox "请选择要删除的行!", vbInformation + vbOKOnly, "未选择行"
      Exit Sub
   End If
   sql = "delete from tblOthers where ID=" & mfgOther.TextMatrix(mfgOther.RowSel, 1) & " and Part_Number ='" & mfgOther.TextMatrix(mfgOther.RowSel, 2) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "删除NEC & 3COM 设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "删除失败"
   End If
   MsgBox "删除H3C设定资料成功!", vbInformation + vbOKOnly, "删除成功"
   renovate ("")
End Sub

Private Sub cmdInsert_Click()
    op = "Insert"
    Reset
''    Me.txtSN.SetFocus
End Sub

Private Sub cmdQuery_Click()
    If txtSN.Enabled = False Then
      MsgBox "请按新增按钮清空就可输入查询内容!", vbOKOnly + vbInformation, "输入查询内容"
    End If
    If rec.State = 1 Then
        rec.Close
     End If
       sql = "SELECT [ID],[Part_Number],[Part_Revision]" & _
        ",case when Print_SV is null then 'N/A' when Print_SV = 0 then 'No' when Print_SV = 1 then 'Yes' end as 'Print_SV'" & _
        ",case when Print_Power is null then 'N/A' when Print_Power = 0 then 'No' when Print_Power = 1 then 'Yes' end as 'Print_Power'" & _
        ",[5000_State],[Remark] FROM [Print].[dbo].[tblOthers] where 1 = 1"
     
     If txtSN.Text <> "" Then
        sql = sql & " and Part_Number like '%" & txtSN.Text & "%'"
     End If
    sql = sql & " order by Part_Number,Part_Revision"
    renovate (sql)

End Sub

Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    op = "Update"
    Reset
End Sub

Private Sub Form_Load()
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   renovate ("")
End Sub

Private Sub mfgOther_Click()
   If mfgOther.RowSel > 0 Then
      txtHV.Text = mfgOther.TextMatrix(mfgOther.RowSel, 3)
      txtSN.Text = mfgOther.TextMatrix(mfgOther.RowSel, 2)
    If UCase(Trim(mfgOther.TextMatrix(mfgOther.RowSel, 4))) = "YES" Then
        Me.chkSVPrint.Value = 1
    ElseIf UCase(Trim(mfgOther.TextMatrix(mfgOther.RowSel, 4))) = "NO" Then
        Me.chkNonPrintSV.Value = 1
    End If

    If UCase(Trim(mfgOther.TextMatrix(mfgOther.RowSel, 5))) = "YES" Then
        Me.chkPrintPC.Value = 1
    ElseIf UCase(Trim(mfgOther.TextMatrix(mfgOther.RowSel, 5))) = "NO" Then
        Me.chkNonPrintPC.Value = 1
    End If

    If Trim(mfgOther.TextMatrix(mfgOther.RowSel, 6)) = "Y" Then
        cb5000.ListIndex = 0
    ElseIf Trim(mfgOther.TextMatrix(mfgOther.RowSel, 6)) = "N" Then
        cb5000.ListIndex = 1
    ElseIf Trim(mfgOther.TextMatrix(mfgOther.RowSel, 6)) = "NA" Then
        cb5000.ListIndex = 2
    ElseIf Trim(mfgOther.TextMatrix(mfgOther.RowSel, 6)) = "TBD" Then
        cb5000.ListIndex = 3
    End If

    txtRemark.Text = mfgOther.TextMatrix(mfgOther.RowSel, 7)
    End If

End Sub
