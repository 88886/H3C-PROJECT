VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmH3C_3COMSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H3C-3COM Setting"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmH3C_3COMSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   12120
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog cdSelect 
      Left            =   2760
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   735
      Left            =   10740
      TabIndex        =   19
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   735
      Left            =   9300
      TabIndex        =   18
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "确定(Confirm)"
      Height          =   735
      Left            =   7620
      TabIndex        =   17
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(Delete)"
      Height          =   735
      Left            =   10740
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修改(Update)"
      Height          =   735
      Left            =   9300
      TabIndex        =   15
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "新增(Insert)"
      Height          =   735
      Left            =   7620
      TabIndex        =   14
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查询(Query)"
      Height          =   735
      Left            =   5940
      TabIndex        =   13
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "导出(Export)"
      Height          =   735
      Left            =   3420
      TabIndex        =   12
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "导入(Import)"
      Height          =   735
      Left            =   3420
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择(Select)"
      Height          =   495
      Left            =   1380
      TabIndex        =   10
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   180
      TabIndex        =   9
      Top             =   6240
      Width           =   3015
   End
   Begin VB.Frame fmH3C 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.CheckBox chkNo 
         Caption         =   "否"
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox chkYES 
         Caption         =   "是"
         Height          =   375
         Left            =   2280
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton opt21 
         Caption         =   "21类"
         Height          =   375
         Left            =   9720
         TabIndex        =   22
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton opt3C 
         Caption         =   "3C类"
         Height          =   375
         Left            =   8400
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtScode 
         Height          =   450
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtModel 
         Height          =   450
         Left            =   8400
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtPart 
         Height          =   450
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "是否打印IP地址:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "类别:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   20
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblEPN 
         Caption         =   "S/N开头:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblCPN 
         Caption         =   "3COM编号:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblSN 
         Caption         =   "产品机种:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgH3C_3COM 
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   6376
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblPath 
      Caption         =   "导入/导出路径:"
      Height          =   375
      Left            =   180
      TabIndex        =   8
      Top             =   5760
      Width           =   2175
   End
End
Attribute VB_Name = "frmH3C_3COMSetting"
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

Private Sub enable()
   txtPart.Enabled = True
   txtPart.BackColor = &HFFFFFF
   txtModel.Enabled = True
   txtModel.BackColor = &HFFFFFF
   txtScode.Enabled = True
   txtScode.BackColor = &HFFFFFF
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = False
   cmdUpdate.Enabled = False
   cmdDelete.Enabled = False
   cmdConfirm.Enabled = True
   cmdCancel.Enabled = True
   chkYES.Enabled = True
   chkNo.Enabled = True
   opt3C.Enabled = True
   opt21.Enabled = True
End Sub

Private Sub unable()
   txtPart.Enabled = False
   txtPart.BackColor = &HE0E0E0
   txtModel.Enabled = False
   txtModel.BackColor = &HE0E0E0
   txtScode.Enabled = False
   txtScode.BackColor = &HE0E0E0
  
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = True
   cmdUpdate.Enabled = True
   cmdDelete.Enabled = True
   cmdConfirm.Enabled = False
   cmdCancel.Enabled = False
   
   chkYES.Enabled = False
   chkNo.Enabled = False
   opt3C.Enabled = False
   opt21.Enabled = False
End Sub

Private Sub chkNo_Click()
 If chkNo.Value = 1 Then
      chkYES.Value = 0
   Else
      chkYES.Value = 1
   End If
End Sub

Private Sub chkYES_Click()
 If chkYES.Value = 1 Then
      chkNo.Value = 0
   Else
      chkNo.Value = 1
   End If
End Sub


Private Sub cmdCancel_Click()
   unable
   op = ""
End Sub

Private Sub cmdConfirm_Click()
   If txtPart.Text = "" Then
      MsgBox "产品编码不能为空!!", vbExclamation + vbOKOnly, "产品编码空"
      txtPart.SetFocus
      Exit Sub
   End If
   If txtModel.Text = "" Then
       MsgBox "产品型号不能为空!", vbExclamation + vbOKOnly, "产品型号空"
       txtModel.SetFocus
       Exit Sub
   End If
   If txtScode.Text = "" Then
      MsgBox "产品Species Code不能为空!", vbExclamation + vbOKOnly, "产品Species Code空"
      txtScode.SetFocus
      Exit Sub
   End If
   
    If chkYES.Value = 0 And chkNo.Value = 0 Then
        MsgBox "是否打印IP地址不可为空！"
        chkYES.SetFocus
        Exit Sub
    End If
    
    If opt3C.Value = 0 And opt21.Value = 0 Then
        MsgBox "类别不可为空！"
        opt3C.SetFocus
        Exit Sub
    End If
    
   
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from H3COM where part='" & txtPart.Text & "'"
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "产品编码已存在!", vbExclamation + vbOKOnly, "产品编号重复"
         txtPart.SetFocus
         Exit Sub
      End If
      rcd.Close
      sql = "Insert into H3COM(ID,Part,Class,Model,Scode,IFPrintIP) " & _
            "Values(" & getmaxID("H3COM") & ",'" & Replace(Trim(txtPart.Text), Chr(13) & Chr(10), "") & "','" & IIf(opt3C.Value = 1, "3C", "21") & "','" & txtModel.Text & "','" & txtScode.Text & "','" & IIf(chkYES.Value = 1, "Yes", "No") & "')"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "新增H3C_3COM设定资料失败!" & "原因是" & status, vbOKOnly + vbInformation, "新增失败"
      End If
      MsgBox "新增H3C_3COM设定资料成功!", vbOKOnly + vbInformation, "新增成功"
      renovate
      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update H3COM set Model='" & txtModel.Text & "',Scode='" & txtScode.Text & "',Class='" & IIf(opt3C.Value = True, "3C", "21") & "',IFPrintIP='" & IIf(chkYES.Value = 1, "Yes", "No") & "'" & _
            " where ID=" & mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 1) & " and part='" & txtPart.Text & "'"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "修改H3C_3COM设定资料失败!" & "原因是" & status, vbOKOnly + vbInformation, "修改失败"
      End If
      MsgBox "修改H3C_3COM设定资料成功!", vbOKOnly + vbInformation, "修改成功"
      renovate
      cmdCancel_Click
   End If
   renovate
End Sub

Private Sub cmdDelete_Click()
   If mfgH3C_3COM.RowSel <= 0 Then
      MsgBox "请选择要删除的行!", vbInformation + vbOKOnly, "未选择行"
      Exit Sub
   End If
   sql = "delete from H3COM where ID=" & mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 1) & " and part='" & mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 2) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "删除H3C_3COM设定资料失败!" & "原因是" & status, vbInformation + vbOKOnly, "删除失败"
   End If
   MsgBox "删除H3C_3COM设定资料成功!", vbInformation + vbOKOnly, "删除成功"
   renovate
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If mfgH3C_3COM.Rows = 0 Then
      MsgBox "无资料可汇出"
      Exit Sub
   End If
   If txtPath.Text <> "" Then
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets.Item(1)
       For i = 0 To mfgH3C_3COM.Rows - 1
         For j = 1 To mfgH3C_3COM.Cols - 1
          xlSheet.Cells(i + 1, j) = mfgH3C_3COM.TextMatrix(i, j)
       Next j
      Next i
      xlBook.SaveAs (txtPath.Text)
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "汇出到EXCEL资料成功!!"
    End If
End Sub

Private Sub cmdImport_Click()
   If txtPath.Text = "" Then
      MsgBox "导入路径不能为空!", vbOKOnly + vbExclamation, "导入路径空"
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
          Dim isexist As Boolean
          If xlSheet.Cells(j, 3) = "" Then
             MsgBox "导入资料格式不正确!", vbOKOnly + vbExclamation, "导入格式错误"
             Exit Sub
          End If
          isexist = False
          For K = 1 To 3
           If K = 3 Then
             cellValue = xlSheet.Cells(j, K)
             If cellValue = "" Then
                MsgBox "导入资料格式不正确!", vbOKOnly + vbExclamation, "导入格式错误"
                Exit Sub
             End If
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from H3COM where Part='" & cellValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("产品编码已存在!", vbAbortRetryIgnore + vbExclamation, "资料重复")
                End If
                
                If action = vbAbort Then
                   MsgBox "资料导入已终止!!", vbOKOnly + vbInformation, "导入终止"
                   rcd.Close
                   Exit Sub
                ElseIf action = vbIgnore And info = True Then
                   MsgBox "重复产品编号资料不会导入,请稍等..!!", vbInformation + vbOKOnly, "重复的不导入"
                   rcd.Close
                   info = False
                   Exit For
                ElseIf action = vbRetry And info = True Then
                   MsgBox "重复产品编号资料会自动更新,请稍等..!!", vbInformation + vbOKOnly, "重复的会自动更新"
                   info = False
                End If
                isexist = True
             Else
                isexist = False
             End If
             rcd.Close
            End If
            
            If K = 3 Then
               If action = vbRetry Then
                   sql = "Update H3COM set Model='" & xlSheet.Cells(j, 3) & "',Scode='" & xlSheet.Cells(j, 4) & "'," & _
                         " where Part='" & xlSheet.Cells(j, 2) & "'"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "修改H3C_3COM设定资料失败!" & "原因是" & status, vbOKOnly + vbExclamation, "修改失败"
                   End If
'                   MsgBox "修改H3C_3COM_3COM设定资料成功!"
               ElseIf isexist = False Then
                   sql = "Insert into H3COM(ID,Part,Model,Scode) " & _
                        "Values(" & getmaxID("H3COM") & ",'" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "','" & xlSheet.Cells(j, 4) & "')"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "新增H3C_3COM设定资料失败!" & "原因是" & status, vbOKOnly + vbInformation, "新增失败"
                   End If
'                   MsgBox "新增H3C_3COM_3COM设定资料成功!"
               End If
           End If
         Next K
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "H3C_3COM_3COM设定资料导入成功!", vbOKOnly + vbInformation, "导入成功"
      renovate
End Sub

Private Sub cmdInsert_Click()
   enable
   txtPart.Text = ""
   txtModel.Text = ""
   txtScode.Text = ""
   op = "Insert"
End Sub

Private Sub cmdQuery_Click()
   If txtPart.Enabled = False Then
      MsgBox "请按新增按钮清空就可输入查询内容!", vbOKOnly + vbInformation, "输入查询内容"
   End If
   If rec.State = 1 Then
      rec.Close
   End If
   sql = "select * from H3COM Where 1=1"
   If txtPart.Text <> "" Then
      sql = sql & " and part like '%" & txtPart.Text & "%'"
   End If
   If txtModel.Text <> "" Then
      sql = sql & " and model like '%" & txtModel.Text & "%'"
   End If
   If txtScode.Text <> "" Then
      sql = sql & " and scode like '%" & txtScode.Text & "%'"
   End If
   sql = sql & " Order by ID,part"
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgH3C_3COM.DataSource = rec
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdSelect_Click()
   cdSelect.CancelError = True
   cdSelect.Filter = "*.xls|*.xls"
   cdSelect.action = 1
   If cdSelect.FileName <> "" Then txtPath.Text = cdSelect.FileName
End Sub

Private Sub cmdUpdate_Click()
   If mfgH3C_3COM.RowSel <= 0 Then
      MsgBox "请选择要修改的行!", vbInformation + vbOKOnly, "未选择行"
      Exit Sub
   End If
   mfgH3C_3COM_Click
   enable
   txtPart.Enabled = False
   txtPart.BackColor = &HE0E0E0
   op = "Update"
End Sub




Private Sub Form_Load()
   unable
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   renovate
End Sub

Private Sub renovate()
   sql = "select * from H3COM order by ID,Part"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgH3C_3COM.DataSource = rec
   With mfgH3C_3COM
        .Cols = rec.Fields.Count + 1

        .ColWidth(0) = 400
        .ColWidth(1) = 1000
        .ColWidth(2) = 2500
        .ColWidth(3) = 1000
        .ColWidth(4) = 3000
        .ColWidth(5) = 3000
        .ColWidth(6) = 1000
        
        .TextMatrix(0, 1) = "序号(ID)"
        .TextMatrix(0, 2) = "产品机种"
        .TextMatrix(0, 3) = "类别"
        .TextMatrix(0, 4) = "3COM编号"
        .TextMatrix(0, 5) = "S/N开头"
        .TextMatrix(0, 6) = "PrintIP"
       
   End With
   rec.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rec.State = 1 Then
      rec.Close
      Set rec = Nothing
   End If
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub



Private Sub mfgH3C_3COM_Click()
   If mfgH3C_3COM.RowSel > 0 Then
      txtPart.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 2)
      If mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 3) = "3C" Then
        opt3C.Value = True
      Else
       opt21.Value = True
      End If
      txtModel.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 4)
      txtScode.Text = mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 5)
      If mfgH3C_3COM.TextMatrix(mfgH3C_3COM.RowSel, 6) = "Yes" Then
        chkYES.Value = 1
      Else
       chkNo.Value = 1
      End If
      
   End If
End Sub

Private Sub mfgH3C_3COM_SelChange()
   mfgH3C_3COM_Click
End Sub


