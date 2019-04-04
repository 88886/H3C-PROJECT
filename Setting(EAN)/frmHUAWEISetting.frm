VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmHUAWEISetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "非3COM类 Setting"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHUAWEISetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   12135
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog cdSelect 
      Left            =   2400
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10620
      TabIndex        =   17
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9180
      TabIndex        =   16
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "确定(Confirm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7500
      TabIndex        =   15
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(Delete)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10620
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修改(Update)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9180
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "新增(Insert)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7500
      TabIndex        =   12
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查询(Query)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5820
      TabIndex        =   11
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "导出(Export)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3300
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "导入(Import)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3300
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择(Select)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1260
      TabIndex        =   8
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   7
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Frame fmH3C 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.Frame framePB 
         Caption         =   "有无铅"
         Height          =   855
         Left            =   8640
         TabIndex        =   18
         Top             =   0
         Width           =   3015
         Begin VB.OptionButton opNonPB 
            Caption         =   "无铅"
            Height          =   585
            Left            =   1560
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opPB 
            Caption         =   "有铅"
            Height          =   345
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox txtSN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtXH 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5400
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblSN 
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCPN 
         Caption         =   "产品型号:"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgHUAWEI 
      Height          =   2895
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5106
      _Version        =   393216
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   4200
      Width           =   2175
   End
End
Attribute VB_Name = "frmHUAWEISetting"
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
   txtSN.Enabled = True
   txtSN.BackColor = &HFFFFFF
   txtXH.Enabled = True
   txtXH.BackColor = &HFFFFFF
  
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

Private Sub unable()
   txtSN.Enabled = False
   txtSN.BackColor = &HE0E0E0
   txtXH.Enabled = False
   txtXH.BackColor = &HE0E0E0
  
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = True
   cmdUpdate.Enabled = True
   cmdDelete.Enabled = True
   cmdConfirm.Enabled = False
   cmdCancel.Enabled = False
End Sub


Private Sub cmdCancel_Click()
   unable
   op = ""
End Sub

Private Sub cmdConfirm_Click()
   Dim result As String
   If txtSN.Text = "" Then
      MsgBox "产品编码不能为空!!", vbExclamation + vbOKOnly, "产品编码空"
      txtSN.SetFocus
      Exit Sub
   End If
   If txtXH.Text = "" Then
       MsgBox "产品型号不能为空!", vbExclamation + vbOKOnly, "产品型号空"
       txtXH.SetFocus
       Exit Sub
   End If
   If Me.opNonPB.Value = True Then
     result = "0"
   Else
    If Me.opPB.Value = True Then
        result = "1"
    Else
'        result = "0"
       MsgBox "铅属性不能为空!", vbExclamation + vbOKOnly, "铅属性空"
       Exit Sub
    End If
    
   End If
   
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from SingleUnit where SN='" & txtSN.Text & "'"
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "产品编码已存在!"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close
      sql = "Insert into SingleUnit(ID,SN,type,PB) " & _
            "Values(" & getmaxID("SingleUnit") & ",'" & txtSN.Text & "','" & txtXH.Text & "'," & result & ")"
            
      sql = sql & " insert into SingleUnit_log(CREATE_USER,SN,TYPE,PB,COMMENT) "
      sql = sql & " Values( '" & golUSERNAME & "','" & txtSN.Text & "','" & txtXH.Text & "'," & result & ",'Insert')"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "新增资料失败!" & "原因是" & status
      End If
      MsgBox "新增资料成功!"
      renovate
      cmdInsert_Click
   ElseIf op = "Update" Then
   
        '' edit on 2015/01/30
        If ConfirmConfig(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 3), txtXH.Text, IIf(mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 4) = "Yes", "有铅", "无铅"), IIf(opNonPB.Value = True, "无铅", "有铅")) Then
            '' do nothing
        Else
            Exit Sub
        End If
        
      sql = "Update SingleUnit set type='" & txtXH.Text & "',PB = " & result & " where ID=" & mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 1) & " and SN='" & txtSN.Text & "'"
      
      sql = sql & " insert into SingleUnit_log(CREATE_USER,SN,TYPE,PB,COMMENT) "
      sql = sql & " Values( '" & golUSERNAME & "','" & txtSN.Text & "','" & txtXH.Text & "'," & result & ",'Update')"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "修改资料失败!" & "原因是" & status
      End If
      MsgBox "修改资料成功!"
      renovate
      cmdCancel_Click
   End If
   renovate
End Sub

Private Sub cmdDelete_Click()
   If mfgHUAWEI.RowSel <= 0 Then
      MsgBox "请选择要删除的行!"
      Exit Sub
   End If
   'sql = "delete from SingleUnit where ID=" & mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 1) & " and SN='" & mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 2) & "'"
   
   sql = " insert into SingleUnit_log(CREATE_USER,SN,TYPE,PB,COMMENT) "
   sql = sql & " select '" & golUSERNAME & "',SN,TYPE,PB,'delete' from SingleUnit where  SN='" & mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 2) & "'"
   sql = sql & "delete from SingleUnit where  SN='" & mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 2) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "删除资料失败!" & "原因是" & status
   End If
   MsgBox "删除资料成功!"
   renovate
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If mfgHUAWEI.Rows = 0 Then
      MsgBox "无资料可汇出"
      Exit Sub
   End If
   If txtPath.Text <> "" Then
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets.Item(1)
       For i = 0 To mfgHUAWEI.Rows - 1
         For j = 1 To mfgHUAWEI.Cols - 1
          xlSheet.Cells(i + 1, j) = mfgHUAWEI.TextMatrix(i, j)
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
      MsgBox "导入路径不能为空!"
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
          If xlSheet.Cells(j, 2) = "" Then
             MsgBox "导入资料格式不正确!"
             Exit Sub
          End If
          isexist = False
          For K = 1 To 2
           If K = 2 Then
             cellValue = xlSheet.Cells(j, K)
             If cellValue = "" Then
                MsgBox "导入资料格式不正确!"
                Exit Sub
             End If
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from SingleUnit where SN='" & cellValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("产品编码已存在!", vbAbortRetryIgnore + vbExclamation, "资料重复")
                End If

                If action = vbAbort Then
                   MsgBox "资料导入已终止!!"
                   rcd.Close
                   Exit Sub
                ElseIf action = vbIgnore And info = True Then
                   MsgBox "重复产品编号资料不会导入,请稍等..!!"
                   rcd.Close
                   info = False
                   Exit For
                ElseIf action = vbRetry And info = True Then
                   MsgBox "重复产品编号资料会自动更新,请稍等..!!"
                   info = False
                End If
                isexist = True
             Else
                isexist = False
             End If
             rcd.Close
            End If

            If K = 2 Then
               If action = vbRetry Then
                   sql = "Update SingleUnit set type='" & xlSheet.Cells(j, 3) & "'" & _
                        " where SN='" & xlSheet.Cells(j, 2) & "'"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "修改资料失败!" & "原因是" & status
                   End If
'                   MsgBox "修改HUAWEI设定资料成功!"
               ElseIf isexist = False Then
                   sql = "Insert into SingleUnit(ID,SN,type) " & _
                        "Values(" & getmaxID("SingleUnit") & ",'" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "')"

                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "新增资料失败!" & "原因是" & status
                   End If
'                   MsgBox "新增HUAWEI设定资料成功!"
               End If
           End If
         Next K
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "资料导入成功!"
      renovate
End Sub

Private Sub cmdInsert_Click()
   enable
   txtSN.Text = ""
   txtXH.Text = ""
   Me.opNonPB.Value = False
   Me.opPB.Value = False
   op = "Insert"
End Sub

Private Sub cmdQuery_Click()
   MsgBox "请按新增按钮清空就可输入查询内容!", vbOKOnly + vbInformation, "输入查询内容"
   If rec.State = 1 Then
      rec.Close
   End If
   sql = "select ID,SN,TYPE,CASE PB WHEN 1 THEN 'Yes' when 0 then 'No' else 'Non'end from SingleUnit Where 1=1"
   If txtSN.Text <> "" Then
      sql = sql & " and SN like '%" & txtSN.Text & "%'"
   End If
   If txtXH.Text <> "" Then
      sql = sql & " and type like '%" & txtXH.Text & "%'"
   End If
     sql = sql & " order by ID,SN"
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgHUAWEI.DataSource = rec
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdSelect_Click()
   On Error Resume Next
   cdSelect.CancelError = True
   cdSelect.Filter = "*.xls|*.xls"
   cdSelect.action = 1
   If cdSelect.Filename <> "" Then txtPath.Text = cdSelect.Filename
End Sub

Private Sub cmdUpdate_Click()
   If mfgHUAWEI.RowSel <= 0 Then
      MsgBox "请选择要修改的行!"
      Exit Sub
   End If
   mfgHUAWEI_Click
   enable
   txtSN.Enabled = False
   txtSN.BackColor = &HE0E0E0
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
   sql = "select ID,SN,TYPE,CASE PB WHEN 1 THEN 'Yes' when 0 then 'No' else 'Non' end from SingleUnit order by ID,SN"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgHUAWEI.DataSource = rec
   With mfgHUAWEI
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 3000
        .ColWidth(4) = 3000
        .TextMatrix(0, 1) = "序号(ID)"
        .TextMatrix(0, 2) = "产品编码(Model Number)"
        .TextMatrix(0, 3) = "产品型号(Product Type)"
        .TextMatrix(0, 4) = "有无铅设定(PB)"
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

Private Sub mfgHUAWEI_Click()
   If mfgHUAWEI.RowSel > 0 Then
      txtSN.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 2)
      txtXH.Text = mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 3)
      If (mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 4) = "Yes") Then
        Me.opPB.Value = 1
      ElseIf (mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 4) = "No") Then
            Me.opNonPB.Value = 1
      ElseIf (mfgHUAWEI.TextMatrix(mfgHUAWEI.RowSel, 4) = "Non") Then
            Me.opNonPB.Value = 0
            Me.opPB.Value = 0
      End If
  End If
End Sub

Private Sub mfgHUAWEI_SelChange()
   mfgHUAWEI_Click
End Sub


