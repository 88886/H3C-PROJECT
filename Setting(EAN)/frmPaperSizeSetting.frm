VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPaperSizeSetting 
   Caption         =   "HP序列号类型维护"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "参数设定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10575
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除"
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查询"
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "退出"
         Height          =   375
         Left            =   8640
         TabIndex        =   11
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存更新"
         Height          =   375
         Left            =   7080
         TabIndex        =   10
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtProject 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton optSingle3 
         Caption         =   "纯HP14.6*7.7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CheckBox chkBackup 
         Caption         =   "是否打印副标签"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   2160
         Width           =   2175
      End
      Begin VB.OptionButton optSingle2 
         BackColor       =   &H0000FFFF&
         Caption         =   "纯HP标签48*6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   1560
         Width           =   2295
      End
      Begin VB.OptionButton optDouble2 
         BackColor       =   &H008080FF&
         Caption         =   "HP双标签48*6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton optDouble1 
         BackColor       =   &H008080FF&
         Caption         =   "HP双标签50*20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optSingle1 
         BackColor       =   &H0000FFFF&
         Caption         =   "纯HP标签50*20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   600
         TabIndex        =   2
         Top             =   1605
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "机种编号"
         BeginProperty Font 
            Name            =   "宋体"
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
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMain 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
Attribute VB_Name = "frmPaperSizeSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String


Private Sub reset()
    Me.txtProject.Text = ""
    Me.chkBackup.Value = False
    Me.optDouble1.Value = False
    Me.optDouble2.Value = False
    Me.optSingle1.Value = False
    Me.optSingle2.Value = False
    Me.optSingle3.Value = False
    renovate ("")
End Sub
Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Trim(Me.txtProject.Text) = "" Then
         MsgBox "删除机种为空!", vbOKOnly + vbInformation, "删除机种内容不能为空"
         Exit Sub
    Else
        If MsgBox("确定删除机种:" + Trim(Me.txtProject.Text) + "相关信息吗?", vbYesNo, "资料删除确认") <> vbYes Then
            Exit Sub
        End If
    End If

     Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "HPPaperSizeHandler"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 16, "Delete")
    cmd.Parameters.Append cmd.CreateParameter("part_number", adVarChar, adParamInput, 32, Me.txtProject.Text)
    cmd.Parameters.Append cmd.CreateParameter("size1", adBoolean, adParamInput, 1, Me.optDouble1.Value)
    cmd.Parameters.Append cmd.CreateParameter("size2", adBoolean, adParamInput, 1, Me.optDouble2.Value)
    cmd.Parameters.Append cmd.CreateParameter("size3", adBoolean, adParamInput, 1, Me.optSingle1.Value)
    cmd.Parameters.Append cmd.CreateParameter("size4", adBoolean, adParamInput, 1, Me.optSingle2.Value)
    cmd.Parameters.Append cmd.CreateParameter("size5", adBoolean, adParamInput, 1, Me.optSingle3.Value)
    cmd.Parameters.Append cmd.CreateParameter("backup", adBoolean, adParamInput, 1, Me.chkBackup.Value)
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    reset
    
End Sub

Private Sub cmdQuery_Click()
    sql = "select part_number, case [double_50*20] when 1 then 'Yes' else 'No' end ,"
    sql = sql + " case [double_48*6] when 1 then 'Yes' else 'No' end,case [single_50*20] when 1 then 'Yes' else 'No' end,"
    sql = sql + " case [single_48*6] when 1 then 'Yes' else 'No' end, case [single_14.6*7.7] when 1 then 'Yes' else 'No' end,"
    sql = sql + " case [backup] when 1 then 'Yes' else 'No' end from EAN_HP_Setting_2014 where part_number like '%" + Trim(Me.txtProject.Text) + "%'"
    renovate (sql)
End Sub

Private Sub cmdSave_Click()
    If IIf(Me.optDouble1.Value, 1, 0) + IIf(Me.optDouble2.Value, 1, 0) + IIf(Me.optSingle1.Value, 1, 0) + IIf(Me.optSingle2.Value, 1, 0) + IIf(Me.optSingle3.Value, 1, 0) = 1 Then
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = conn
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "HPPaperSizeHandler"
        cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 16, "Update")
        cmd.Parameters.Append cmd.CreateParameter("part_number", adVarChar, adParamInput, 32, Me.txtProject.Text)
        cmd.Parameters.Append cmd.CreateParameter("size1", adBoolean, adParamInput, 1, Me.optDouble1.Value)
        cmd.Parameters.Append cmd.CreateParameter("size2", adBoolean, adParamInput, 1, Me.optDouble2.Value)
        cmd.Parameters.Append cmd.CreateParameter("size3", adBoolean, adParamInput, 1, Me.optSingle1.Value)
        cmd.Parameters.Append cmd.CreateParameter("size4", adBoolean, adParamInput, 1, Me.optSingle2.Value)
        cmd.Parameters.Append cmd.CreateParameter("size5", adBoolean, adParamInput, 1, Me.optSingle3.Value)
        cmd.Parameters.Append cmd.CreateParameter("backup", adBoolean, adParamInput, 1, Me.chkBackup.Value)
        cmd.Execute
        Set cmd.ActiveConnection = Nothing
        reset
    Else
         MsgBox "选择错误!", vbOKOnly + vbInformation, "前五个选项中只能选择其中一个"
    End If
    
End Sub



Private Sub Form_Load()
'    unable
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    Me.chkBackup.Enabled = False
    Me.chkBackup.Value = 0
   
    renovate ("")
    
End Sub

Private Sub renovate(sql As String)
   gridMain.Clear
   If Trim(sql) = "" Then
      sql = "select part_number, case [double_50*20] when 1 then 'Yes' else 'No' end ,"
      sql = sql + " case [double_48*6] when 1 then 'Yes' else 'No' end,case [single_50*20] when 1 then 'Yes' else 'No' end,"
      sql = sql + " case [single_48*6] when 1 then 'Yes' else 'No' end, case [single_14.6*7.7] when 1 then 'Yes' else 'No' end,"
      sql = sql + " case [backup] when 1 then 'Yes' else 'No' end from EAN_HP_Setting_2014 order by part_number"
   End If
   
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set gridMain.DataSource = rec
   With gridMain
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 100
        .ColWidth(1) = 2000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500

'HP双标签50*20   HP双标签48*6    纯HP标签50*20   纯HP标签48*6    纯HP14.6*7.7

        
        .TextMatrix(0, 1) = "产品编码(Model Number)"
        .TextMatrix(0, 2) = "HP双标签50*20"
        .TextMatrix(0, 3) = "HP双标签48*6"
        .TextMatrix(0, 4) = "纯HP标签50*20"
        .TextMatrix(0, 5) = "纯HP标签48*6"
        .TextMatrix(0, 6) = "纯HP14.6*7.7"
        .TextMatrix(0, 7) = "副标签是否打印"

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

Private Sub gridMain_Click()
    If gridMain.RowSel > 0 Then
      txtProject.Text = gridMain.TextMatrix(gridMain.RowSel, 1)
      
      If gridMain.TextMatrix(gridMain.RowSel, 2) = "Yes" Then
        Me.optDouble1.Value = True
      Else
        Me.optDouble1.Value = False
      End If
      
      If gridMain.TextMatrix(gridMain.RowSel, 3) = "Yes" Then
        Me.optDouble2.Value = True
      Else
        Me.optDouble2.Value = False
      End If
      
      If gridMain.TextMatrix(gridMain.RowSel, 4) = "Yes" Then
        Me.optSingle1.Value = True
      Else
        Me.optSingle1.Value = False
      End If
      
      If gridMain.TextMatrix(gridMain.RowSel, 5) = "Yes" Then
        Me.optSingle2.Value = True
      Else
        Me.optSingle2.Value = False
      End If
      
      If gridMain.TextMatrix(gridMain.RowSel, 6) = "Yes" Then
        Me.optSingle3.Value = True
      Else
        Me.optSingle3.Value = False
      End If
      
      If gridMain.TextMatrix(gridMain.RowSel, 7) = "Yes" Then
        Me.chkBackup.Value = 1
      Else
        Me.chkBackup.Value = 0
      End If
      
   End If
End Sub
   
Private Sub optDouble1_Click()
    If Me.optDouble1.Value = True Then
        Me.chkBackup.Enabled = False
        Me.chkBackup.Value = 0
    End If
    
End Sub

Private Sub optDouble2_Click()
    If Me.optDouble2.Value = True Then
        Me.chkBackup.Enabled = False
        Me.chkBackup.Value = 0
    End If
End Sub

Private Sub optSingle1_Click()
    If Me.optSingle1.Value = True Then
        Me.chkBackup.Enabled = False
        Me.chkBackup.Value = 0
    End If
End Sub

Private Sub optSingle2_Click()
    If Me.optSingle2.Value = True Then
        Me.chkBackup.Enabled = False
        Me.chkBackup.Value = 0
    End If
End Sub

Private Sub optSingle3_Click()
    If Me.optSingle3.Value = True Then
        Me.chkBackup.Enabled = True
    Else
        Me.chkBackup.Enabled = False
        Me.chkBackup.Value = 0
    End If
End Sub
