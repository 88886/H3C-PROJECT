VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmModelVer 
   Caption         =   "软件版本维护"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13170
   LinkTopic       =   "软件版本维护"
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13170
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdRefh 
      Caption         =   "刷 新"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   17
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清 空"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查 询"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   15
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修 改"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   14
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删 除"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   13
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "新 增"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   12
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CheckBox chkNo 
      Caption         =   "否"
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
      Left            =   10680
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.CheckBox chkYes 
      Caption         =   "是"
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
      Left            =   9600
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtNowVer 
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
      Left            =   4800
      TabIndex        =   7
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtPerVer 
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
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtModel 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   495
      Left            =   9120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      Format          =   103350273
      CurrentDate     =   40908
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgVer 
      Height          =   6015
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   10610
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "是否自动抓取:"
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
      Left            =   7560
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "现版本:"
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
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "有效日期:"
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
      Left            =   7680
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPerVer 
      Caption         =   "前版本:"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblModel 
      Caption         =   "机种:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmModelVer"
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
Dim con As ADODB.Connection
Dim rs3 As ADODB.Recordset
Dim str As String
Dim com As ADODB.Command
Dim status As String

Private Sub enable()

End Sub

Private Sub unable()
   
End Sub

Private Sub chkNo_Click()
    If chkNo.Value = 1 Then
        chkYes.Value = 0
    Else
        chkYes.Value = 1
    End If
    
End Sub

Private Sub chkYes_Click()
    If chkYes.Value = 1 Then
        chkNo.Value = 0
    Else
        chkNo.Value = 1
    End If
End Sub


Private Sub cmdClear_Click()
        txtModel.Text = ""
        txtPerVer.Text = ""
        txtNowVer.Text = ""
End Sub

Public Function excuteUpdateSql(sSQLStatement As String) As String

    If con.State = 1 Then
        con.Close
    End If
    con.Open
    
    On Error GoTo errorHandler
    con.Execute (sSQLStatement)
    excuteUpdateSql = ""
    
    Exit Function
errorHandler:
    excuteUpdateSql = Err.Description
    
End Function

Private Sub cmdDelete_Click()
    If mfgVer.RowSel <= 0 Or Trim(txtModel.Text) = "" Then
      MsgBox "请选择要删除的行!"
      Exit Sub
   End If
   
   sql = " insert into tblSoftVersion_log(CREATE_USER,Model, beforeVer, endDate, nowVer, searchFlag,comment) "
   sql = sql & " select '" & golUSERNAME & "',Model, beforeVer, endDate, nowVer, searchFlag,'delete' from tblSoftVersion where  Model='" & mfgVer.TextMatrix(mfgVer.RowSel, 1) & "'"
   
   sql = sql & " delete from tblSoftVersion where Model='" & mfgVer.TextMatrix(mfgVer.RowSel, 1) & "'"
   
   status = excuteUpdateSql(sql)
   If status <> "" Then
      MsgBox "删除资料失败!" & "原因是" & status
      Exit Sub
   End If
   MsgBox "删除资料成功!"
   
   renovate
   
End Sub

Private Sub cmdInsert_Click()
    If txtModel.Text = "" Then
        MsgBox "机种不能为空!!", vbExclamation + vbOKOnly, "机种空"
        txtModel.SetFocus
        Exit Sub
    End If
   
    If txtNowVer.Text = "" Then
        MsgBox "现版本不能为空!!", vbExclamation + vbOKOnly, "现版本空"
        txtNowVer.SetFocus
        Exit Sub
    End If
    
   If chkYes.Value = 0 And chkNo.Value = 0 Then
        MsgBox "是否自动抓取不能为空!!", vbExclamation + vbOKOnly, "是否自动抓取空"
        txtModel.SetFocus
        Exit Sub
   End If
   
   '==================
        
        If con.State = 1 Then
            con.Close
        End If
        con.Open
        
        Set rs3 = New ADODB.Recordset
        Set rs3.ActiveConnection = con
        rs3.CursorType = adOpenDynamic
        
        str = "select count(*) from tblSoftVersion where model='" & Trim(txtModel.Text) & "'"
        rs3.Open str, con, adOpenKeyset, adLockOptimistic
        If rs3.Fields(0) > 0 Then
            MsgBox "此机种资料已经设置，如需修改请先查询", vbOKOnly + vbExclamation, "警告"
            txtModel.Text = ""
            txtPerVer.Text = ""
            txtNowVer.Text = ""
            chkYes.Value = 0
            chkNo.Value = 0
            rs3.Close
            con.Close
            
            Exit Sub
        Else
            Dim chkyn As String
            If chkYes.Value = 1 Then
                chkyn = "Y"
            Else
                chkyn = "N"
            End If
            
            str = "Insert into tblSoftVersion( Model, beforeVer, endDate, nowVer, searchFlag) values " & _
            " ('" & UCase(Trim(txtModel.Text)) & "','" & UCase(Trim(txtPerVer.Text)) & "','" & UCase(Trim(dtpEndDate.Value)) & "','" & UCase(Trim(txtNowVer.Text)) & "','" & chkyn & "')"
            str = str & " Insert into tblSoftVersion_log( create_user,Model, beforeVer, endDate, nowVer, searchFlag,comment) values " & _
            " ('" & golUSERNAME & "','" & UCase(Trim(txtModel.Text)) & "','" & UCase(Trim(txtPerVer.Text)) & "','" & UCase(Trim(dtpEndDate.Value)) & "','" & UCase(Trim(txtNowVer.Text)) & "','" & chkyn & "','Insert')"

            
            
            Set com = New ADODB.Command
            Set com.ActiveConnection = con
            com.CommandText = str
            com.CommandType = adCmdText
            com.Execute
            
            txtModel.Text = ""
            txtPerVer.Text = ""
            txtNowVer.Text = ""
            chkYes.Value = 0
            chkNo.Value = 0
            
        End If
        rs3.Close
        con.Close
        
      renovate
End Sub

Private Sub cmdRefh_Click()
    txtModel.Text = ""
    txtPerVer.Text = ""
    txtNowVer.Text = ""
    chkYes.Value = 0
    chkNo.Value = 0
    
    renovate
        
End Sub

Private Sub cmdSearch_Click()
    If txtModel.Text = "" Then
        MsgBox "请输入机种作为查询条件!!", vbExclamation + vbOKOnly, "产品编码空"
        txtModel.SetFocus
        Exit Sub
    End If
    
    sql = "select * from tblSoftVersion where model='" & Trim(txtModel.Text) & "'"
    
    If con.State = 1 Then
      con.Close
   End If
   
   con.Open
    
    Set rs3 = New ADODB.Recordset
    rs3.ActiveConnection = con

    rs3.Open sql, con, adOpenKeyset, adLockOptimistic
    
   Set mfgVer.DataSource = rs3
   With mfgVer
        .Cols = rs3.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 2000
        .ColWidth(2) = 3500
        .ColWidth(3) = 2000
        .ColWidth(4) = 3500
        .ColWidth(5) = 1000
        
        .TextMatrix(0, 1) = "机种"
        .TextMatrix(0, 2) = "前版本"
        .TextMatrix(0, 3) = "有效日期"
        .TextMatrix(0, 4) = "现版本"
        .TextMatrix(0, 5) = "是否自动抓取"
   End With
   rs3.Close
   con.Close
   
End Sub

Private Sub cmdUpdate_Click()
   If mfgVer.RowSel <= 0 Or Trim(txtModel.Text) = "" Then
      MsgBox "请选择要修改的行!"
      Exit Sub
   End If
   
   If mfgVer.TextMatrix(mfgVer.RowSel, 1) <> UCase(Trim(txtModel.Text)) Then
        MsgBox "机种不可以修改!"
      Exit Sub
   End If
   
   If txtModel.Text = "" Then
        MsgBox "机种不能为空!!", vbExclamation + vbOKOnly, "机种空"
        txtModel.SetFocus
        Exit Sub
    End If
   
    If txtNowVer.Text = "" Then
        MsgBox "现版本不能为空!!", vbExclamation + vbOKOnly, "现版本空"
        txtNowVer.SetFocus
        Exit Sub
    End If
    
   If chkYes.Value = 0 And chkNo.Value = 0 Then
        MsgBox "是否自动抓取不能为空!!", vbExclamation + vbOKOnly, "是否自动抓取空"
        txtModel.SetFocus
        Exit Sub
   End If
   
   Dim chkyn As String
            If chkYes.Value = 1 Then
                chkyn = "Y"
            Else
                chkyn = "N"
            End If
            
   sql = "Update tblSoftVersion set beforeVer='" & UCase(Trim(txtPerVer.Text)) & "',endDate='" & dtpEndDate.Value & "',nowVer='" & UCase(Trim(txtNowVer.Text)) & "',searchFlag='" & chkyn & "'  where Model='" & mfgVer.TextMatrix(mfgVer.RowSel, 1) & "'"
   
   sql = sql & " Insert into tblSoftVersion_log( create_user,Model, beforeVer, endDate, nowVer, searchFlag,comment) values " & _
            " ('" & golUSERNAME & "','" & mfgVer.TextMatrix(mfgVer.RowSel, 1) & "','" & UCase(Trim(txtPerVer.Text)) & "','" & UCase(Trim(dtpEndDate.Value)) & "','" & UCase(Trim(txtNowVer.Text)) & "','" & chkyn & "','Update')"
   status = excuteUpdateSql(sql)
   If status <> "" Then
      MsgBox "修改资料失败!" & "原因是" & status
      Exit Sub
   End If
   MsgBox "修改资料成功!"
   
   renovate
   
    txtModel.Text = ""
    txtPerVer.Text = ""
    txtNowVer.Text = ""
    chkYes.Value = 0
    chkNo.Value = 0
            
End Sub

Private Sub Form_Load()
   
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   
    Set con = New ADODB.Connection
    con.CursorLocation = adUseClient

    con.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
    con.ConnectionTimeout = 100
        
   renovate
End Sub

Private Sub renovate()
   sql = "select * from tblSoftVersion order by model"
   If con.State = 1 Then
      con.Close
   End If
   
   con.Open
   
    Set rs3 = New ADODB.Recordset
    rs3.ActiveConnection = con

    rs3.Open sql, con, adOpenKeyset, adLockOptimistic
    
   Set mfgVer.DataSource = rs3
   With mfgVer
        .Cols = rs3.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 2000
        .ColWidth(2) = 3500
        .ColWidth(3) = 2000
        .ColWidth(4) = 3500
        .ColWidth(5) = 1000
        
        .TextMatrix(0, 1) = "机种"
        .TextMatrix(0, 2) = "前版本"
        .TextMatrix(0, 3) = "有效日期"
        .TextMatrix(0, 4) = "现版本"
        .TextMatrix(0, 5) = "是否自动抓取"
   End With
   rs3.Close
   con.Close
   
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

Private Sub mfgVer_Click()
   If mfgVer.RowSel > 0 Then
        
      txtModel.Text = mfgVer.TextMatrix(mfgVer.RowSel, 1)
      txtPerVer.Text = mfgVer.TextMatrix(mfgVer.RowSel, 2)
      dtpEndDate.Value = mfgVer.TextMatrix(mfgVer.RowSel, 3)
      txtNowVer.Text = mfgVer.TextMatrix(mfgVer.RowSel, 4)
      If mfgVer.TextMatrix(mfgVer.RowSel, 5) = "Y" Then
        chkYes.Value = 1
      Else
        chkNo.Value = 1
      End If
    
   End If
End Sub

Private Sub mfgVer_SelChange()
   mfgVer_Click
End Sub


