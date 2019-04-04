VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDataupdate 
   Caption         =   "装箱清单后台维护"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   11085
   StartUpPosition =   2  '屏幕中心
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
      Left            =   5520
      TabIndex        =   12
      Top             =   960
      Width           =   1455
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
      Left            =   7200
      TabIndex        =   11
      Top             =   960
      Width           =   1455
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
      Left            =   8880
      TabIndex        =   10
      Top             =   960
      Width           =   1575
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
      Left            =   3720
      TabIndex        =   9
      Top             =   960
      Width           =   1575
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
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
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
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox chkYes 
      Caption         =   "是"
      Enabled         =   0   'False
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
      Left            =   9240
      TabIndex        =   5
      Top             =   240
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox chkNo 
      Caption         =   "否"
      Enabled         =   0   'False
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
      Left            =   10320
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtVer 
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
      TabIndex        =   2
      Top             =   240
      Width           =   2175
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgVer 
      Height          =   3975
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7011
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "装箱清单报废:"
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
      Left            =   7200
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblPerVer 
      Caption         =   "整机版本:"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   1335
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
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmDataupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim op As String
Dim con As ADODB.Connection
Dim rs3 As ADODB.Recordset
Dim str As String
Dim com As ADODB.Command
Dim status As String


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
        txtVer.Text = ""
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
   
   sql = "delete from tblPackList where Model='" & mfgVer.TextMatrix(mfgVer.RowSel, 1) & "'"

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
   
    If txtVer.Text = "" Then
        MsgBox "版本不能为空!!", vbExclamation + vbOKOnly, "版本空"
        txtNowVer.SetFocus
        Exit Sub
    End If
    
   If chkYes.Value = 0 And chkNo.Value = 0 Then
        MsgBox "装箱清单报废不能为空!!", vbExclamation + vbOKOnly, "装箱清单报废空"
        txtModel.SetFocus
        Exit Sub
   End If
   
   '==================
        
        If con.State = 1 Then
            con.Close
        End If
        con.Open
        
        Dim chkyn As String
        If chkYes.Value = 1 Then
            chkyn = "Y"
        Else
            chkyn = "N"
        End If
            
        Set rs3 = New ADODB.Recordset
        Set rs3.ActiveConnection = con
        rs3.CursorType = adOpenDynamic
        
        str = "select count(*) from tblPackList where model='" & Trim(txtModel.Text) & "' and Version='" & UCase(Trim(txtVer.Text)) & "'"
        rs3.Open str, con, adOpenKeyset, adLockOptimistic
        If rs3.Fields(0) > 0 Then
            MsgBox "此机种资料已经设置，如需修改请先查询", vbOKOnly + vbExclamation, "警告"
            txtModel.Text = ""
            txtVer.Text = ""
            chkNo.Value = 0
            rs3.Close
            con.Close
            Exit Sub
        Else
           
            str = "Insert into tblPackList( Model, Version, UseFlag) values " & _
            " ('" & UCase(Trim(txtModel.Text)) & "','" & UCase(Trim(txtVer.Text)) & "','" & chkyn & "')"

            Set com = New ADODB.Command
            Set com.ActiveConnection = con
            com.CommandText = str
            com.CommandType = adCmdText
            com.Execute
            
            txtModel.Text = ""
            txtVer.Text = ""
            chkNo.Value = 0
            
        End If
        rs3.Close
        con.Close
        
      renovate
End Sub

Private Sub cmdRefh_Click()
    txtModel.Text = ""
    txtVer.Text = ""
    chkNo.Value = 0
    
    renovate
        
End Sub


Private Sub cmdSearch_Click()
    If txtModel.Text = "" Then
        MsgBox "请输入机种作为查询条件!!", vbExclamation + vbOKOnly, "产品编码空"
        txtModel.SetFocus
        Exit Sub
    End If
    
    sql = "select * from tblPackList where model='" & Trim(txtModel.Text) & "'"
    
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
        .ColWidth(1) = 3000
        .ColWidth(2) = 3000
        .ColWidth(3) = 4000
        
        .TextMatrix(0, 1) = "机种"
        .TextMatrix(0, 2) = "整机版本"
        .TextMatrix(0, 3) = "装箱清单报废"
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
   
    If txtVer.Text = "" Then
        MsgBox "版本不能为空!!", vbExclamation + vbOKOnly, "版本空"
        txtNowVer.SetFocus
        Exit Sub
    End If
    
   If chkYes.Value = 0 And chkNo.Value = 0 Then
        MsgBox "装箱清单报废不能为空!!", vbExclamation + vbOKOnly, "装箱清单报废空"
        txtModel.SetFocus
        Exit Sub
   End If
   
   Dim chkyn As String
            If chkYes.Value = 1 Then
                chkyn = "Y"
            Else
                chkyn = "N"
            End If
            
   sql = "Update tblPackList set Version='" & UCase(Trim(txtVer.Text)) & "',UseFlag='" & chkyn & "'  where Model='" & mfgVer.TextMatrix(mfgVer.RowSel, 1) & "'"

   status = excuteUpdateSql(sql)
   If status <> "" Then
      MsgBox "修改资料失败!" & "原因是" & status
      Exit Sub
   End If
   MsgBox "修改资料成功!"
   
   renovate
   
    txtModel.Text = ""
    txtVer.Text = ""
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
   sql = "select * from tblPackList order by model"
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
        .ColWidth(2) = 2000
        .ColWidth(3) = 4000

        
        .TextMatrix(0, 1) = "机种"
        .TextMatrix(0, 2) = "整机版本"
        .TextMatrix(0, 3) = "装箱清单报废"

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
      txtVer.Text = mfgVer.TextMatrix(mfgVer.RowSel, 2)

      If mfgVer.TextMatrix(mfgVer.RowSel, 3) = "Y" Then
        chkYes.Value = 1
      Else
        chkNo.Value = 1
      End If
    
   End If
End Sub

Private Sub mfgVer_SelChange()
   mfgVer_Click
End Sub
