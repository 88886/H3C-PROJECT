VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmHPDescMaintain 
   Caption         =   "HP 描述维护1"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10230
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除"
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
      Left            =   6000
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查询"
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
      Left            =   2880
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtDesc2 
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
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   6615
   End
   Begin VB.TextBox txtDesc 
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
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Width           =   6615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退出"
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
      Left            =   7440
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存更新"
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
      Left            =   4320
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
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
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMain 
      Height          =   3975
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7011
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "HP 描述维护2"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "HP 机种"
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
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label label5 
      Caption         =   "HP 描述维护1"
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
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmHPDescMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset

Private Sub reset()
    Me.txtProject.Text = ""
    Me.txtDesc2.Text = ""
    Me.txtDesc.Text = ""
    Me.gridMain.Clear
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
    cmd.CommandText = "HPDescHandler"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 16, "Delete")
    cmd.Parameters.Append cmd.CreateParameter("part_number", adVarChar, adParamInput, 16, Me.txtProject.Text)
    cmd.Parameters.Append cmd.CreateParameter("hp_desc", adVarChar, adParamInput, 50, Me.txtDesc.Text)
    cmd.Parameters.Append cmd.CreateParameter("hp_desc2", adVarChar, adParamInput, 50, Me.txtDesc2.Text)
 
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    reset
End Sub

Private Sub cmdQuery_Click()
    sql = "select part_number,hp_desc,hp_desc2 from ean_hpDesc_maintain"
    sql = sql + " where part_number like '%" + Me.txtProject + "%'"
    renovate (sql)
End Sub

Private Sub cmdSave_Click()
    If Trim(Me.txtProject.Text) = "" Then
         MsgBox "机种为空!", vbOKOnly + vbInformation, "机种内容不能为空"
         Exit Sub
    End If
    
    If Trim(Me.txtDesc.Text) = "" Then
         MsgBox "描述1为空!", vbOKOnly + vbInformation, "描述1内容不能为空"
         Exit Sub
    End If
    
    If Me.txtDesc2.Text = "" Then
         MsgBox "描述2为空!", vbOKOnly + vbInformation, "描述2内容不能为空"
         Exit Sub
    End If
  
        
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "HPDescHandler"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 16, "Update")
    cmd.Parameters.Append cmd.CreateParameter("part_number", adVarChar, adParamInput, 16, Me.txtProject.Text)
    cmd.Parameters.Append cmd.CreateParameter("hp_desc", adVarChar, adParamInput, 50, Me.txtDesc.Text)
    cmd.Parameters.Append cmd.CreateParameter("hp_desc2", adVarChar, adParamInput, 50, Me.txtDesc2.Text)
 
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    reset
    
End Sub



Private Sub Form_Load()
'    unable
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
   renovate ("")
End Sub
Private Sub renovate(sql As String)
   gridMain.Clear
   If Trim(sql) = "" Then
        sql = "select part_number,hp_desc,hp_desc2 from ean_hpDesc_maintain"
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
        .ColWidth(2) = 4000
        .ColWidth(3) = 4000

'HP双标签50*20   HP双标签48*6    纯HP标签50*20   纯HP标签48*6    纯HP14.6*7.7

        
        .TextMatrix(0, 1) = "产品编码(Model Number)"
        .TextMatrix(0, 2) = "HP描述1"
        .TextMatrix(0, 3) = "HP描述2"

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
      
      If gridMain.TextMatrix(gridMain.RowSel, 2) <> "" Then
        Me.txtDesc.Text = gridMain.TextMatrix(gridMain.RowSel, 2)
      Else
        Me.txtDesc.Text = gridMain.TextMatrix(gridMain.RowSel, 2)
      End If
      
      If gridMain.TextMatrix(gridMain.RowSel, 3) <> "" Then
        Me.txtDesc2.Text = gridMain.TextMatrix(gridMain.RowSel, 3)
      End If
      
   End If
End Sub
   

