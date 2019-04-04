VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCustomType 
   Caption         =   "整机序列号品牌维护"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "frmCustomType"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9255
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6165
      _Version        =   393216
      AllowUserResizing=   1
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
   Begin VB.CheckBox chkNonH3C 
      Caption         =   "非H3C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox chkH3C 
      Caption         =   "H3C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtModel 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      MaxLength       =   12
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "类别:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "机种:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCustomType"
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

Private Sub chkH3C_Click()
    If chkH3C.Value = 1 Then
        chkNonH3C.Value = 0
    Else
        chkNonH3C.Value = 1
    End If
End Sub

Private Sub chkNonH3C_Click()
    If chkNonH3C.Value = 1 Then
        chkH3C.Value = 0
    Else
        chkH3C.Value = 1
    End If
End Sub


Private Sub renovate()
   sql = "select * from tblCustomType order by PartNumber"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set MSHFlexGrid1.DataSource = rec
   With MSHFlexGrid1
        .Cols = rec.Fields.Count + 1

        .ColWidth(0) = 400
        .ColWidth(1) = 3000
        .ColWidth(2) = 3000

        
        .TextMatrix(0, 1) = "产品机种"
        .TextMatrix(0, 2) = "产品类别"
       
   End With
   rec.Close
End Sub

Private Sub Form_Load()

   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   renovate
End Sub

Private Sub cmdAdd_Click()
    If txtModel.Text = "" Then
        MsgBox "产品编码不能为空!!", vbExclamation + vbOKOnly, "产品编码空"
        txtModel.SetFocus
        Exit Sub
    End If
   
   If chkH3C.Value = 0 And chkNonH3C.Value = 0 Then
        MsgBox "产品类别不能为空!!", vbExclamation + vbOKOnly, "产品类别空"
        txtModel.SetFocus
        Exit Sub
   End If
   
   
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from tblCustomType where PartNumber='" & Trim(txtModel.Text) & "' "
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
            sql = "update tblCustomType set Type='" & IIf(chkH3C.Value = 1, "H3C", "Non-H3C") & "' where PartNumber= '" & Trim(txtModel.Text) & "'"
            status = Connect.excuteUpdate(sql)
            If status <> "" Then
                MsgBox "新增资料失败!" & "原因是" & status, vbOKOnly + vbInformation, "新增失败"
            End If
            MsgBox "新增资料成功!", vbOKOnly + vbInformation, "新增成功"
      Else
            sql = "Insert into tblCustomType(PartNumber,Type) " & _
            "Values(N'" & Replace(Trim(txtModel.Text), Chr(13) & Chr(10), "") & "','" & IIf(chkH3C.Value = 1, "H3C", "Non-H3C") & "')"
            status = Connect.excuteUpdate(sql)
            If status <> "" Then
                MsgBox "新增资料失败!" & "原因是" & status, vbOKOnly + vbInformation, "新增失败"
            End If
            MsgBox "新增资料成功!", vbOKOnly + vbInformation, "新增成功"
      End If
      rcd.Close
      
      
      renovate
   
End Sub
