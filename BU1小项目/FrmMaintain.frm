VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmMaintain 
   Caption         =   "Description"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   10215
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "刷新"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtID 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMain 
      Height          =   2655
      Left            =   1560
      TabIndex        =   10
      Top             =   2520
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "创建"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除"
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtDescription 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtCustomerPN 
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtFlashPN 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "ID:"
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Description:"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Customer PN:"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Flash PN:"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Description setting up"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "FrmMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub renovate()
  Dim rec As New ADODB.Recordset
   sql = "select [part_number],[customer_part],[description],[ID] from [BU1_PrintPartMapping]"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set gridMain.DataSource = rec
   With gridMain
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 1500
        .ColWidth(1) = 1500
        .ColWidth(2) = 2000
        .ColWidth(4) = 500
        .TextMatrix(0, 1) = "Flash PN"
        .TextMatrix(0, 2) = "Customer PN"
        .TextMatrix(0, 3) = "Description"
        .TextMatrix(0, 4) = "ID"
   End With
   rec.Close
End Sub


Private Sub cmdCreate_Click()
    Me.txtFlashPN.Text = ""
    Me.txtCustomerPN.Text = ""
    Me.txtDescription.Text = ""
    Me.txtID.Text = ""
    Me.txtFlashPN.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim flashPN As String
    Dim customerPN As String
    Dim description As String
    Dim id As Integer
    flashPN = Trim(Me.txtFlashPN.Text)
    customerPN = Trim(Me.txtCustomerPN.Text)
    description = Trim(Me.txtDescription.Text)
    If (Trim(Me.txtID.Text) = "") Then
        txtID.Text = "0"
    End If
    id = CInt(Me.txtID)
    Dim res As Integer
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BU1PrintHandler"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 8, "delete")
    cmd.Parameters.Append cmd.CreateParameter("part_number", adVarChar, adParamInput, 32, Me.txtFlashPN.Text)
    cmd.Parameters.Append cmd.CreateParameter("customer_part", adVarChar, adParamInput, 32, Me.txtCustomerPN.Text)
    cmd.Parameters.Append cmd.CreateParameter("decription", adVarChar, adParamInput, 64, Me.txtDescription.Text)
    cmd.Parameters.Append cmd.CreateParameter("id", adInteger, adParamInput, 4, id)
    cmd.Execute
    MsgBox "Updated Successfully"
    Set cmd.ActiveConnection = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim flashPN As String
    Dim customerPN As String
    Dim description As String
    Dim id As Integer
    flashPN = Trim(Me.txtFlashPN.Text)
    customerPN = Trim(Me.txtCustomerPN.Text)
    description = Trim(Me.txtDescription.Text)
    If (Trim(Me.txtID.Text) = "") Then
        txtID.Text = "0"
    End If
    id = CInt(Me.txtID)
    If (flashPN = "" Or customerPN = "" Or description = "") Then
        MsgBox "料号或者客户料号或者描述不能为空"
        txtFlashPN.SetFocus
        Exit Sub
    End If
    
    
    Dim res As Integer
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BU1PrintHandler"
    cmd.Parameters.Append cmd.CreateParameter("action", adVarChar, adParamInput, 8, "save")
    cmd.Parameters.Append cmd.CreateParameter("part_number", adVarChar, adParamInput, 32, Me.txtFlashPN.Text)
    cmd.Parameters.Append cmd.CreateParameter("customer_part", adVarChar, adParamInput, 32, Me.txtCustomerPN.Text)
    cmd.Parameters.Append cmd.CreateParameter("decription", adVarChar, adParamInput, 64, Me.txtDescription.Text)
    cmd.Parameters.Append cmd.CreateParameter("id", adInteger, adParamInput, 4, id)
    cmd.Execute
    MsgBox "Updated Successfully"
    Set cmd.ActiveConnection = Nothing
End Sub

Private Sub Command1_Click()
    renovate
End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring()
        conn.Open
    End If
    
    renovate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If conn.State = 1 Then
        conn.Close
    End If
    
End Sub


Private Sub gridMain_Click()
    If Me.gridMain.RowSel > 0 Then
        txtFlashPN.Text = Me.gridMain.TextMatrix(Me.gridMain.RowSel, 1)
        txtCustomerPN.Text = Me.gridMain.TextMatrix(Me.gridMain.RowSel, 2)
        txtDescription.Text = Me.gridMain.TextMatrix(Me.gridMain.RowSel, 3)
        txtID.Text = Me.gridMain.TextMatrix(Me.gridMain.RowSel, 4)
    End If
End Sub
