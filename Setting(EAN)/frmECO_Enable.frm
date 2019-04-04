VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmECO_Enable 
   Caption         =   "条码ECO版本防呆"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10335
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgMain 
      Height          =   5055
      Left            =   480
      TabIndex        =   8
      Top             =   1440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8916
      _Version        =   393216
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CommandDelRight 
      Caption         =   "删除"
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修改"
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查询"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   120
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
      Height          =   375
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmbAddRight 
      Caption         =   "加入"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtNonuse 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "版本:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmECO_Enable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String


Private Sub cmbAddRight_Click()
    If Len(Trim(Me.txtModel.Text)) <> 11 Then
        MsgBox "机种编号必须是11位", vbInformation + vbOKOnly, "机种编号不正确"
    End If
    
    If Mid(UCase(Me.txtModel.Text), 1, 3) <> "HWF" And Mid(UCase(Me.txtModel.Text), 1, 3) <> "HUV" Then
        MsgBox "机种编号开头必须是HWF或者HUV", vbInformation + vbOKOnly, "机种编号不正确"
    End If
    
    
    If Trim(txtModel.Text) <> "" Then
        If txtNonuse.Text <> "" Then
            sql = "Insert into tblECO_Ver(PartNumber,Version,Active) " & _
            "Values('" & UCase(Trim(txtModel.Text)) & "','" & UCase(Trim(txtNonuse.Text)) & "',0)"
            status = Connect.excuteUpdate(sql)
           renovate_right
        End If
    End If
    
    Me.txtModel.Text = ""
    Me.txtNonuse.Text = ""
    
    'renovate_right
End Sub



Private Sub cmdSearch_Click()
      renovate_right
End Sub


Private Sub renovate_right()
    
   sql = "select PartNumber,Version, case when Active = 0 then 'Yes' when Active = 1 then 'No' else 'Non' end from tblECO_Ver Where 1 = 1"
   
   If (Trim(Me.txtModel.Text) <> "") Then
        sql = sql + " and PartNumber='" & UCase(Trim(Me.txtModel.Text)) & "'"
   End If
   
   If Trim(Me.txtNonuse.Text) <> "" Then
        sql = sql & " and Version like '" & UCase(Trim(Me.txtNonuse.Text)) & "'"
     End If
   If rec.State = 1 Then
      rec.Close
   End If
   
   rec.Open sql, conn, adOpenKeyset, adLockReadOnly
   
   Set mfgMain.DataSource = rec
   
    With mfgMain
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 2000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
       
        
        .TextMatrix(0, 1) = "料号"
        .TextMatrix(0, 2) = "版本"
        .TextMatrix(0, 3) = "是否禁用"
      
   End With
    
   rec.Close

End Sub

Private Sub cmdUpdate_Click()
   If Trim(Me.txtModel.Text) <> "" And Trim(Me.txtNonuse.Text) <> "" Then
        sql = "update tblECO_Ver set Active = 0 where PartNumber = '" & UCase(Trim(Me.txtModel.Text)) & "' and Version ='" & UCase(Trim(Me.txtNonuse.Text)) & "'"
        status = Connect.excuteUpdate(sql)
        If status <> "" Then
           MsgBox "更新ECO Version 禁用设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "删除失败"
        End If
        MsgBox "更新ECO设定资料成功!", vbInformation + vbOKOnly, "更新成功"
        renovate_right
   End If
   
End Sub

Private Sub CommandDelRight_Click()
   If mfgMain.RowSel <= 0 Then
      MsgBox "请选择要删除的行!", vbInformation + vbOKOnly, "未选择删除行"
      Exit Sub
   End If
   sql = "delete from tblECO_Ver where PartNumber = '" & mfgMain.TextMatrix(mfgMain.RowSel, 1) & "' and Version ='" & mfgMain.TextMatrix(mfgMain.RowSel, 2) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "删除ECO Version 设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "删除失败"
   End If
   MsgBox "删除H3C设定资料成功!", vbInformation + vbOKOnly, "删除成功"
   renovate_right
End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
    End If
End Sub


