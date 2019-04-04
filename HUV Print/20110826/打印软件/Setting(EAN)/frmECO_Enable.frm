VERSION 5.00
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
   Begin VB.CommandButton cmdRefesh 
      Caption         =   "刷新"
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查询"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   120
      Width           =   975
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
      Left            =   3960
      MaxLength       =   12
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmbAddRight 
      Caption         =   "加入"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   3
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
      Left            =   3960
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "禁用"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4815
      Left            =   3240
      TabIndex        =   0
      Top             =   1680
      Width           =   4095
      Begin VB.ListBox lstUnuse 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4260
         ItemData        =   "frmECO_Enable.frx":0000
         Left            =   240
         List            =   "frmECO_Enable.frx":0002
         TabIndex        =   6
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10320
      Y1              =   600
      Y2              =   600
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
      Left            =   3000
      TabIndex        =   4
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
      Left            =   3000
      TabIndex        =   2
      Top             =   840
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

    
    If Trim(txtModel.Text) <> "" Then
        If txtNonuse.Text <> "" Then
            sql = "Insert into tblECO_Ver(PartNumber,Version,Active) " & _
            "Values('" & Trim(txtModel.Text) & "','" & Trim(txtNonuse.Text) & "',0)"
            status = Connect.excuteUpdate(sql)
     
            If status <> "" Then
            Else
                lstUnuse.AddItem UCase(Trim(Me.txtNonuse.Text))
            End If
        End If
    End If
    

    Me.txtNonuse.Text = ""
    
    'renovate_right
End Sub

Private Sub cmdRefesh_Click()
    Me.txtModel.Text = ""
    Me.txtModel.Enabled = True
    Me.cmdSearch.Enabled = True
    Me.cmbAddRight.Enabled = False

    Me.lstUnuse.Clear
    
End Sub

Private Sub cmdSearch_Click()

    If Trim(txtModel.Text) <> "" Then
        
      cmbAddRight.Enabled = True

      cmdSearch.Enabled = False
      
      renovate_right
      
    End If
    
End Sub


Private Sub renovate_right()
    
    sql = "select Version from tblECO_Ver where PartNumber='" & Trim(Me.txtModel.Text) & "' and Active=0 order by version "
   If rec.State = 1 Then
      rec.Close
   End If
   
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    
    Do While rec.EOF = False
    
        lstUnuse.AddItem rec.Fields(0)
        rec.MoveNext
    Loop
    
   rec.Close

End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
    End If
End Sub
