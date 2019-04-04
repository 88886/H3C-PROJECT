VERSION 5.00
Begin VB.Form frmECO 
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
   Begin VB.TextBox lblModel 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
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
      Left            =   4560
      TabIndex        =   17
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdRefesh 
      Caption         =   "刷新"
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查询"
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdToLeft 
      Caption         =   "<"
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdToRight 
      Caption         =   ">"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   3360
      Width           =   615
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
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmbAddRight 
      Caption         =   "加入"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmbAddLeft 
      Caption         =   "加入"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1320
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
      Left            =   6480
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtActive 
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
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
      Left            =   5640
      TabIndex        =   1
      Top             =   2040
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
         ItemData        =   "frmECO.frx":0000
         Left            =   240
         List            =   "frmECO.frx":0002
         TabIndex        =   11
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "可用"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   4815
      Left            =   720
      TabIndex        =   0
      Top             =   2040
      Width           =   4095
      Begin VB.ListBox lstActive 
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
         ItemData        =   "frmECO.frx":0004
         Left            =   240
         List            =   "frmECO.frx":0006
         TabIndex        =   10
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Label Label3 
      Caption         =   "当前显示机种:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   840
      Width           =   1695
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
      TabIndex        =   8
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
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Lable1 
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
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "frmECO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String

Private Sub cmbAddLeft_Click()

    If Trim(lblModel.Text) = "" Then
       lblModel.Text = Trim(txtModel.Text)
    End If
    
    If Trim(lblModel.Text) <> "" Then
    
         If Trim(txtActive.Text) <> "" Then
            sql = "Insert into tblECO_Ver(PartNumber,Version,Active) " & _
            "Values('" & lblModel.Text & "','" & Trim(txtActive.Text) & "',1)"
            status = Connect.excuteUpdate(sql)
     
            If status <> "" Then
            Else
                lstActive.AddItem UCase(Trim(Me.txtActive.Text))
            End If
        End If
            
    End If
    

    Me.txtActive.Text = ""
    
    'renovate_left
    
End Sub

Private Sub cmbAddRight_Click()
    If Trim(lblModel.Text) = "" Then
       lblModel.Text = Trim(txtModel.Text)
    End If
    
    If Trim(lblModel.Text) <> "" Then
        If txtNonuse.Text <> "" Then
            sql = "Insert into tblECO_Ver(PartNumber,Version,Active) " & _
            "Values('" & lblModel.Text & "','" & Trim(txtNonuse.Text) & "',0)"
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
    Me.cmbAddLeft.Enabled = False
    Me.cmbAddRight.Enabled = False
    Me.lblModel.Text = ""
    
    Me.lstActive.Clear
    Me.lstUnuse.Clear
    
End Sub

Private Sub cmdSearch_Click()

    If Trim(txtModel.Text) <> "" Then
        
      lblModel.Text = Trim(txtModel.Text)
      cmbAddLeft.Enabled = True
      cmbAddRight.Enabled = True
      txtModel.Text = ""
      
      cmdSearch.Enabled = False
      
      renovate_left
      renovate_right
      
    End If
    
End Sub

Private Sub renovate_left()

   sql = "select Version from tblECO_Ver where PartNumber='" & Trim(lblModel.Text) & "' and Active=1 order by version "
   If rec.State = 1 Then
      rec.Close
   End If
   
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    
    Do While rec.EOF = False
    
        lstActive.AddItem rec.Fields(0)
        rec.MoveNext
    Loop
    
   rec.Close
End Sub

Private Sub renovate_right()
    
    sql = "select Version from tblECO_Ver where PartNumber='" & Trim(lblModel.Text) & "' and Active=0 order by version "
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

Private Sub cmdToLeft_Click()
    If lstUnuse.ListIndex >= 0 Then
        
        sql = "Update tblECO_Ver set Active=1 where PartNumber='" & Trim(lblModel.Text) & "' and Version='" & lstUnuse.Text & "'"
        status = Connect.excuteUpdate(sql)
     
        If status <> "" Then
        Else
            lstActive.AddItem lstUnuse.Text
            lstUnuse.RemoveItem lstUnuse.ListIndex
        End If
        
   
    End If
End Sub

Private Sub cmdToRight_Click()

    If lstActive.ListIndex >= 0 Then
        
         sql = "Update tblECO_Ver set Active=0 where PartNumber='" & Trim(lblModel.Text) & "' and Version='" & lstActive.Text & "'"
        status = Connect.excuteUpdate(sql)
     
        If status <> "" Then
        Else
            lstUnuse.AddItem lstActive.Text
            lstActive.RemoveItem lstActive.ListIndex
        End If
        
      
    End If

End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
    End If
End Sub
