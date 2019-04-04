VERSION 5.00
Begin VB.Form PrintLabel 
   Caption         =   "打印料号"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "PrintLabel.frx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
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
      Left            =   960
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtPCS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtPN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "PCS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "工单数量:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "包材料号:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "PrintLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim myApp As New LabelManager2.Application
Dim mydoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects

Private Sub cmdCancel_Click()
    txtPN.Text = ""
    txtPCS.Text = ""
End Sub

Private Sub cmdPrint_Click()

        If Len(txtPN.Text) <> 8 Then
            MsgBox "料号必须为8位!"
            txtPN.Text = ""
            txtPN.SetFocus
            Exit Sub
        End If
        
        e = Val(txtPCS.Text)
        If e = 0 Then
            k = MsgBox("此栏只能输入不为0的数字，不包含其他字符！ ", vbExclamation)
            txtPCS.Text = ""
            txtPCS.SetFocus
            Exit Sub
        End If
        
        OpenLppx
        
        Dim i As Integer
        For i = 0 To CInt(txtPCS.Text) - 1
            myVars.Item("SN").Value = Trim(txtPN.Text)
        
            mydoc.PrintLabel 1
            mydoc.FormFeed
        Next
        
        
    
        UnloadLppx
        
        
        clearForm
        
End Sub

Private Sub clearForm()
   
   Me.txtPCS.Text = ""
   Me.txtPN.Text = ""
    
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set mydoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "包材料号.lab")
   Me.MousePointer = vbDefault
   Set myVars = mydoc.Variables
   Set myObjs = mydoc.DocObjects
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

