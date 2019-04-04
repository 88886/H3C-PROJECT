VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   Caption         =   "Main"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   6345
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "UINS打印保修卡"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "逆向打印保修卡"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   2
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "01M打印保修卡"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "H3C打印保修卡"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    frmH3CPrintWarranty.Show

End Sub

Private Sub Command2_Click()
   frmH3CPrintWarrantyByPart.Show
End Sub

Private Sub Command3_Click()
   frmNiXiangH3CPrintWarranty.Show
End Sub

Private Sub Command4_Click()
   frmUINSPrintWarranty.Show
End Sub
