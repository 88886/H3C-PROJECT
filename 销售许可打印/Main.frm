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
   Begin VB.CommandButton Command2 
      Caption         =   "H3C销售许可打印（TASK）"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   2880
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "H3C销售许可打印"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    frmSalesLicense.Show

End Sub

Private Sub Command2_Click()
     frmSalesLicense_TASK.Show
End Sub
