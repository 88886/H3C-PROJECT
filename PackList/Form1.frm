VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6180
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MainForm.Show
End Sub

Private Sub Command2_Click()
    NXPrint.Show
End Sub
