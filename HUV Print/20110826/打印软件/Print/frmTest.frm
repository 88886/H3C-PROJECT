VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   7650
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function timeGetTime Lib "winmm.dll" () As Long


Private Sub Command1_Click()
    Dim Savetime As Double
    Command1.Enabled = False
    Command1.Caption = "Waitting..."
    
    Text1.Text = "timeGetTime begin"
    Savetime = timeGetTime '记下开始时的时间
    While timeGetTime < Savetime + 5000 '循环等待
        DoEvents '转让控制权，以便让操作系统处理其它的事件。
    Wend
    Text1.Text = "timeGetTime end"
    Command1.Enabled = True
    Command1.Caption = "Print"
End Sub

Private Sub Command2_Click()
    If CDbl(Text2.Text) Mod 20 > 0 Then
        Text3.Text = ">0"
    Else
        Text3.Text = "=0"
    End If
    
    
End Sub
