VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Main Form"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8805
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8805
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   615
      Left            =   600
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "EAN类3COM"
      Height          =   735
      Left            =   3240
      TabIndex        =   11
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "0302"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "0303"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "0305"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "H3C"
      Height          =   375
      Left            =   5760
      MaskColor       =   &H000080FF&
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "非H3C"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "非3COM整机类"
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C 
      Caption         =   "非3COM单板类"
      Height          =   855
      Left            =   3240
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdHUAWEI 
      Caption         =   "3C类3COM"
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C_3COM 
      Caption         =   "21类3COM"
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退  出(Exit)"
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   5400
      X2              =   5760
      Y1              =   1920
      Y2              =   1560
   End
   Begin VB.Line Line6 
      X1              =   5400
      X2              =   5760
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line5 
      X1              =   5400
      X2              =   5760
      Y1              =   1920
      Y2              =   2160
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   360
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   5400
      X2              =   5760
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   5400
      X2              =   5760
      Y1              =   3120
      Y2              =   2880
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "标签打印选择(Label Printed Select)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   End
End Sub

Private Sub cmdH3C_3COM_Click()
   frmH3C_3COMPrint.Show
End Sub

Private Sub cmdHUAWEI_Click()
   frmH3COMPrint.Show
End Sub



Private Sub Command1_Click()
'frm21Print.Show
End Sub

Private Sub Command2_Click()
frm21huaweiPrint.Show
End Sub

Private Sub Command3_Click()
frm21H3CPrint.Show
End Sub

Private Sub Command4_Click()
frm0305Print.Show
End Sub

Private Sub Command5_Click()
frm0303Print.Show
End Sub

Private Sub Command6_Click()
frm0302Print.Show
End Sub

Private Sub Command7_Click()
frmEANPrint.Show
End Sub

Private Sub Command8_Click()
frmTest.Show
End Sub
