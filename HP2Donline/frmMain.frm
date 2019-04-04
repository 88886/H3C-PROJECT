VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Main Form"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7815
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
   ScaleHeight     =   4800
   ScaleWidth      =   7815
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdH3C 
      Caption         =   "H3C 2D 标签打印"
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退  出(Exit)"
      Height          =   735
      Left            =   5760
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "2013.03.20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Image imgH3C 
      Height          =   810
      Left            =   960
      Picture         =   "frmMain.frx":13652
      Top             =   1560
      Width           =   1785
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "2D 标签打印选择(Label Printed Select)"
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
      TabIndex        =   2
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
   If conn1.State = 1 Then
      conn1.Close
   Set conn1 = Nothing
   End If
   End
End Sub


Private Sub cmdH3C_Click()
   frmH3CPrint.Show
End Sub


