VERSION 5.00
Begin VB.Form frmInformation 
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "请确认"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   6960
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirm 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lablRight 
      BackColor       =   &H0000FF00&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblLeft 
      BackColor       =   &H0000FF00&
      Caption         =   "ren"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
   result = ""
   Unload Me
End Sub

Private Sub cmdConfirm_Click()
   result = "OK"
   Unload Me
End Sub

Private Sub Form_Load()
   lblLeft.Caption = info
   lablRight.Caption = nver
End Sub

