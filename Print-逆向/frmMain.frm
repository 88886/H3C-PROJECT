VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Main Form"
   ClientHeight    =   7248
   ClientLeft      =   48
   ClientTop       =   408
   ClientWidth     =   8808
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7248
   ScaleWidth      =   8808
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "H3C整机2D打印"
      Height          =   375
      Left            =   5760
      MaskColor       =   &H000080FF&
      TabIndex        =   17
      Top             =   6360
      Width           =   2412
   End
   Begin VB.CommandButton Command12 
      Caption         =   "0212二维码打印"
      Height          =   375
      Left            =   5760
      TabIndex        =   16
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0080FF80&
      Caption         =   "0303二维码"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080FF80&
      Caption         =   "0302二维码"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "H3C批量打印"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      Caption         =   "非H3C批量打印"
      Height          =   405
      Left            =   5760
      TabIndex        =   12
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton Command7 
      Caption         =   "EAN类3COM"
      Height          =   735
      Left            =   360
      TabIndex        =   11
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "0302"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "0303"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "0305"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "H3C"
      Height          =   375
      Left            =   5880
      MaskColor       =   &H000080FF&
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "非H3C"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "非3COM整机类"
      Height          =   615
      Left            =   3360
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C 
      Caption         =   "非3COM单板类"
      Height          =   855
      Left            =   3360
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdHUAWEI 
      Caption         =   "3C类3COM"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C_3COM 
      Caption         =   "21类3COM"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   2880
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
   Begin VB.Line Line10 
      X1              =   5520
      X2              =   5880
      Y1              =   2520
      Y2              =   3360
   End
   Begin VB.Line Line9 
      X1              =   5520
      X2              =   5880
      Y1              =   2520
      Y2              =   2880
   End
   Begin VB.Line Line8 
      X1              =   5520
      X2              =   5760
      Y1              =   4200
      Y2              =   5640
   End
   Begin VB.Line Line7 
      X1              =   5520
      X2              =   5760
      Y1              =   4200
      Y2              =   5040
   End
   Begin VB.Line Line4 
      X1              =   5520
      X2              =   5880
      Y1              =   2520
      Y2              =   1440
   End
   Begin VB.Line Line6 
      X1              =   5520
      X2              =   5880
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line5 
      X1              =   5520
      X2              =   5880
      Y1              =   2520
      Y2              =   2400
   End
   Begin VB.Line Line2 
      X1              =   5520
      X2              =   5880
      Y1              =   4200
      Y2              =   4560
   End
   Begin VB.Line Line1 
      X1              =   5520
      X2              =   5880
      Y1              =   4200
      Y2              =   3960
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

Private Sub Command10_Click()
    frm0302_2D.Show
End Sub

Private Sub Command11_Click()
    frm0303_2D.Show
End Sub

Private Sub Command12_Click()
frm0212huaweiPrint.Show
End Sub

Private Sub Command13_Click()
    If Connect.AccessCheck(golUSERNAME, "Reverse") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm21H3CPrint2DMAC.Show
    End If
    
End Sub

Private Sub Command2_Click()
    frm21huaweiPrint.Show
End Sub

Private Sub Command3_Click()
    If Connect.AccessCheck(golUSERNAME, "Reverse") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm21H3CPrint.Show
    End If
    
End Sub

Private Sub Command4_Click()
frm0305Print.Show
End Sub

Private Sub Command5_Click()
    If Connect.AccessCheck(golUSERNAME, "Reverse") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm0303Print.Show
    End If
End Sub

Private Sub Command6_Click()
    If Connect.AccessCheck(golUSERNAME, "Reverse") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm0302Print.Show
    End If
    
End Sub

Private Sub Command7_Click()
    frmEANPrint.Show
End Sub

Private Sub Command8_Click()
     frmHWBatchPrint.Show
End Sub

Private Sub Command9_Click()
    frmBatchPrint.Show
End Sub
