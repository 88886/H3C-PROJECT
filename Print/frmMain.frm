VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Main Form"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12540
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
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12540
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command28 
      BackColor       =   &H00C0E0FF&
      Caption         =   "H3C2D补印"
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command27 
      Caption         =   "H3C02122D"
      Height          =   375
      Left            =   10320
      TabIndex        =   32
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H00C0FFC0&
      Caption         =   "0303二维码8X8打印"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5160
      Width           =   3135
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00C0C0C0&
      Caption         =   "H3C2D"
      Height          =   375
      Left            =   10320
      MaskColor       =   &H000080FF&
      TabIndex        =   30
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00C0FFC0&
      Caption         =   "0302二维码8X8打印"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00C0E0FF&
      Caption         =   "GW MAC地址补印"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7320
      Width           =   2895
   End
   Begin VB.CommandButton cmdMAC 
      BackColor       =   &H00C0E0FF&
      Caption         =   "MAC地址补印"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton Command10 
      Caption         =   "新MAC地址打印"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton Command22 
      Caption         =   "GW MAC地址打印"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6720
      Width           =   2895
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0E0FF&
      Caption         =   "二维码MAC地址补印"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton Command20 
      Caption         =   "二维码MAC地址打印"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "H3C"
      Height          =   375
      Left            =   10320
      MaskColor       =   &H000080FF&
      TabIndex        =   22
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command18 
      Caption         =   "非H3C"
      Height          =   375
      Left            =   10320
      TabIndex        =   21
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command17 
      Caption         =   "H3C 整机类"
      Height          =   615
      Left            =   7200
      TabIndex        =   20
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0FFC0&
      Caption         =   "0303二维码打印"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton Command15 
      Caption         =   "0303一维码打印"
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton Command14 
      Caption         =   "0302一维码打印"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0FFC0&
      Caption         =   "0302二维码打印"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command12 
      Caption         =   "H3C 单板类"
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "0303二维码打印"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "0302二维码打印"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "HP本体标签"
      Height          =   735
      Left            =   720
      TabIndex        =   12
      Top             =   7080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "EAN类3COM"
      Height          =   735
      Left            =   3480
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "0302一维码打印"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "0303一维码打印"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "0305"
      Height          =   375
      Left            =   10080
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "H3C"
      Height          =   375
      Left            =   10560
      MaskColor       =   &H000080FF&
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "非H3C"
      Height          =   375
      Left            =   10560
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HUV 整机类"
      Height          =   615
      Left            =   7200
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C 
      Caption         =   "HUV 单板类"
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmdHUAWEI 
      Caption         =   "3C类3COM"
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C_3COM 
      Caption         =   "21类3COM"
      Height          =   735
      Left            =   3480
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退  出(Exit)"
      Height          =   735
      Left            =   9480
      TabIndex        =   0
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Line Line18 
      X1              =   10440
      X2              =   9960
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line17 
      X1              =   9960
      X2              =   9960
      Y1              =   2880
      Y2              =   3600
   End
   Begin VB.Line Line16 
      X1              =   9960
      X2              =   10320
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line15 
      X1              =   9960
      X2              =   9960
      Y1              =   1800
      Y2              =   2880
   End
   Begin VB.Line Line14 
      X1              =   9360
      X2              =   10320
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line13 
      X1              =   3240
      X2              =   3840
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line12 
      X1              =   3240
      X2              =   3840
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line11 
      X1              =   3240
      X2              =   3840
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line10 
      X1              =   2760
      X2              =   3840
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line9 
      X1              =   3240
      X2              =   3240
      Y1              =   4800
      Y2              =   3360
   End
   Begin VB.Line Line8 
      X1              =   9360
      X2              =   10560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line3 
      X1              =   3360
      X2              =   3360
      Y1              =   720
      Y2              =   2250
   End
   Begin VB.Line Line7 
      X1              =   2760
      X2              =   3840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line4 
      X1              =   3360
      X2              =   3840
      Y1              =   1250
      Y2              =   1250
   End
   Begin VB.Line Line6 
      X1              =   3360
      X2              =   3840
      Y1              =   1750
      Y2              =   1750
   End
   Begin VB.Line Line5 
      X1              =   3360
      X2              =   3840
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line2 
      X1              =   9960
      X2              =   10560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   9960
      X2              =   9960
      Y1              =   1200
      Y2              =   720
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "H3C标签打印选择(Label Printed Select)"
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
      Width           =   12615
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

Private Sub cmdMAC_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        FrmMACReprint.Show
    End If
End Sub

Private Sub Command1_Click()
'frm21Print.Show
End Sub

Private Sub Command10_Click()
     If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        FrmMACUpgrade.Show
    End If
End Sub

Private Sub Command11_Click()
If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm0303_2D.Show
    End If
    
End Sub



Private Sub Command13_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm022D_H3C.Show
    End If
    
End Sub

Private Sub Command14_Click()
' to do
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm0302Print_H3C.Show
    End If
End Sub

Private Sub Command15_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm0303Print_H3C.Show
    End If
End Sub

Private Sub Command16_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm0303_2D_H3C.Show
    End If
End Sub



Private Sub Command18_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm21huaweiPrint_H3C.Show
    End If
End Sub

Private Sub Command19_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm21H3CPrint_H3C.Show
    End If
End Sub

Private Sub Command2_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm21huaweiPrint.Show
    End If
End Sub

Private Sub Command20_Click()
     If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        Frm2DMACUpgrade.Show
    End If
End Sub

Private Sub Command21_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        Frm2DMACReprint.Show
    End If
End Sub

Private Sub Command22_Click()
    FrmMACforGW.Show
End Sub

Private Sub Command23_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        FrmMACReprintforGW.Show
    End If

End Sub

Private Sub Command24_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm022D_H3C_8X8.Show
    End If
End Sub

Private Sub Command25_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm21H3CPrint_H3C_2D.Show
    End If
End Sub

Private Sub Command26_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm0303_2D_H3C_8X8.Show
    End If
End Sub

Private Sub Command27_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frmNOH3C2DPrint.Show
    End If
    
End Sub

Private Sub Command28_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm21H3CPrint_H3C_2D_Reprint.Show
    End If
End Sub

Private Sub Command3_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
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
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm0303Print.Show
    End If
    
End Sub

Private Sub Command6_Click()
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
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
    If Connect.AccessCheck(golUSERNAME, "offline") = False Then
        MsgBox "您没有该界面的权限!"
        Exit Sub
    Else
        frm022D.Show
    End If
    
End Sub

Private Sub Command9_Click()
frmHPSelfPrint.Show
End Sub

