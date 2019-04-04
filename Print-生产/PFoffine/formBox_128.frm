VERSION 5.00
Begin VB.Form formBox_128 
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   Picture         =   "formBox_128.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   10530
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox tbMAC 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox tbSN 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "MAC"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "SN"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   4920
      Width           =   495
   End
End
Attribute VB_Name = "formBox_128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Function validateMAC()
    Dim sn As String
    If (Len(Trim(tbMAC.Text)) <> 12) Then
        MsgBox "条码长度必修为12位"
        tbMAC.SelStart = 0
        tbMAC.SelLength = Len(tbMAC.Text)
        validateMAC = False
    Else
        validateMAC = True
    End If
End Function

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Dim Ctr As Control
    For Each Ctr In Me.Controls
        If TypeOf Ctr Is TextBox Then
            Ctr.Text = ""
        End If
    Next
End Sub

Private Sub cmdPrint_Click()
    Dim sn As String, mac As String
    sn = Trim(tbSN.Text)
    mac = Trim(tbMAC.Text)
    If (Len(sn) <> 11 Or Len(mac) <> 12) Then
        MsgBox "条码长度必修是11，MAC长度必须是12"
        Exit Sub
    End If
    
    Dim File As String
    Dim Dic As New dictionary
    
    File = "\\sz-fs01\Labels\Packet Front\New folder\GP-MPC480-MB Shipping Label.btw"
    Call Dic.PutValue("MAC", mac)
    Call Dic.PutValue("SN", sn)
        
    Call Printer.PrintLabel(File, Dic)
    
    
    tbSN.Text = ""
    tbMAC.Text = ""
    tbSN.SetFocus
    
End Sub

Private Sub Command1_Click()
    mac = Connect.getMACOfPackFront("F84613O1703")
    MsgBox mac
End Sub

Private Sub tbMAC_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (validateMAC = False) Then
            Exit Sub
        End If
        cmdPrint_Click
    End If
End Sub

Private Function validateSN()
    Dim sn As String
    If (Len(Trim(tbSN.Text)) <> 11) Then
        MsgBox "条码长度必修为11位"
        tbSN.SelStart = 0
        tbSN.SelLength = Len(tbSN.Text)
        validateSN = False
    Else
        tbMAC.Text = ""
        validateSN = True
    End If
End Function
Private Sub tbSN_KeyPress(KeyAscii As Integer)
    Dim mac As String
    If (KeyAscii = 13) Then
        If (validateSN = True) Then
            mac = getMACOfPackFront(tbSN.Text)
            tbMAC.Text = mac
            If (Len(mac) = 12 And chkPrint.value = 1) Then
                cmdPrint_Click
            End If
        End If

    End If
End Sub
