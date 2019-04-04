VERSION 5.00
Begin VB.Form frmEANPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EAN���ǩ��ӡ"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEANPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ӡ(Print) "
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   3960
      Width           =   6495
      Begin VB.CheckBox chkPA 
         BackColor       =   &H80000009&
         Caption         =   "Check1"
         Height          =   330
         Left            =   3720
         TabIndex        =   16
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox txtEAN 
         Enabled         =   0   'False
         Height          =   450
         Left            =   2280
         TabIndex        =   13
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2280
         TabIndex        =   8
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtCPN 
         Height          =   405
         Left            =   2280
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "�Ƿ��ӡ3COM��ַ"
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "EAN����"
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ӡ����:"
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   0
      Picture         =   "frmEANPrint.frx":13652
      ScaleHeight     =   3945
      ScaleWidth      =   6465
      TabIndex        =   1
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��Ʒ����:"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EAN Numwe"
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��Ʒ����:"
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��Ʒ����:"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
End
Attribute VB_Name = "frmEANPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects




Private Sub cmdCancel_Click()
   txtCPN.Text = ""
   txtQty.Text = ""
   txtEAN.Text = ""
   chkPA.Value = 0
   txtCPN.SetFocus
End Sub

Private Sub cmdPrint_Click()
   If txtCPN.Text = "" Then
      MsgBox "��Ʒ����δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ�����Ʒ����"
      txtCPN.SetFocus
      Exit Sub
   End If
   
   If txtQty.Text = "" Then
      MsgBox "����δ���룬���ܴ�ӡ��", vbInformation + vbOKOnly, "δ��������"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If txtEAN.Text = "" Then
      MsgBox "ϵͳδȡ��EAN���룬���ܴ�ӡ�����������Ʒ�����س���ȡEAN����", vbInformation + vbOKOnly, "δ��������"
      txtQty.SetFocus
      Exit Sub
   End If
      
   If CInt(txtQty.Text) = 0 Then
      MsgBox "��������ȷ��������", vbInformation + vbOKOnly, "��������"
      txtQty.SetFocus
      Exit Sub
   End If
   

   

   
   

   

   Dim i, qty As Integer
   'Dim leftstr, rightstr, str As String
   qty = CInt(txtQty.Text)
     OpenLppx
     
   For i = 0 To qty - 1

      

 
   myVars.Item("sn").Value = txtEAN.Text

   myVars.Item("pn").Value = txtCPN.Text
   
 
  
   'myApp.Visible = True
  
   myDoc.PrintLabel 1
   myDoc.FormFeed

   Next
   
   UnloadLppx
    

   cmdCancel_Click
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub







Private Sub Form_Load()
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub




Private Sub lblVer_Click()

End Sub

Private Sub txtCPN_KeyPress(KeyAscii As Integer)
     If (KeyAscii = 13) Then
        sql = "select * from ean where SN='" & Mid(txtCPN.Text, 1, 8) & "'"
        'Print sql
        rec.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rec.EOF = True Then
         MsgBox "�˲�ƷEAN Numberδ��������!"
         txtCPN.Text = ""
         txtCPN.SetFocus
         rec.Close
         Exit Sub
      Else
        txtCPN.Text = rec.Fields(1)
        txtEAN.Text = rec.Fields(2)
        If rec.Fields(3) = "Y" Then
            chkPA.Value = 1
        End If

    
    
      End If
      rec.Close
      txtQty.SetFocus
   End If

End Sub

Private Sub txtQty_Change()
If txtQty.Text <> "" Then
    If Asc(Right(txtQty.Text, 1)) > 57 Or Asc(Right(txtQty.Text, 1)) < 48 Then
       MsgBox "ֻ���������֣�", vbInformation + vbOKOnly, "���벻��ȷ"
       SendKeys "{backspace}"
       txtQty.SetFocus
       Exit Sub
    End If
End If
End Sub

'Private Sub txtQty_KeyPress(KeyAscii As Integer)
 ' If (KeyAscii = 13) Then
  '   txtVer.SetFocus
  'End If
'End Sub


Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   If chkPA.Value = 0 Then
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "EAN_NA.lab")
   Else
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "EAN.lab")
   End If
   'Set myDoc = myApp.Documents.Open("G:\flash\��ǩģ��\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub


