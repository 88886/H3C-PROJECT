VERSION 5.00
Begin VB.Form FrmMACReprint 
   Caption         =   "MAC��ַ��ӡ"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   11445
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   615
      Left            =   7560
      TabIndex        =   10
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ӡ(Print) &p"
      Height          =   615
      Left            =   1440
      Picture         =   "FrmMACReprint.frx":0000
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   10935
      Begin VB.TextBox txtMac 
         Height          =   400
         Left            =   2040
         TabIndex        =   11
         Text            =   "txtMac"
         Top             =   500
         Width           =   3015
      End
      Begin VB.TextBox txtQty2 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   9240
         TabIndex        =   4
         Text            =   "1"
         Top             =   500
         Width           =   1455
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   6720
         TabIndex        =   3
         Text            =   "1"
         Top             =   500
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "һʽ����:"
         Height          =   375
         Left            =   7920
         TabIndex        =   7
         Top             =   500
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "����:"
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   500
         Width           =   1215
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ʼMac��ַ:"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   500
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   360
      Picture         =   "FrmMACReprint.frx":13652
      ScaleHeight     =   1275
      ScaleWidth      =   6915
      TabIndex        =   1
      Top             =   600
      Width           =   6975
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC��ַ��ǩ��"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "FrmMACReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim isPrint As String
Private Sub cmdCancel_Click()
   txtMac.Text = ""
   txtMac.SetFocus
End Sub

Private Sub cmdPrint_Click()

    If txtMac.Text = "" Or Len(txtMac.Text) < 12 Then
          MsgBox "δ������ȷ��MAC��", vbInformation + vbOKOnly, "δ������ȷ��MAC"
          txtMac.SetFocus
          Exit Sub
    End If

   If txtQty.Text = "" Then
      MsgBox "δ����������", vbInformation + vbOKOnly, "δ��������"
      Exit Sub
   End If
   If txtQty2.Text = "" Then
      MsgBox "δ����һʽ����������", vbInformation + vbOKOnly, "δ����һʽ��������"
      Exit Sub
   End If
   
   If CInt(txtQty.Text) = 0 Then
      MsgBox "��������ȷ��������", vbInformation + vbOKOnly, "��������"
      Exit Sub
   End If
   If CInt(txtQty2.Text) = 0 Then
      MsgBox "��������ȷ��һʽ����������", vbInformation + vbOKOnly, "һʽ������������"
      Exit Sub
   End If
   
   
'   Dim sn, mac, ip, model, part As String
'   Dim weishu, qty, qty2 As Integer
'   part = UCase(txtPart.Text)
'   mac = UCase(txtMac.Text)
'   weishu = CInt(txtWeishu.Text)
'   qty = CInt(txtQty.Text)
'   qty2 = CInt(txtQty2.Text)
'
'   '��ʼ����
'    Dim arr() As Double
'    Dim number2 As Integer
'    ReDim Preserve arr(qty) As Double
'
'    arr(0) = HEXTODEC(mac)
'    For i = 1 To qty - 1
'     arr(i) = arr(i - 1) + weishu
'     'MsgBox arr(I)
'    Next
'
'    OpenLppx
'    Dim j As Integer
'    For i = 0 To qty - 1
'        For j = 0 To qty2 - 1
'          myVars.Item("MAC").Value = dectohex(arr(i))
'          myApp.Visible = False
'          myDoc.PrintLabel 1
'          myDoc.FormFeed
'        Next
'    Next
    

    OpenLppx
    myVars.Item("MAC").Value = txtMac.Text
    myApp.Visible = False
    myDoc.PrintLabel 1
    myDoc.FormFeed
    UnloadLppx
    cmdCancel_Click
End Sub

Private Sub cmdReturn_Click()
    If (conn.State <> 1) Then
        conn.Close
    End If
    Unload Me
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "MAC��ַ.Lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\��ǩģ��\" & "H3C-3COM.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub


Private Sub Form_Load()
txtMac.Text = ""
txtQty.Text = 1
txtQty2.Text = 1
End Sub
