VERSION 5.00
Begin VB.Form FrmMAC 
   Caption         =   "MAC��ַ��ӡ"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   11490
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   360
      Picture         =   "FrmMAC.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   6915
      TabIndex        =   15
      Top             =   600
      Width           =   6975
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   10935
      Begin VB.TextBox txtMac 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtPart 
         Height          =   405
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtQty 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6720
         TabIndex        =   6
         Text            =   "1"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtWeishu 
         Height          =   450
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtQty2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9240
         TabIndex        =   4
         Text            =   "1"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ʼMac��ַ:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "����:"
         Height          =   375
         Left            =   5880
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Mac��ַλ��:"
         Height          =   375
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "һʽ����:"
         Height          =   375
         Left            =   7920
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   615
      Left            =   7680
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ӡ(Print) &p"
      Height          =   615
      Left            =   2760
      Picture         =   "FrmMAC.frx":1A66A
      TabIndex        =   0
      Top             =   4560
      Width           =   1815
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
      TabIndex        =   14
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "FrmMAC"
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
Dim isPrint As String

Function sums(ByVal X As String, ByVal Y As String) As String ' sum of two hugehexnum����������֮�ͣ�
Dim max As Long, temp As Long, i As Long, result As Variant
max = IIf(Len(X) >= Len(Y), Len(X), Len(Y))
X = Right(String(max, "0") & X, max)
Y = Right(String(max, "0") & Y, max)
ReDim result(0 To max)
For i = max To 1 Step -1
result(i) = Val(Mid(X, i, 1)) + Val(Mid(Y, i, 1))
Next
For i = max To 1 Step -1
temp = result(i) \ 10
result(i) = result(i) Mod 10
result(i - 1) = result(i - 1) + temp
Next
If result(0) = 0 Then result(0) = ""
sums = Join(result, "")
Erase result

End Function

Function multi(ByVal X As String, ByVal Y As String) As String 'multi of two huge hexnum����������֮����
Dim result As Variant
Dim xl As Long, yl As Long, temp As Long, i As Long
xl = Len(Trim(X))
yl = Len(Trim(Y))

ReDim result(1 To xl + yl)
For i = 1 To xl
For temp = 1 To yl
result(i + temp) = result(i + temp) + Val(Mid(X, i, 1)) * Val(Mid(Y, temp, 1))
Next
Next

For i = xl + yl To 2 Step -1
temp = result(i) \ 10
result(i) = result(i) Mod 10
result(i - 1) = result(i - 1) + temp
Next

If result(1) = "0" Then result(1) = ""
multi = Join(result, "")
Erase result

End Function
Function POWERS(ByVal X As Integer) As String ' GET 16777216^X,ie 16^(6*x)��16777216��X �η���
POWERS = 1
Dim i As Integer
For i = 1 To X
POWERS = multi(POWERS, CLng(&H1000000))
Next
End Function
Function half(ByVal X As String) As String 'get half of x��ȡ�룩
X = 0 & X
Dim i As Long
Dim result As Variant
ReDim result(2 To Len(X)) As String
For i = 2 To Len(X)
result(i) = CStr(Val(Mid(X, i, 1)) \ 2 + IIf(Val(Mid(X, i - 1, 1)) Mod 2 = 1, 5, 0))
Next
half = Join(result, "")
If Left(half, 1) = "0" Then half = Right(half, Len(half) - 1) ' no zero ahead
End Function


'��һ�����õĺ�����
Function POWERXY(ByVal X As Integer, ByVal Y As Integer) As String 'GET X^Y��X �� Y �η���
Dim i As Integer
POWERXY = X
For i = 2 To Y
POWERXY = multi(POWERXY, X)
Next
End Function

'����ת��������


'16 to 10
Function HEXTODEC(ByVal X As String) As String
Dim A() As String, i As Long, UNIT As Integer
For i = 1 To Len(X)
If Not IsNumeric("&h" & Mid(X, i, 1)) Then MsgBox "NOT A HEX FORMAT!", 64, "INFO": Exit Function
Next
X = String((6 - Len(X) Mod 6) Mod 6, "0") & X

UNIT = Len(X) \ 6 - 1
ReDim A(UNIT)
For i = 0 To UNIT
A(i) = CLng("&h" & Mid(X, i * 6 + 1, 6))
Next
For i = 0 To UNIT
A(i) = multi(A(i), POWERS(UNIT - i))
HEXTODEC = sums(HEXTODEC, A(i))
Next
End Function




' 10 to 16
Function dectohex(ByVal hugenum As String) As String ' trans hugenum to hex

Do While Len(hugenum) > 2
dectohex = Hex(Val(Right(hugenum, 4)) Mod 16) & dectohex
For i = 1 To 4 'devide hugenum by 16
hugenum = half(hugenum)
Next
Loop
Dim tmp As String
Dim k As Integer

tmp = Hex(Val(hugenum)) & dectohex
For k = 1 To 12
    If Len(tmp) < 12 Then
        tmp = "0" & tmp
    Else
        Exit For
    End If
Next
dectohex = tmp
End Function


Private Sub cmdCancel_Click()
   txtPart.Text = ""
   txtMac.Text = ""
   txtQty.Text = ""
   txtWeishu.Text = ""
   txtPart.SetFocus
End Sub

Private Sub cmdPrint_Click()
''   If txtPart.Text = "" Then
''      MsgBox "��Ʒ����δ����,���ܴ�ӡ,�����������Ʒ����!", vbInformation + vbOKOnly, "δ��������"
''      txtSN.SetFocus
''      Exit Sub
''   End If
   If txtQty.Text = "" Then
      MsgBox "δ����������", vbInformation + vbOKOnly, "δ��������"
      txtQty.SetFocus
      Exit Sub
   End If
   If txtQty2.Text = "" Then
      MsgBox "δ����һʽ����������", vbInformation + vbOKOnly, "δ����һʽ��������"
      txtQty2.SetFocus
      Exit Sub
   End If
    If txtWeishu.Text = "" Then
      MsgBox "δ����λ����", vbInformation + vbOKOnly, "δ����λ��"
      'txtQty.SetFocus
      Exit Sub
   End If
   
   If CInt(txtQty.Text) = 0 Then
      MsgBox "��������ȷ��������", vbInformation + vbOKOnly, "��������"
      txtQty.SetFocus
      Exit Sub
   End If
   If CInt(txtQty2.Text) = 0 Then
      MsgBox "��������ȷ��һʽ����������", vbInformation + vbOKOnly, "һʽ������������"
      txtQty2.SetFocus
      Exit Sub
   End If
   
   If CInt(txtWeishu.Text) = 0 Then
      MsgBox "��������ȷ��λ����", vbInformation + vbOKOnly, "λ������"
      txtWeishu.SetFocus
      Exit Sub
   End If
   
   
   
   Dim sn, mac, ip, model, part As String
   Dim weishu, qty, qty2 As Integer
   part = UCase(txtPart.Text)
   mac = UCase(txtMac.Text)
   weishu = CInt(txtWeishu.Text)
   qty = CInt(txtQty.Text)
   qty2 = CInt(txtQty2.Text)

   '��ʼ����
    Dim arr() As Double
    Dim number2 As Integer
    ReDim Preserve arr(qty) As Double
    
    arr(0) = HEXTODEC(mac)
    For i = 1 To qty - 1
     arr(i) = arr(i - 1) + weishu
     'MsgBox arr(I)
    Next
    
    OpenLppx
    Dim j As Integer
    For i = 0 To qty - 1
        For j = 0 To qty2 - 1
          myVars.Item("MAC").Value = dectohex(arr(i))
          myApp.Visible = False
          myDoc.PrintLabel 1
          myDoc.FormFeed
        Next
    Next

   UnloadLppx
   cmdCancel_Click
End Sub

Private Sub cmdReturn_Click()
    If (conn.State <> 1) Then
        conn.Close
    End If
    Unload Me
End Sub


Private Sub Form_Load()
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
End Sub

Private Sub txtMac_Change()
    If txtMac.Text <> "" Then
        If (Asc(Right(txtMac.Text, 1)) <= 57 And Asc(Right(txtMac.Text, 1)) >= 48) Or (Asc(Right(txtMac.Text, 1)) >= 65 And Asc(Right(txtMac.Text, 1)) <= 70) Then
        Else
           MsgBox "ֻ���������ֻ����ַ�ABCDEF��", vbInformation + vbOKOnly, "���벻��ȷ"
           'SendKeys "{backspace}"
           txtMac.SetFocus
           Exit Sub
        End If
    End If
End Sub

Private Sub txtMac_GotFocus()
    Me.txtMac.Text = Clipboard.GetText
    If (Len(txtMac.Text) <> 12) Then
       MsgBox "MAC��ַ��Ϊ12λ���޸�!"
    End If
    Clipboard.Clear
End Sub

Private Sub txtMac_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (Len(txtMac.Text) <> 12) Then
            MsgBox "MAC��ַ��Ϊ12λ���޸�!"
            Me.txtMac.SetFocus
        Else
            Me.txtWeishu.SetFocus
        End If
    End If
End Sub

Private Sub txtPart_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
      If Len(Trim(txtPart.Text)) <> 8 Then
         MsgBox "��Ʒ��ų��ȱ���Ϊ8!"
         txtPart.SetFocus
         Exit Sub
      End If
        
      sql = "select * from H3COM where Part='" & UCase(Trim(txtPart.Text)) & "'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = False Then
         MsgBox "3COM���֣����ܵ�����ӡMAC��ַ!"
         txtPart.Text = ""
         txtQty.Text = ""
         rec.Close
         Exit Sub
      Else
        rec.Close
        txtMac.SetFocus
      End If
      
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

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
     If txtQty.Text = "" Then
       MsgBox "��ӡ��������Ϊ�գ�"
       txtQty.SetFocus
     Else
       txtQty2.SetFocus
     End If
 End If
End Sub

Private Sub txtQty2_Change()
If txtQty2.Text <> "" Then
    If Asc(Right(txtQty2.Text, 1)) > 57 Or Asc(Right(txtQty2.Text, 1)) < 48 Then
       MsgBox "ֻ���������֣�", vbInformation + vbOKOnly, "���벻��ȷ"
       SendKeys "{backspace}"
       txtQty2.SetFocus
       Exit Sub
    End If
End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "MAC��ַ.Lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\��ǩģ��\" & "H3C-3COM.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub txtWeishu_Change()
If txtWeishu.Text <> "" Then
    If Asc(Right(txtWeishu.Text, 1)) > 57 Or Asc(Right(txtWeishu.Text, 1)) < 48 Then
       MsgBox "ֻ���������֣�", vbInformation + vbOKOnly, "���벻��ȷ"
       SendKeys "{backspace}"
       txtWeishu.SetFocus
       Exit Sub
    End If
End If
End Sub

Private Sub txtWeishu_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
     If txtWeishu.Text = "" Then
       MsgBox "Macλ������Ϊ�գ�"
       txtWeishu.SetFocus
     Else
       txtQty.SetFocus
     End If
 End If
End Sub

