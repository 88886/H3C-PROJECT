VERSION 5.00
Begin VB.Form frmH3C_3COMPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H3C-3COM Label Print"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmH3C_3COMPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11160
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ӡ(Print) &p"
      Height          =   615
      Left            =   3000
      Picture         =   "frmH3C_3COMPrint.frx":13652
      TabIndex        =   4
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   615
      Left            =   7920
      TabIndex        =   6
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   4680
      Width           =   10935
      Begin VB.TextBox txtQty2 
         Height          =   450
         Left            =   9000
         TabIndex        =   19
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtWeishu 
         Height          =   450
         Left            =   7680
         TabIndex        =   17
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtQty 
         Height          =   450
         Left            =   6360
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtPart 
         Height          =   405
         Left            =   2400
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtModel 
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtMac 
         Height          =   405
         Left            =   7680
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "һʽ���ݣ�"
         Height          =   375
         Left            =   7560
         TabIndex        =   18
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "Mac��ַλ��:"
         Height          =   375
         Left            =   5520
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "������"
         Height          =   375
         Left            =   5520
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ�ͺ�:"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ��ʼ����:"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ʼMac��ַ:"
         Height          =   375
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      Picture         =   "frmH3C_3COMPrint.frx":26CA4
      ScaleHeight     =   4185
      ScaleWidth      =   10905
      TabIndex        =   8
      Top             =   360
      Width           =   10935
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "21-3COM ��ǩ��"
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
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmH3C_3COMPrint"
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

Private Sub cmdCancel_Click()
   txtSN.Text = ""
   txtModel.Text = ""
   txtPart.Text = ""
   txtMac.Text = ""
   txtQty.Text = ""
   txtWeishu.Text = ""
   txtSN.SetFocus
End Sub

Private Sub cmdPrint_Click()
   If txtSN.Text = "" Then
      MsgBox "��Ʒ��ʼ����δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ�����Ʒ��ʼ����"
      txtSN.SetFocus
      Exit Sub
   End If
   If txtPart.Text = "" Then
      MsgBox "��Ʒ����δ����,���ܴ�ӡ,�����������Ʒ����!", vbInformation + vbOKOnly, "δ��������"
      txtSN.SetFocus
      Exit Sub
   End If
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
      txtQty.SetFocus
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
   part = UCase(txtModel.Text)
   sn = UCase(txtSN.Text)
   mac = UCase(txtMac.Text)
   weishu = CInt(txtWeishu.Text)
   qty = CInt(txtQty.Text)
   qty2 = CInt(txtQty2.Text)
  ' model = UCase(txtModel.Text)
   '��ʼ����
Dim arr() As Double
Dim number2 As Integer
ReDim Preserve arr(qty) As Double

arr(0) = HEXTODEC(mac)
For i = 1 To qty - 1
 arr(i) = arr(i - 1) + weishu
 'MsgBox arr(I)
Next

Dim ip1, ip2 As Integer
Dim str1, str2, leftstr, rightstr, str As String
Dim arr2() As String
ReDim Preserve arr2(qty, 2) As String
  
leftstr = UCase(Left(txtSN.Text, 14))
rightstr = Right(txtSN.Text, 6)
 OpenLppx
Dim j As Integer
For i = 0 To qty - 1

  str = leftstr & Right("000000" & CStr(CInt(rightstr) + i), 6)   'sn

  arr2(i, 0) = dectohex(arr(i))                                   'mac
  str1 = Mid(arr2(i, 0), Len(arr2(i, 0)) - 3, 2)
  str2 = Mid(arr2(i, 0), Len(arr2(i, 0)) - 1, 2)
  
  ip1 = CInt(Val("&H" & str1))
  ip2 = CInt(Val("&H" & str2))
  
  arr2(i, 1) = "169.254." & CStr(ip1) & "." & CStr(ip2)            'ip
  
 For j = 0 To qty2 - 1
  
   myVars.Item("Sn").Value = str
   myVars.Item("Model").Value = part
   myVars.Item("Mac").Value = arr2(i, 0)
   'myVars.Item("Ip").Value = arr2(i, 1)
   If isPrint = "Y" Or isPrint = "Yes" Then
   myVars.Item("Ip").Value = arr2(i, 1)
   Else
   myVars.Item("Ip").Value = "N/A"
   End If
 
   myApp.Visible = False
   myDoc.PrintLabel 1
   myDoc.FormFeed
Next
     
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
   UnloadLppx
End Sub



Private Sub txtMac_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 13) Then
     If txtMac.Text = "" Then
       MsgBox "Mac��ʼ��ַ����Ϊ�գ�"
       txtMac.SetFocus
     Else
       txtWeishu.SetFocus
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

Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      If Len(Trim(txtSN.Text)) <> 20 Then
         MsgBox "��Ʒ��ų��ȱ���Ϊ20!"
         txtSN.SetFocus
         Exit Sub
      End If
     
     txtPart.Text = Mid(UCase(txtSN.Text), 3, 8)
     
      sql = "select * from H3COM where Part='" & Trim(txtPart.Text) & "' and Class='21'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "�˲�Ʒ����δ��������!"
         txtModel.Text = ""
         txtPart.Text = ""
         txtQty.Text = ""
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
        txtSN.Text = UCase(txtSN.Text)
        
        
        isPrint = rec.Fields(5)
        
        txtModel.Text = rec.Fields("Model")
        txtMac.SetFocus
      rec.Close
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
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "213COM.lab")
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
dectohex = Hex(Val(hugenum)) & dectohex
End Function

