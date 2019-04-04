VERSION 5.00
Begin VB.Form frmMACIP 
   Caption         =   "Print IP"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   Icon            =   "frmMACIP.frx":0000
   LinkTopic       =   "MACIP"
   MaxButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8175
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdBack 
      Caption         =   "返 回"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   14
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打 印"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtWeidizhi 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtQty 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtWeishu 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox txtMac 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2640
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   0
      Picture         =   "frmMACIP.frx":073E
      ScaleHeight     =   2355
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "请检查位地址是否正确"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "位"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "个"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC尾地址:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC打印数量:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC位数:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblSN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC首地址:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
End
Attribute VB_Name = "frmMACIP"
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

Private Sub cmdBack_Click()
 Unload Me
End Sub

Private Sub cmdCancel_Click()
    txtMac.Text = ""
    txtWeishu.Text = ""
    txtQty.Text = ""
    txtWeidizhi.Text = ""
    cmdPrint.Enabled = False
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

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "MAC与IP.Lab")

   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub


Private Sub cmdPrint_Click()

    If txtWeishu.Text = "" Then
      MsgBox "未输入位数！", vbInformation + vbOKOnly, "未输入位数"
      txtQty.SetFocus
      Exit Sub
    End If
   
    If CInt(txtWeishu.Text) = 0 Then
      MsgBox "请输入正确的位数！", vbInformation + vbOKOnly, "位数不对"
      txtWeishu.SetFocus
      Exit Sub
    End If
    
    If txtQty.Text = "" Then
      MsgBox "未输入数量！", vbInformation + vbOKOnly, "未输入数量"
      txtQty.SetFocus
      Exit Sub
    End If
    
    If CInt(txtQty.Text) = 0 Then
      MsgBox "请输入正确的数量！", vbInformation + vbOKOnly, "数量不对"
      txtQty.SetFocus
      Exit Sub
   End If
    
    
   Dim mac, ip As String
   Dim weishu, qty As Integer
   
   
   mac = UCase(txtMac.Text)
   weishu = CInt(txtWeishu.Text)
   qty = CInt(txtQty.Text)
  
   '开始计算
    Dim arr() As Double
    Dim number2 As Integer
    ReDim Preserve arr(qty) As Double

    arr(0) = HEXTODEC(mac)
    For i = 1 To qty - 1
        arr(i) = arr(i - 1) + weishu
    Next

    Dim ip1, ip2 As Integer
    Dim str1, str2, leftstr, rightstr As String
    Dim arr2() As String
    ReDim Preserve arr2(qty, 2) As String
  
    'leftstr = UCase(Left(txtSN.Text, 14))
    'rightstr = Right(txtSN.Text, 6)
    OpenLppx
    Dim j As Integer
    
    For i = 0 To qty - 1

        arr2(i, 0) = dectohex(arr(i))                                   'mac
        str1 = Mid(arr2(i, 0), Len(arr2(i, 0)) - 3, 2)
        str2 = Mid(arr2(i, 0), Len(arr2(i, 0)) - 1, 2)
  
        ip1 = CInt(Val("&H" & str1))
        ip2 = CInt(Val("&H" & str2))
  
        arr2(i, 1) = "169.254." & CStr(ip1) & "." & CStr(ip2)            'ip
  
        myVars.Item("address").Value = "MAC address:"
        myVars.Item("MAC or IP").Value = arr2(i, 0)
 
        myApp.Visible = False
        myDoc.PrintLabel 1
        myDoc.FormFeed
        
        myVars.Item("address").Value = "IP Default address:"
        myVars.Item("MAC or IP").Value = arr2(i, 1)
 
        myApp.Visible = False
        myDoc.PrintLabel 1
        myDoc.FormFeed
     
     
    Next

   UnloadLppx
   cmdCancel_Click
  
End Sub

Private Sub txtMac_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 13) Then
     If txtMac.Text = "" Then
       MsgBox "Mac起始地址不能为空！"
       txtMac.SetFocus
     Else
       txtWeishu.SetFocus
     End If
 End If
End Sub


Private Sub txtQty_Change()
If txtQty.Text <> "" Then
    If Asc(Right(txtQty.Text, 1)) > 57 Or Asc(Right(txtQty.Text, 1)) < 48 Then
       MsgBox "只能输入数字！", vbInformation + vbOKOnly, "输入不正确"
       SendKeys "{backspace}"
       txtQty.SetFocus
       Exit Sub
    End If
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
     If txtQty.Text = "" Then
       MsgBox "打印数量不能为空！"
       txtQty.SetFocus
     Else

       Dim mac, ip As String
       Dim weishu, qty As Integer
       
       mac = UCase(txtMac.Text)
       weishu = CInt(txtWeishu.Text)
       qty = CInt(txtQty.Text)
       
        Dim arr() As Double
        Dim number2 As Integer
        ReDim Preserve arr(qty) As Double

        arr(0) = HEXTODEC(mac)
        
        
        For i = 1 To qty - 1
            arr(i) = arr(i - 1) + weishu
        Next
        
        txtWeidizhi.Text = CStr(dectohex(arr(qty - 1)))
        
        cmdPrint.Enabled = True
     End If
 End If
End Sub

Private Sub txtWeishu_Change()
If txtWeishu.Text <> "" Then
    If Asc(Right(txtWeishu.Text, 1)) > 57 Or Asc(Right(txtWeishu.Text, 1)) < 48 Then
       MsgBox "只能输入数字！", vbInformation + vbOKOnly, "输入不正确"
       SendKeys "{backspace}"
       txtWeishu.SetFocus
       Exit Sub
    End If
End If
End Sub

Private Sub txtWeishu_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
     If txtWeishu.Text = "" Then
       MsgBox "Mac位数不能为空！"
       txtWeishu.SetFocus
     Else
       txtQty.SetFocus
     End If
 End If
End Sub


Function sums(ByVal X As String, ByVal Y As String) As String ' sum of two hugehexnum（两个大数之和）
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

Function multi(ByVal X As String, ByVal Y As String) As String 'multi of two huge hexnum（两个大数之积）
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
Function POWERS(ByVal X As Integer) As String ' GET 16777216^X,ie 16^(6*x)（16777216的X 次方）
POWERS = 1
Dim i As Integer
For i = 1 To X
POWERS = multi(POWERS, CLng(&H1000000))
Next
End Function
Function half(ByVal X As String) As String 'get half of x（取半）
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


'另一个有用的函数：
Function POWERXY(ByVal X As Integer, ByVal Y As Integer) As String 'GET X^Y（X 的 Y 次方）
Dim i As Integer
POWERXY = X
For i = 2 To Y
POWERXY = multi(POWERXY, X)
Next
End Function

'进制转换函数：


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


