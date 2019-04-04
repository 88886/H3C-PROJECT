VERSION 5.00
Begin VB.Form frmHPSNAndMac 
   Caption         =   "HP5020"
   ClientHeight    =   10935
   ClientLeft      =   3750
   ClientTop       =   210
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   13245
   Begin VB.TextBox tempText 
      Height          =   375
      Left            =   8760
      TabIndex        =   47
      Top             =   8880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtWeishu 
      Height          =   375
      Left            =   6480
      TabIndex        =   46
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox txtMACEnd3 
      Height          =   375
      Left            =   7560
      TabIndex        =   44
      Top             =   8040
      Width           =   1815
   End
   Begin VB.TextBox txtMACEnd2 
      Height          =   375
      Left            =   7560
      TabIndex        =   43
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox txtMacStart3 
      Height          =   375
      Left            =   3360
      TabIndex        =   40
      Top             =   8040
      Width           =   1815
   End
   Begin VB.TextBox txtMacStart2 
      Height          =   375
      Left            =   3360
      TabIndex        =   39
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox txtQty 
      Height          =   375
      Left            =   5760
      TabIndex        =   36
      Top             =   9360
      Width           =   1095
   End
   Begin VB.TextBox txtQty1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   34
      Text            =   "1"
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txtMACEnd 
      Height          =   375
      Left            =   7560
      TabIndex        =   32
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox txtMacStart 
      Height          =   375
      Left            =   3360
      TabIndex        =   30
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox txtOrder 
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   8640
      Width           =   1695
   End
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   3000
      TabIndex        =   26
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CheckBox chkN 
      Caption         =   "N*"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   24
      Top             =   5640
      Width           =   615
   End
   Begin VB.CheckBox chkN4 
      Caption         =   "N4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   23
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtModel 
      Height          =   285
      Left            =   8760
      TabIndex        =   22
      Top             =   11520
      Width           =   495
   End
   Begin VB.CheckBox chkY2 
      Caption         =   "Y2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   5640
      Width           =   615
   End
   Begin VB.CheckBox chkY 
      Caption         =   "Y*"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtRevision 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtPart 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txtWorkOrder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtSN 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox txtProduct 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox txtDesc1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox txtDesc2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CommandButton cmdPrint_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "打 印"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   10320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "取 消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   10320
      Width           =   1095
   End
   Begin VB.CommandButton cmdReturn_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "返 回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   10320
      Width           =   1095
   End
   Begin VB.CommandButton cmdMPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "批量打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   10320
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2100
      Left            =   1680
      Picture         =   "frmHPSN&MAC.frx":0000
      ScaleHeight     =   2040
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label19 
      Caption         =   "MAC位数:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   45
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "MAC结束地址:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   42
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "MAC结束地址:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   41
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "MAC起始地址:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   38
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "MAC起始地址:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   37
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "打印数量:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   35
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "一式几份:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   33
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "MAC结束地址:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   31
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "工单号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   28
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "MAC起始地址:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   27
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "MAC打印单号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   25
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000004&
      Caption         =   "环保属性:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "版本:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "机种:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "工单:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   5040
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8400
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "产品序列号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "产品编号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "产品描述1:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      Caption         =   "产品描述2:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
End
Attribute VB_Name = "frmHPSNAndMac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New Recordset
Dim bom_code As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim rs As New Recordset
Dim newLableFlag As Boolean
Dim lTESettingAssignMAC As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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

Private Sub cmdCancel_HPSN_Click()
    Me.txtSN.Text = ""
    Me.txtProduct.Text = ""
    Me.txtDesc1.Text = ""
    Me.txtDesc2.Text = ""
    Me.txtWorkOrder.Text = ""
    Me.txtPart.Text = ""
    Me.txtRevision.Text = ""
    Me.chkY.Value = 0
    Me.chkY2.Value = 0
    Me.chkN.Value = 0
    Me.chkN4.Value = 0
End Sub

Private Sub cmdMPrint_Click()
    Dim model As String
    Dim mac As String
    Dim printMac As String
    Dim qty As Integer
    Dim arr() As Double
    qty = CInt(txtQty.Text)
    
   ReDim Preserve arr(qty) As Double
    
   mac = UCase(txtMacStart.Text)
   weishu = CInt(txtWeishu.Text)
   If lTESettingAssignMAC = True Then
        arr(0) = HEXTODEC(mac)
    For i = 0 To qty - 1
     If i > 0 Then
         arr(i) = arr(i - 1) + weishu
         End If
         
    Next
    
    End If
    
    If Me.txtPart.Text = "" And txtModel.Text = "" Then
        MsgBox "该机种信息不能打印!"
        Exit Sub
    ElseIf Me.txtPart.Text <> "" Then
        model = Mid(txtPart.Text, 4, 8)
    ElseIf txtModel.Text <> "" Then
        model = Trim(txtModel.Text)
    End If
    

 If Connect.checkPrintPreCondition(model, 3) = False Then
    MsgBox "该机种没有在HP序列号类型维护为[50*20]打印选项!"
    Exit Sub
End If
 cmdReturn_HPSN.Enabled = False
'cmdPrint_HPSN.Enabled = False
cmdCancel_HPSN.Enabled = False
sql = "select ITEM_CODE,BARCODE from tblHP_Print where isnull(BARCODE,'')<>'' and isnull(ITEM_CODE,'')<>'' order by BARCODE"
If conn1.State = 0 Then
    conn1.Open
End If
rs.Open sql, conn1, adOpenStatic, adLockReadOnly
If rs.EOF = True Then
    MsgBox ("序列号未导入！")
    rs.Close
    cmdReturn_HPSN.Enabled = True
    'cmdPrint_HPSN.Enabled = True
    cmdCancel_HPSN.Enabled = True
    Exit Sub
Else
    For i = 1 To rs.RecordCount
    
        txtSN.Text = rs("BARCODE")
        txtModel.Text = rs("ITEM_CODE")
        'begin
        If Len(txtSN.Text) < 10 Then
            MsgBox "产品序号长度不能小于10!"
            txtSN.SetFocus
            Exit Sub
        End If
        If InStr(1, txtPart.Text, txtModel.Text) <= 0 Then
            MsgBox ("该工单料号和条码对应的料号不一致，请确认输入工单是否正确!")
            rs.Close
            Exit Sub
        End If
        updateHPInformation
        tempText.Text = dectohex(arr(i - 1))
        cmdPrint_HPSN_Click
        rs.MoveNext
        If i Mod 100 = 0 Then
            Sleep (1000 * 10)
       End If
    Next
    UnloadLppx
    cmdCancel_HPSN_Click
    rs.Close
End If
'del_excel
del_sql
cmdReturn_HPSN.Enabled = True
'cmdPrint_HPSN.Enabled = True
cmdCancel_HPSN.Enabled = True
'MsgBox ("批量打印成功！")
End Sub

Private Sub cmdPrint_HPSN_Click()
    Dim Pb As String
    
    If txtSN.Text = "" Then
        MsgBox ("序列号未输入，不能打印！")
        txtSN.SetFocus
        Exit Sub
    End If
    If txtProduct.Text = "" Then
        MsgBox ("产品编码未带出，不能打印！")
        Exit Sub
    End If
    If txtDesc1.Text = "" Then
        MsgBox ("产品描述1未带出，不能打印！")
        Exit Sub
    End If
     If txtModel.Text = "" Then
        MsgBox ("导入资料中ITEM_CODE栏不能为空！")
        Exit Sub
    End If
    OpenLppx
    myVars.Item("SN").Value = UCase(txtSN.Text)
    myVars.Item("PN").Value = UCase(txtProduct.Text)
    myVars.Item("Model").Value = UCase(txtModel.Text)
    myVars.Item("Rev").Value = UCase(txtRevision.Text)
    myVars.Item("MAC").Value = UCase(tempText.Text)
    If (Me.chkY.Value = 1) Then
        Pb = "Y*"
        OverridePb Pb
        myVars.Item("Rohs").Value = Pb
    ElseIf (Me.chkY2.Value = 1) Then
        Pb = "Y2"
        OverridePb Pb
        myVars.Item("Rohs").Value = Pb
    ElseIf (Me.chkN.Value = 1) Then
        Pb = "N*"
        OverridePb Pb
        myVars.Item("Rohs").Value = Pb
    ElseIf (Me.chkN4.Value = 1) Then
        Pb = "N4"
        OverridePb Pb
        myVars.Item("Rohs").Value = Pb
    Else
        MsgBox "环保属性未选择，不能打印!"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    
    If UploadH3CInfo(Pb, Trim(UCase(txtSN.Text)), Trim(UCase(txtRevision.Text)), "NA", "N/A", "CHINA", "frmHPSNAndMac") = False Then
        MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
    If UploadH3C_PB(Pb, Trim(UCase(txtSN.Text)), Trim(UCase(txtRevision.Text)), "NA", "N/A", "CHINA", "frmHPSNAndMac") = False Then
        MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    
    '======Add by mike 2015.3.24 for data upload to FTPC============
    
    'add by robin for HP SNAndMacPrint
    
     Call Connect.addPrintedMACAndSN(golUSERNAME, Trim(UCase(txtSN.Text)), UCase(tempText.Text), Trim(txtID.Text), txtMacStart.Text, txtMACEnd.Text, "", "")
    'add by robin end
    
    
    If txtDesc2.Text <> "" Then
        myVars.Item("ID-1").Value = txtDesc1.Text
        myVars.Item("ID-2").Value = txtDesc2.Text
    Else
        myVars.Item("ID-1").Value = txtDesc1.Text
        myVars.Item("ID-2").Value = ""
    End If
    'OpenLppx
    myDoc.PrintLabel 1
    myDoc.FormFeed
End Sub

Private Sub OverridePb(ByRef Pb As String)
    Dim labelHistory As New Label_History
    Dim sn As String
    sn = txtSN.Text
    If labelHistory.Init(sn) Then
        Pb = labelHistory.Pb
    End If
End Sub

Private Sub cmdReturn_HPSN_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    If conn1.State = 0 Then
      conn1.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
      conn1.Open
   End If
   If connFTPC.State = 0 Then
        connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
        connFTPC.Open
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   
    If conn1.State = 1 Then
      conn1.Close
      Set conn1 = Nothing
   End If
   
   If connFTPC.State = 1 Then
        connFTPC.Close
        Set connFTPC = Nothing
   End If
End Sub

Private Sub updateHPInformation()

      sql = "select * from hp where hp_sn_iii=substring('" & Trim(txtSN.Text) & "',5,3) and h3c_bom_code = '" + txtModel.Text + "'"
      If rec.State = 1 Then
        rec.Close
      End If
      
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      
      If rec.EOF = True Then
          MsgBox "此序列号未维护信息!"
          txtSN.Text = ""
          txtSN.SetFocus
          rec.Close
          Exit Sub
      Else
          If IsNull(rec.Fields("hpsnproduct")) Then
              MsgBox ("此序列号未维护产品编码!")
              rec.Close
              Exit Sub
          Else
              txtProduct = rec.Fields("hpsnproduct")
          End If
'          hpsnproduct
    
          If IsNull(rec.Fields("hp_desc1")) Then
              MsgBox ("此序列号未维护描述信息!")
              rec.Close
              Exit Sub
          Else
              txtDesc1 = rec.Fields("hp_desc1")
          End If
    
          If Not IsNull(rec.Fields("hp_desc2")) Then
              txtDesc2 = rec.Fields("hp_desc2")
          End If
          
          If IsNull(rec.Fields("new_label")) Or Trim(rec.Fields("new_label")) = "" Then
              newLableFlag = False
              MsgBox ("此机种序列号没有维护new_label选项，请联系ME!")
              rec.Close
              Exit Sub
          Else
              newLableFlag = True
          End If
      End If
    
End Sub
'Private Sub txtSN_KeyPress(KeyAscii As Integer)
'
'End Sub



Private Sub Text1_Change()
If KeyAscii = 13 Then
        If rec.State = 1 Then
            rec.Close
        End If
        sql = "select a.*,(select PrintAllMAC from MAC_PART with(NOLOCK) where PART=a.PART) PrintAllMAC from mac_record a with(NOLOCK) where a.ID = '" & Trim(Me.txtID.Text) & "'  "
        If connFTPC.State = 0 Then
           connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
           connFTPC.Open
        End If
        rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
        If rec.EOF = True Then
            lTESettingAssignMAC = False
            rec.Close
            connFTPC.Close

        Else
            lTESettingAssignMAC = True
            
            If Len(Trim(rec.Fields("ORDER_NUMBER"))) <= 7 Then
                txtOrder.Text = Trim(rec.Fields("ORDER_NUMBER"))
            Else
                txtOrder.Text = Mid(Trim(rec.Fields("ORDER_NUMBER")), 1, 7)
            End If
            txtQty.Text = rec.Fields("WO_QTY")
            txtWeishu.Text = rec.Fields("MAC_QTY")
            txtMacStart.Text = rec.Fields("MAC_START")
            txtMACEnd.Text = rec.Fields("MAC_END")
'            txtQty2.Text = rec.Fields("MAC_COPIES")
'            txtRemark.Text = rec.Fields("Remark")
            If rec.Fields("PrintAllMAC") = "1" Then
                txtWeishu.Text = "1"
                txtQty.Text = CInt(rec.Fields("WO_QTY")) * CInt(rec.Fields("MAC_QTY"))
            End If
            
            rec.MoveNext
            If rec.EOF = True Then
                rec.Close
                connFTPC.Close
            Else
                txtMacStart2.Text = rec.Fields("MAC_START")
                txtMACEnd2.Text = rec.Fields("MAC_END")
                rec.MoveNext
                If rec.EOF = True Then
                    rec.Close
                    connFTPC.Close
                Else
                    txtMacStart2.Text = rec.Fields("MAC_START")
                    txtMACEnd2.Text = rec.Fields("MAC_END")
                End If
            End If
            
        End If
        
        If rec.State <> 0 Then rec.Close
        If connFTPC.State <> 0 Then connFTPC.Close
        
        If Trim(txtPart.Text) <> Trim(txtCPN.Text) Then
            MsgBox "流水号对应的产品编码和打印单号对应的产品机种不一致，请联系ME确认"
            txtID.SetFocus
            Exit Sub
        End If
        txtWO.SetFocus

Else
    txtOrder.Text = ""
    txtWeishu.Text = ""
    txtMacStart.Text = ""
    txtMACEnd.Text = ""
    txtMacStart2.Text = ""
    txtMACEnd2.Text = ""
    txtMacStart3.Text = ""
    txtMACEnd3.Text = ""
End If
End Sub
Private Sub txtID_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        If rec.State = 1 Then
            rec.Close
        End If
        sql = "select a.*,(select PrintAllMAC from MAC_PART with(NOLOCK) where PART=a.PART) PrintAllMAC from mac_record a with(NOLOCK) where a.ID = '" & Trim(Me.txtID.Text) & "'  "
        If connFTPC.State = 0 Then
           connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
           connFTPC.Open
        End If
        rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
        If rec.EOF = True Then
            lTESettingAssignMAC = False
            rec.Close
            connFTPC.Close

        Else
            lTESettingAssignMAC = True
            
            If Len(Trim(rec.Fields("ORDER_NUMBER"))) <= 7 Then
                txtOrder.Text = Trim(rec.Fields("ORDER_NUMBER"))
            Else
                txtOrder.Text = Mid(Trim(rec.Fields("ORDER_NUMBER")), 1, 7)
            End If
            
'            If Trim(txtOrder.Text) <> Trim(txtWorkOrder.Text) Then
'               MsgBox "工单号和分配工单号不一致，请联系ME确认"
'               txtID.SetFocus
'               Exit Sub
'            End If
               
            
            txtQty.Text = rec.Fields("WO_QTY")
            txtWeishu.Text = rec.Fields("MAC_QTY")
            txtMacStart.Text = rec.Fields("MAC_START")
            txtMACEnd.Text = rec.Fields("MAC_END")
'            txtQty2.Text = rec.Fields("MAC_COPIES")
'            txtRemark.Text = rec.Fields("Remark")
            If rec.Fields("PrintAllMAC") = "1" Then
                txtWeishu.Text = "1"
                txtQty.Text = CInt(rec.Fields("WO_QTY")) * CInt(rec.Fields("MAC_QTY"))
            End If
            
            rec.MoveNext
            If rec.EOF = True Then
                rec.Close
                connFTPC.Close
            Else
                txtMacStart2.Text = rec.Fields("MAC_START")
                txtMACEnd2.Text = rec.Fields("MAC_END")
                rec.MoveNext
                If rec.EOF = True Then
                    rec.Close
                    connFTPC.Close
                Else
                    txtMacStart2.Text = rec.Fields("MAC_START")
                    txtMACEnd2.Text = rec.Fields("MAC_END")
                End If
            End If
            
        End If
        
        If rec.State <> 0 Then rec.Close
        If connFTPC.State <> 0 Then connFTPC.Close
        
Else
    txtOrder.Text = ""
    txtWeishu.Text = ""
    txtMacStart.Text = ""
    txtMACEnd.Text = ""
    txtMacStart2.Text = ""
    txtMACEnd2.Text = ""
    txtMacStart3.Text = ""
    txtMACEnd3.Text = ""
End If
End Sub

'Private Sub txtSN_Change()
'
'End Sub

Private Sub txtWorkOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPart.Text = ""
        Me.txtRevision.Text = ""
        Me.chkY.Value = 0
        Me.chkY2.Value = 0
        Me.chkN.Value = 0
        Me.chkN4.Value = 0
        If Trim(Me.txtWorkOrder.Text) <> "" Then
            If rec.State = 1 Then
                rec.Close
            End If
           
            If connFTPC.State = 0 Then
                connFTPC.Open
            End If
            Dim tempWO As String
         
            tempWO = Trim(Me.txtWorkOrder.Text)
            sql = "select b.part_number,b.part_revision from WORK_ORDER a,WORK_ORDER_ITEMS b where a.order_key = b.order_key and a.order_number = '" & tempWO & "'"
            rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                MsgBox ("该工单不存在，请确认输入工单是否正确!")
                rec.Close
                Exit Sub
            Else
'
                Me.txtPart.Text = rec.Fields("part_number")
                Me.txtRevision.Text = rec.Fields("part_revision")
                
                
                If Connect.getPartList(Trim(Me.txtWorkOrder.Text)) <> "" Then
                    lPB = Connect.GetPBState(Connect.getPartList(Trim(Me.txtWorkOrder.Text)))
                    If (lPB = "NPb") Then
                        Me.chkY2.Value = 1
                        Me.chkY.Value = 0
                        Me.chkN.Value = 0
                        Me.chkN4.Value = 0
                    ElseIf (lPB = "N*") Then
                        Me.chkY2.Value = 0
                        Me.chkY.Value = 0
                        Me.chkN.Value = 1
                        Me.chkN4.Value = 0
                    ElseIf (lPB = "N4") Then
                        Me.chkY2.Value = 0
                        Me.chkY.Value = 0
                        Me.chkN.Value = 0
                        Me.chkN4.Value = 1
                    Else
                        Me.chkY.Value = 0
                        Me.chkY2.Value = 0
                        Me.chkN.Value = 0
                        Me.chkN4.Value = 0
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub
Private Sub OpenLppx()
    
        If newLableFlag = True Then
            Me.MousePointer = vbHourglass
            Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\" & "HPE 二维码新SN标签5020.lab")
        Else
            Me.MousePointer = vbHourglass
            
            Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\HP本体标签正向\" & "HP SN标签5020.lab")
        End If
    
'    Me.MousePointer = vbHourglass
'    Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\HP本体标签正向\" & "HP SN标签5020.lab")
'    If txtDesc2.Text = "" Then
'        Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP序列号标签小于30位描述.lab")
'    Else
'        Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP序列号标签大于30位描述.lab")
'    End If
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub

Sub del_sql()
    Dim delsql As String
    delsql = "delete from tblHP_Print"
    conn1.Execute delsql
End Sub

