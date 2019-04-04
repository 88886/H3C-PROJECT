VERSION 5.00
Begin VB.Form frm21H3CPrint_H3C_2D 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H3C 整机模块类标签打印"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   12015
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox tbFirst 
      Enabled         =   0   'False
      Height          =   405
      Left            =   240
      TabIndex        =   25
      Top             =   7800
      Width           =   375
   End
   Begin VB.CommandButton cmdGoon 
      Caption         =   "继续"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9720
      TabIndex        =   21
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8280
      TabIndex        =   20
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   960
      TabIndex        =   15
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   5880
      TabIndex        =   14
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   3360
      TabIndex        =   13
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   11775
      Begin VB.CheckBox chkAutoTest 
         Caption         =   "有"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   48
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox chkNonAutoTest 
         BackColor       =   &H0000C000&
         Caption         =   "无"
         Enabled         =   0   'False
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
         Left            =   3720
         TabIndex        =   47
         Top             =   4080
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.TextBox txtWeishu 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8040
         TabIndex        =   36
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtPart 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         TabIndex        =   35
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtMacStart 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   34
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   33
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtMACEnd 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         TabIndex        =   32
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtOrder 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   31
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtMACEnd2 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         TabIndex        =   30
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtMacStart2 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   29
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtMACEnd3 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         TabIndex        =   28
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox txtMacStart3 
         BackColor       =   &H80000010&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   27
         Top             =   2640
         Width           =   3015
      End
      Begin VB.CheckBox chkN4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N4"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   26
         Top             =   3600
         Width           =   855
      End
      Begin VB.CheckBox chkN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N*"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtWO 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   1680
         TabIndex        =   23
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   10320
         TabIndex        =   19
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox txtXH 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   8040
         TabIndex        =   17
         Top             =   3120
         Width           =   3135
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y*"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   3600
         Width           =   735
      End
      Begin VB.CheckBox chkY2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y2"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtQty1 
         Height          =   405
         Left            =   8040
         TabIndex        =   3
         Text            =   "1"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   8040
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtVer 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   450
         Left            =   3840
         TabIndex        =   1
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "自动测试机柜:"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "Mac位数:"
         Height          =   375
         Left            =   6480
         TabIndex        =   46
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblEPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品机种:"
         Height          =   375
         Left            =   6480
         TabIndex        =   45
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "起始Mac:"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "打印单号:"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "结束mac:"
         Height          =   375
         Left            =   6480
         TabIndex        =   42
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "工单:"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "结束mac:"
         Height          =   375
         Left            =   6480
         TabIndex        =   40
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "起始Mac:"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "结束mac:"
         Height          =   375
         Left            =   6480
         TabIndex        =   38
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "起始Mac:"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "工单号:"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "打印数量:"
         Height          =   375
         Left            =   9000
         TabIndex        =   18
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品型号:"
         Height          =   375
         Left            =   6480
         TabIndex        =   16
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "环保属性:"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   6480
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "起始条码:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "一式几份:"
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本:"
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   3120
         Width           =   720
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      Picture         =   "frm21H3CPrint_H3C_2D.frx":0000
      ScaleHeight     =   2265
      ScaleWidth      =   11865
      TabIndex        =   6
      Top             =   120
      Width           =   11895
   End
End
Attribute VB_Name = "frm21H3CPrint_H3C_2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim sql, sql99 As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Dim bRun As Boolean
Dim lTESettingAssignMAC, lPrintMAC, lTESetting, lMESetting As Boolean


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


Private Sub chkCE_Click()
   If chkCE.Value = 1 Then
      chkNonCE.Value = 0
   Else
      chkNonCE.Value = 1
   End If
End Sub

Private Sub chkNonCE_Click()
   If chkNonCE.Value = 1 Then
      chkCE.Value = 0
   Else
      chkCE.Value = 1
   End If
End Sub

Private Sub chkNonChinaRoHS_Click()
   If chkNonChinaRoHS.Value = 1 Then
      chkChinaRoHS.Value = 0
   'Else
   '   chkChinaRoHS.Value = 1
   End If
End Sub

Private Sub chkNonWEEE_Click()
   If chkNonWEEE.Value = 1 Then
      chkWEEE.Value = 0
   Else
      chkWEEE.Value = 1
   End If
End Sub

Private Sub chkOS_Click()
   If chkOS.Value = 1 Then
      txtOS.Enabled = True
      txtOS.BackColor = &H80000005
   Else
      txtOS.Enabled = False
      txtOS.BackColor = &HC0C0C0
   End If
End Sub

Private Sub chkChinaRoHS_Click()
   If chkChinaRoHS.Value = 1 Then
      chkNonChinaRoHS.Value = 0
   'Else
   '   chkNonChinaRoHS.Value = 1
   End If
End Sub

Private Sub chkWEEE_Click()
   If chkWEEE.Value = 1 Then
      chkNonWEEE.Value = 0
   Else
      chkNonWEEE.Value = 1
   End If
End Sub

Private Sub cmdCancel_Click()
    Dim Ctr As Control
    For Each Ctr In Me.Controls
        If TypeOf Ctr Is TextBox Then
            Ctr.Text = ""
        End If
        If TypeOf Ctr Is CheckBox Then
            Ctr.Value = 0
        End If
    Next

   txtSN.SetFocus
End Sub

Private Sub cmdGoon_Click()
    bRun = True
    cmdPrint.Enabled = False
    cmdCancel.Enabled = True
    cmdReturn.Enabled = True
    cmdStop.Enabled = True
    cmdGoon.Enabled = False
End Sub

Private Sub cmdPrint_Click()

    
    If Me.chkY.Value = 0 And Me.chkY2.Value = 0 And Me.chkN.Value = 0 And Me.chkN4.Value = 0 Then
        MsgBox "环保属性未输入,不能打印!", vbInformation + vbOKOnly, "未输入环保属性"
        txtSN.SetFocus
        Exit Sub
    End If
    
    If Me.chkY.Value + Me.chkY2.Value + Me.chkN.Value + Me.chkN4.Value > 1 Then
        MsgBox "环保属性输入多个,不能打印!", vbInformation + vbOKOnly, "输入环保属性多个"
        txtSN.SetFocus
        Exit Sub
    End If
    
    If Me.chkAutoTest.Value = 0 And Me.chkNonAutoTest.Value = 0 Then
        MsgBox "自动测试机柜未选择,不能打印!", vbInformation + vbOKOnly, "未输入自动测试机柜"
        txtSN.SetFocus
        'Exit Sub
    End If
    
    If Me.chkAutoTest.Value + Me.chkNonAutoTest.Value > 1 Then
        MsgBox "自动测试机柜选择多个,不能打印!", vbInformation + vbOKOnly, "输入多个自动测试机柜"
        txtSN.SetFocus
        'Exit Sub
    End If
    
    sql = "select active from tblECO_Ver where PartNumber='" & Trim(txtCPN.Text) & "' and Version='" & Trim(txtVer.Text) & "'"
    If rec.State = 1 Then
      rec.Close
    End If
   
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   If rec.EOF = False Then
        If rec.Fields(0).Value = "False" Then
            MsgBox "此版本已经被禁用,不能打印!", vbInformation + vbOKOnly, "版本已经被禁用"
            txtSN.SetFocus
            Exit Sub
        End If
   End If
   rec.Close


  If txtSN.Text = "" And lMESetting = True And lTESetting = True Then
      MsgBox "产品条码未输入,不能打印!", vbInformation + vbOKOnly, "未输入产品条码"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If txtQty.Text = "" Then
      MsgBox "数量未输入，不能打印！", vbInformation + vbOKOnly, "未输入数量"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If CInt(txtQty.Text) = 0 Then
      MsgBox "请输入正确的数量！", vbInformation + vbOKOnly, "数量不对"
      txtQty.SetFocus
      Exit Sub
   End If
   
    If txtQty1.Text = "" Then
      MsgBox "一式几份未输入，不能打印！", vbInformation + vbOKOnly, "未输入数量"
      txtQty1.SetFocus
      Exit Sub
   End If
   
   If CInt(txtQty1.Text) = 0 Then
      MsgBox "请输入正确的数量！", vbInformation + vbOKOnly, "数量不对"
      txtQty1.SetFocus
      Exit Sub
   End If
   
   
   If txtVer.Text = "" Then
      MsgBox "版本未输入,不能打印!", vbInformation + vbOKOnly, "未输入版本"
      txtWO.SetFocus
      Exit Sub
   End If
   
   If txtXH.Text = "" Then
      MsgBox "型号未输入,不能打印!", vbInformation + vbOKOnly, "未输入型号"
      txtXH.SetFocus
      Exit Sub
   End If
   
   If Trim(txtPart.Text) <> Trim(txtCPN.Text) And lMESetting = True And lTESetting = True Then
        MsgBox "流水号对应的产品编码和打印单号对应的产品机种不一致，请联系ME确认"
        txtID.SetFocus
        Exit Sub
   End If
   
    If Trim(txtWO.Text) <> Trim(txtOrder.Text) And lMESetting = True And lTESetting = True Then
        MsgBox "你输入的工单号与打印单号对应的工单不一致，请联系ME确认"
        txtWO.Text = ""
        txtVer.Text = ""
        txtWO.SetFocus
        Exit Sub
    End If
    
    '''''''''''''''''Add by carson 20171205
'    lMESetting = False
'    sql = "select 1 from C_MACAndSN_Part where EFFE_FLAG='1' AND Part_Number='" & txtCPN.Text & "'"
'    rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
'    If rec.EOF = True Then
'        lMESetting = False
'    Else
'        lMESetting = True
'    End If
'    rec.Close
'
'    lTESetting = False
'    sql = "select 1 from C_PrintMAC_Part where EFFE_FLAG='1' AND Part_Number='" & txtCPN.Text & "'"
'    rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
'    If rec.EOF = True Then
'        lTESetting = False
'    Else
'        lTESetting = True
'    End If
'    rec.Close
    
    
    lPrintMAC = False
    If lMESetting = False And lTESetting = True And lTESettingAssignMAC = True Then '没有有
        MsgBox "ME未在MAC和SN合并中维护当前机种,而TE维护的打印MAC,请联系ME,TE确认"
        txtWO.Text = ""
        txtVer.Text = ""
        txtWO.SetFocus
        Exit Sub
    ElseIf lMESetting = False And lTESetting = True And lTESettingAssignMAC = False Then '没有没
        MsgBox "ME未在MAC和SN合并中维护当前机种,而TE维护的打印MAC,请联系ME,TE确认"
        txtWO.Text = ""
        txtVer.Text = ""
        txtWO.SetFocus
        Exit Sub
    ElseIf lMESetting = False And lTESetting = False And lTESettingAssignMAC = True Then '没没有
        MsgBox "ME未在MAC和SN合并中维护当前机种,而TE有分配MAC地址,请联系ME,TE确认"
        txtWO.Text = ""
        txtVer.Text = ""
        txtWO.SetFocus
        Exit Sub
    ElseIf lMESetting = False And lTESetting = Flase And lTESettingAssignMAC = False Then '没没没
        '不打印MAC
        lPrintMAC = False
    ElseIf lMESetting = True And lTESetting = True And lTESettingAssignMAC = False Then '有有没
        MsgBox "ME和TE维护的打印MAC,但TE没有分配对应的MAC地址,请联系ME,TE确认"
        txtWO.Text = ""
        txtVer.Text = ""
        txtWO.SetFocus
        Exit Sub
    ElseIf lMESetting = True And lTESetting = Flase And lTESettingAssignMAC = False Then '有没没
        MsgBox "ME维护的打印MAC,TE没有维护打印MAC,请联系ME,TE确认"
        txtWO.Text = ""
        txtVer.Text = ""
        txtWO.SetFocus
        Exit Sub
     ElseIf lMESetting = True And lTESetting = Flase And lTESettingAssignMAC = True Then '有没有
        MsgBox "ME维护的打印MAC,TE没有维护打印MAC,但TE有分配对应的MAC地址,请联系ME,TE确认"
        txtWO.Text = ""
        txtVer.Text = ""
        txtWO.SetFocus
        Exit Sub
     ElseIf lMESetting = True And lTESetting = True And lTESettingAssignMAC = True Then '有有有
        lPrintMAC = True
        '打印MAC
     End If
    '''''''''''''''''''Add by carson 20171205 end
   
   cmdPrint.Caption = "执行中..."
   cmdPrint.Enabled = False
   cmdStop.Enabled = True
    
   Dim i, j, qty, qty1 As Integer
   Dim leftstr, rightstr, str As String
   qty = CInt(txtQty.Text)
   qty1 = CInt(txtQty1.Text)
   leftstr = UCase(Left(txtSN.Text, 14))
   rightstr = tbFirst.Text + Right(txtSN.Text, 5)

   '' To double check the value of pb
    Dim Pb As String
    If chkY2.Value = 1 Then
         Pb = "Y2"
    ElseIf chkY.Value = 1 Then
         Pb = "Y*"
    ElseIf chkN.Value = 1 Then
         Pb = "N*"
    ElseIf chkN4.Value = 1 Then
         Pb = "N4"
    End If
'    If (MsgBox("环保属性为<" & pb & ">,是否继续打印", vbYesNo, "确认信息") <> vbYes) Then
'        cmdPrint.Caption = "打印(Print) &p"
'        cmdPrint.Enabled = True
'        Exit Sub
'    End If


   If txtPart.Text = "" And lPrintMAC = True Then
      MsgBox "产品编码未带出,不能打印,请重新输入产品条码!", vbInformation + vbOKOnly, "未带出编码"
      txtSN.SetFocus
      Exit Sub
   End If
   If txtQty.Text = "" Then
      MsgBox "未输入数量！", vbInformation + vbOKOnly, "未输入数量"
      txtQty.SetFocus
      Exit Sub
   End If
'   If txtQty2.Text = "" Then
'      MsgBox "未输入一式几份数量！", vbInformation + vbOKOnly, "未输入一式几份数量"
'      txtQty2.SetFocus
'      Exit Sub
'   End If
    If txtWeishu.Text = "" And lPrintMAC = True Then
      MsgBox "未输入位数！", vbInformation + vbOKOnly, "未输入位数"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If CInt(txtQty.Text) = 0 Then
      MsgBox "请输入正确的数量！", vbInformation + vbOKOnly, "数量不对"
      txtQty.SetFocus
      Exit Sub
   End If
   If CInt(txtQty1.Text) = 0 Then
      MsgBox "请输入正确的一式几份数量！", vbInformation + vbOKOnly, "一式几份数量不对"
      txtQty2.SetFocus
      Exit Sub
   End If
   
   If Trim(txtWeishu.Text) = "" Then
        txtWeishu.Text = "0"
   End If
   If CInt(txtWeishu.Text) = 0 And lMESetting = True And lTESetting = True Then
      MsgBox "请输入正确的位数！", vbInformation + vbOKOnly, "位数不对"
      txtWeishu.SetFocus
      Exit Sub
   End If


   Dim sn, mac, ip, model, part, reprintmac As String
   Dim weishu, qty2 As Integer
   part = UCase(txtPart.Text)
   mac = UCase(txtMacStart.Text)
   weishu = CInt(txtWeishu.Text)
   qty = CInt(txtQty.Text)
'   qty2 = CInt(txtQty2.Text)
'   reprintmac = UCase(txtReprintMac.Text)

   '开始计算
    Dim arr() As Double
    Dim number2 As Integer
    ReDim Preserve arr(qty) As Double
    Dim MACindex As Integer
    
    MACindex = 1
    If lPrintMAC = True Then
        arr(0) = HEXTODEC(mac)
    End If
'    For i = 1 To qty - 1
'        arr(i) = arr(i - 1) + weishu
'        If dectohex(arr(i)) > UCase(txtMACEnd.Text) And MACindex = 1 Then
'            arr(i) = HEXTODEC(UCase(txtMacStart2.Text))
'            MACindex = 2
'        End If
'        If dectohex(arr(i)) > UCase(txtMACEnd2.Text) And MACindex = 2 Then
'            arr(i) = HEXTODEC(UCase(txtMacStart3.Text))
'            MACindex = 3
'        End If
'        'MsgBox arr(i)
'    Next


   OpenLppx
     
    bRun = True
    Dim k As Integer
    k = 0
    Dim strPreviousLength As Integer, strFinalLength As Integer
    Dim strFinal, strConstant As String
'    If (Me.chkY2.Value = True) Then
'        strConstant = "900000"
'    Else
'        strConstant = "000000"
'    End If


''''''''''qty 为打印数量
   For i = 0 To qty - 1
'      str = leftstr & Right("000000" & CStr(CInt(rightstr) + i), 6)
'==================edit by ben 2011-10-14 start========================
       strPreviousLength = Len(rightstr)
       strFinal = CStr(CLng(rightstr) + i)
       strFinalLength = Len(strFinal)
       For m = strprevisouslength To strFinalLength - 1
            strFinal = "0" + strFinal
       Next
       str = leftstr & "0" & Right("000000" & strFinal, 5)

       
    '======Add by mike 2015.3.24 for data upload to FTPC============
    If UploadH3C_PB(Pb, Trim(str), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False Then
        MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
    
    '======Add By Robin 2018.8.9 for h3c_pb 版本抓取问题修改 start=========
    
      If UploadH3C_PB_Version(Trim(str), Trim(Me.txtVer.Text)) = False Then
        MsgBox "PB版本保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
     '======Add By Robin 2018.8.9 for h3c_pb 版本抓取问题修改 END=========
     
    
      '===============add by Carson 2015-12-15 start===============
    If reprint = False Then
       If Connect.isPrintedLabel(Trim(str), Me.Name) = True Then
           MsgBox ("此序列号已打印！")
           txtSN.SetFocus
           UnloadLppx
           cmdCancel_Click
           cmdPrint.Caption = "打印(Print) &p"
           cmdPrint.Enabled = True
           Exit Sub
       End If
               
    End If
    '===============add by Carson 2015-12-15 end=================
    
    If lPrintMAC = True Then
        If i > 0 Then
            arr(i) = arr(i - 1) + weishu
            If dectohex(arr(i)) > UCase(txtMACEnd.Text) And MACindex = 1 Then
                arr(i) = HEXTODEC(UCase(txtMacStart2.Text))
                MACindex = 2
            End If
            If dectohex(arr(i)) > UCase(txtMACEnd2.Text) And MACindex = 2 Then
                arr(i) = HEXTODEC(UCase(txtMacStart3.Text))
                MACindex = 3
            End If
        End If
    End If


'==================edit by ben 2011-10-14 end==========================
''''''''''qty1 为一式几份
    For j = 0 To qty1 - 1
 
        If bRun = True Then
            If k > 0 And k Mod 100 = 0 Then
                Savetime = timeGetTime '记下开始时的时间
                While timeGetTime < Savetime + 30000 '循环等待
                    DoEvents '转让控制权，以便让操作系统处理其它的事件。
                Wend
            End If
keepprint:
            myVars.Item("sn").Value = str
            'myVars.Item("Item").Value = "03" & UCase(Left(txtSN.Text, 6))
            If txtVer.Text = "" Or txtVer.Text = "/" Then
                'myObjs("Sver").Top = 5
                myVars.Item("rev").Value = "N/A"
            ElseIf Me.txtVer.Text <> "" Then
                'modified by noel.zhou
                myVars.Item("rev").Value = Trim(txtVer.Text)
            Else
                'myObjs("Sver").Top = 5
                myVars.Item("rev").Value = UCase(txtVer.Text)
            End If
            myVars.Item("PID").Value = txtXH.Text
            If Me.chkY.Value = 1 Then
                myVars.Item("Rohs").Value = "Y*"
            ElseIf Me.chkY2.Value = 1 Then
                myVars.Item("Rohs").Value = "Y2"
            ElseIf Me.chkN.Value = 1 Then
                myVars.Item("Rohs").Value = "N*"
            ElseIf Me.chkN4.Value = 1 Then
                myVars.Item("Rohs").Value = "N4"
            End If
 
'            If chkChinaRoHS.Value = 1 Then
'                If Left(txtSN.Text, 2) = "03" Or Left(txtSN.Text, 2) = "21" Then
'                    myVars.Item("Rohs").Value = "Y*"
'                End If
'                If Left(txtSN.TabIndex, 2) = "02" Then
'                    myVars.Item("Rohs").Value = "Y2"
'                End If
'            Else
'                If Left(txtSN.Text, 2) = "03" Or Left(txtSN.Text, 2) = "21" Then
'                    myVars.Item("Rohs").Value = "N*"
'                End If
'                If Left(txtSN.Text, 2) = "02" Then
'                    myVars.Item("Rohs").Value = "Y1"
'                End If
'            End If
 
            'myApp.Visible = True
            If lPrintMAC = False Then
                mac = ""
                myObjs("text3").Top = 10000
                myObjs("MAC").Top = 10000
                myObjs("MAC1").Top = 10000
                
                If chkNonAutoTest.Value = 1 Then
                    myObjs("SN&MAC").Top = 10000
                    myObjs("SN2").Top = 10000
                    myObjs("MAC(2)").Top = 10000
                End If
            Else
                mac = dectohex(arr(i))
                myVars.Item("MAC").Value = mac
                If chkNonAutoTest.Value = 1 Then
                    myObjs("SN&MAC").Top = 10000
                    myObjs("SN2").Top = 10000
                    myObjs("MAC(2)").Top = 10000
                End If
            End If
            
'            If chkNonAutoTest.Value = 1 Then
'                myObjs("SN&MAC").Top = 10000
'                myObjs("SN2").Top = 10000
'                myObjs("MAC(2)").Top = 10000
'            End If
    '==============================Add by carson 20171207
If lPrintMAC = True Then
    sql = "select serial_number,MAC from C_MACAndSN_PrintRecord where EFFE_FLAG='1' AND (serial_number='" & str & "' or mac='" & mac & "')"
    If connFTPC.State = 0 Then
        connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
        connFTPC.Open
    End If
    rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
    If rec.EOF = False Then
        lMAC = rec.Fields("MAC")
        lSN = rec.Fields("serial_number")
        rec.Close
        connFTPC.Close
        If lMAC <> mac Then
            MsgBox ("序列号：" & str & " 打印过,其MAC地址为：" & lMAC & ",新分配的MAC地址为：" & mac & ",两者不一样,请联系ME确认！")
            txtSN.SetFocus
            UnloadLppx
            cmdCancel_Click
            cmdPrint.Caption = "打印(Print) &p"
            cmdPrint.Enabled = True
            Exit Sub
        ElseIf lSN <> str Then
            MsgBox ("MAC：" & mac & " 打印过,其序列号为：" & lSN & ",新分配的序列号为：" & str & ",两者不一样,请联系ME确认！")
            txtSN.SetFocus
            UnloadLppx
            cmdCancel_Click
            cmdPrint.Caption = "打印(Print) &p"
            cmdPrint.Enabled = True
            Exit Sub
        End If
    
End If
End If


    If chkAutoTest.Value = 1 Then
        lAutoTest = "Y"
    ElseIf chkNonAutoTest.Value = 1 Then
        lAutoTest = "N"
    Else
        lAutoTest = ""
    End If
    
    
    
'    If addPrintedMACAndSN(ByVal Operator As String, ByVal serial_number As String, ByVal mac As String, ByVal MAC_RECORD_ID As String, ByVal MAC_START As String, ByVal MAC_END As String) Then
'        MsgBox ("此序列号已打印！")
'        txtSN.SetFocus
'        UnloadLppx
'        cmdCancel_Click
'        cmdPrint.Caption = "打印(Print) &p"
'        cmdPrint.Enabled = True
'        Exit Sub
'    End If
    '==============================Add by carson 20171207 end
            
            myDoc.PrintLabel 1
            myDoc.FormFeed
            'Call Connect.addPrintedLabel(Trim(str), Me.Name)
            Call Connect.addPrintedLabelMAC(Trim(str), mac, Me.Name)
            Call Connect.addPrintedMACAndSN(golUSERNAME, Trim(str), mac, Trim(txtID.Text), txtMacStart.Text, txtMACEnd.Text, txtXH.Text, lAutoTest)
            k = k + 1
            
            DoEvents
            
        Else
            While (bRun = False)
                'sleep 1000
                DoEvents
            Wend
            
            GoTo keepprint
        End If
   Next
   Next
   
   UnloadLppx
    

   cmdCancel_Click
   
   cmdPrint.Caption = "打印(Print) &p"
   cmdPrint.Enabled = True
   
End Sub

Private Sub cmdStop_Click()
    bRun = False
    cmdPrint.Enabled = False
    cmdCancel.Enabled = False
    cmdReturn.Enabled = False
    cmdStop.Enabled = False
    cmdGoon.Enabled = True
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   If connFTPC.State = 0 Then
      connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
      connFTPC.Open
   End If
   
   lTESettingAssignMAC = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   bRun = False
End Sub

Private Sub txtHV_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 13) Then
     txtMS.SetFocus
  End If
End Sub

Private Sub txtMS_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtNAL.SetFocus
  End If
End Sub



Private Sub txtNAL_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtRemark.SetFocus
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
            txtPart.Text = rec.Fields("PART")
            
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
    txtPart.Text = ""
    txtMacStart.Text = ""
    txtMACEnd.Text = ""
    txtMacStart2.Text = ""
    txtMACEnd2.Text = ""
    txtMacStart3.Text = ""
    txtMACEnd3.Text = ""
    chkNonAutoTest.Value = 0
    chkAutoTest.Value = 0
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
     'txtVer.SetFocus
     cmdPrint_Click
  End If
End Sub



Private Sub txtRemark_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     cmdPrint_Click
  End If
End Sub

Private Sub txtQty1_Change()
If txtQty1.Text <> "" Then
    If Asc(Right(txtQty1.Text, 1)) > 57 Or Asc(Right(txtQty1.Text, 1)) < 48 Then
       MsgBox "只能输入数字！", vbInformation + vbOKOnly, "输入不正确"
       SendKeys "{backspace}"
       txtQty1.SetFocus
       Exit Sub
    End If
End If
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      txtVer.Enabled = False
      If Len(Trim(txtSN.Text)) <> 20 Then
         MsgBox "产品序号长度必须为20位!"
         txtSN.SetFocus
         Exit Sub
      End If


        Dim rcd As New ADODB.Recordset
        sql = "select * from tblCustomType where PartNumber='" & Mid(txtSN.Text, 3, 8) & "'"
        rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rcd.EOF = True Then
           MsgBox "品牌未维护!"
           txtSN.Text = ""
           txtSN.SetFocus
           rcd.Close
           Exit Sub
        Else
            If rcd.Fields(1) = "Non-H3C" Then
                MsgBox "请使用[非H3C整机模块类标签程序]打印!"
                txtSN.Text = ""
                txtSN.SetFocus
                rcd.Close
                Exit Sub
            End If
        End If
        rcd.Close
        

      sql = "select ID,SN,TYPE,CASE PB WHEN 1 THEN 'Yes' when 0 then 'No' else 'Non' end from SingleUnit where SN='" & Mid(txtSN.Text, 3, 8) & "'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "此产品编码未进行设置!"
         txtVer.Text = ""
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
        txtCPN.Text = Trim(rec.Fields(1))
        txtXH.Text = Trim(rec.Fields(2))
      End If
      rec.Close
      txtID.SetFocus
   Else
   
    chkN.Value = 0
    chkY.Value = 0
    chkY2.Value = 0
    chkN4.Value = 0
    txtWO.Text = ""
    txtCPN.Text = ""
    txtVer.Text = ""
    txtXH.Text = ""
    txtID.Text = ""
    txtOrder.Text = ""
    txtWeishu.Text = ""
    txtPart.Text = ""
    txtMacStart.Text = ""
    txtMACEnd.Text = ""
    txtMacStart2.Text = ""
    txtMACEnd2.Text = ""
    txtMacStart3.Text = ""
    txtMACEnd3.Text = ""
    chkNonAutoTest.Value = 0
    chkAutoTest.Value = 0
    
    
   End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\10.11.1.25\Public\Manufacture\打印中心\" & "整机SN标签40x15x15.lab")
   'Set myDoc = myApp.Documents.Open("\\10.11.1.25\Public\Manufacture\标签模板\" & "21H3C.lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub



Private Sub txtWO_KeyPress(KeyAscii As Integer)
    Dim tempWO As String
    If (KeyAscii = 13) Then
        If Len(Trim(txtCPN.Text)) <> 8 Then
            MsgBox "产品编码长度必须为8位!"
            txtSN.SetFocus
            Exit Sub
        Else
            tempWO = txtWO.Text
            If tempWO = "" Or tempWO = Null Then Return
            If UCase(tempWO) = "TASK" Then
                txtVer.Enabled = True
                chkY.Enabled = True
                chkY2.Enabled = True
                chkN.Enabled = True
                chkN4.Enabled = True
                chkNonAutoTest.Enabled = True
                chkAutoTest.Enabled = True
                
                Exit Sub
            End If
            
          '''''''''''''''''''''''''''''''
            lMESetting = False
            sql = "select 1 from C_MACAndSN_Part where EFFE_FLAG='1' AND Part_Number='" & txtCPN.Text & "'"
            If connFTPC.State = 0 Then
                connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
                connFTPC.Open
            End If
            rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                lMESetting = False
            Else
                lMESetting = True
            End If
            rec.Close
            
            lTESetting = False
            sql = "select 1 from C_PrintMAC_Part where EFFE_FLAG='1' AND Part_Number='" & txtCPN.Text & "'"
            rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                lTESetting = False
            Else
                lTESetting = True
            End If
            rec.Close
            connFTPC.Close
          ''''''''''''''''''''''''''''''''
            
            If Trim(txtWO.Text) <> Trim(txtOrder.Text) And lMESetting = True And lTESetting = True Then
                MsgBox "你输入的工单号与打印单号对应的工单不一致，请联系ME确认"
                txtWO.Text = ""
                txtVer.Text = ""
                txtWO.SetFocus
                Exit Sub
            End If
            
'            While (Len(tempWO) < 12)
'                tempWO = "0" & tempWO
'            Wend
'            sql = "select MaterialRevision from [10.11.1.17].dsActive.dbo.SAP_WO " & _
'                "where WorkOrderNumber = '" & tempWO & "' and ( MaterialNumber like 'HWF" & txtCPN.Text & "%' " & _
'                "or MaterialNumber like 'HUV" & txtCPN.Text & "%' ) "
            sql = "select part_revision,part_number,(select order_type_S from [10.11.1.130].afg_active_90.dbo.UDA_Order where object_key=A.order_key) order_type from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A,[10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B " & _
                "WHERE A.order_key = B.order_key AND A.order_number ='" & tempWO & "' and ( part_number like 'HWF" & txtCPN.Text & "%')"
            rec.Open sql, conn, adOpenForwardOnly, adLockReadOnly
            
            
            If rec.EOF = True Then
                MsgBox "SAP中此工单的编码号与此产品编码不一致或者该工单是HUV工单!"
                txtWO.Text = ""
                txtVer.Text = ""
                txtWO.SetFocus
                rec.Close
                Exit Sub
            Else
                txtVer.Text = Trim(rec.Fields(0))
                If rec.Fields(2) = "PP05" Then
                    txtVer.Enabled = True
                    chkY.Enabled = True
                    chkY2.Enabled = True
                    chkN.Enabled = True
                    chkN4.Enabled = True
                    rec.Close
                    Exit Sub
                End If
                If Mid(Trim(rec.Fields(1)), InStr(Trim(rec.Fields(1)), "0"), 4) = "0212" Then
                    sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order =(select top 1 leading_order from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport where order_number='" & tempWO & "') and (assembly like 'HWF0302%' or assembly like 'HUV0302%')"
                Else
                    sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & tempWO & "' and (assembly like 'HWF0302%' or assembly like 'HUV0302%')"
                End If
                rec.Close
                rec.Open sql, conn, adOpenKeyset, adLockReadOnly
                If rec.EOF = True Then
                    If Trim(txtWO.Text) <> "740026" Then
'                        MsgBox "SAP中此工单不包含0302阶单板不能打印,请确认!"
'                        txtWO.Text = ""
'                        txtVer.Text = ""
'                        txtWO.SetFocus
'                        rec.Close
'                        Exit Sub
                        rec.Close
'                        sql = "select distinct assembly from [10.11.1.130].afg_active_90.dbo.afg_b_SAPWIPReport A  where A.leading_order ='" & tempWO & "' and ( assembly like 'HWF0231%' or assembly like 'HWF0235%')"
'                        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
'                        If rec.EOF = True Then
'                            getPbByPartList = "Non"
'                            MsgBox "SAP中此工单不包含0302阶或整机阶,不能打印,请确认!"
'                            rec.Close
'                            Exit Sub
                        sql = "select distinct B.part_number from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A JOIN [10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B ON A.order_key=B.order_key where order_number='" & tempWO & "' and ( B.part_number like 'HWF0231%' or B.part_number like 'HWF0235%')"
                        rec.Open sql, conn, adOpenKeyset, adLockReadOnly
                        If rec.EOF = True Then
                            getPbByPartList = "Non"
                            MsgBox "SAP中此工单不包含0302阶或整机阶,不能打印,请确认!"
                            rec.Close
                            Exit Sub
                        Else
                            Do While Not rec.EOF
                                partlist = partlist + Mid(rec!part_number, 4, 8) + ";"
                                rec.MoveNext
                            Loop
                        End If
                    Else
                        Me.chkY2.Value = 1
                        Me.chkY2.Enabled = False
                        Me.chkY.Value = 0
                        Me.chkY.Enabled = False
                        Me.chkN.Value = 0
                        Me.chkN.Enabled = False
                        Me.chkN4.Value = 0
                        Me.chkN4.Enabled = False
                        rec.Close
                        Exit Sub
                    End If
                    
                Else
                    Do While Not rec.EOF
                        partlist = partlist + Mid(rec!assembly, 4, 8) + ";"
                        rec.MoveNext
                    Loop
                End If
                rec.Close
                Dim cmd As New ADODB.Command
                cmd.ActiveConnection = conn
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "[PbHandler]"
                cmd.Parameters.Append cmd.CreateParameter("partlist", adVarChar, adParamInput, 8000, partlist)
                cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 8)
                cmd.Parameters.Append cmd.CreateParameter("first", adVarChar, adParamOutput, 1)
                cmd.Execute
'                tbFirst.Text = cmd("first")        '' Cancel the 9 Principle
                tbFirst.Text = "0"
                Select Case cmd("res")
                    Case "No"
                        Me.chkY2.Value = 1
                        Me.chkY2.Enabled = False
                        Me.chkY.Value = 0
                        Me.chkY.Enabled = False
                        Me.chkN.Value = 0
                        Me.chkN.Enabled = False
                        Me.chkN4.Value = 0
                        Me.chkN4.Enabled = False
                    Case "Non"
                        MsgBox "此工单包含0302阶单板未设定有铅无铅,请相关ME去设定!"
                        txtWO.Text = ""
                        txtVer.Text = ""
                        txtWO.SetFocus
                        Exit Sub
                    Case "Half"
                        Me.chkY2.Value = 0
                        Me.chkY2.Enabled = False
                        Me.chkY.Value = 0
                        Me.chkY.Enabled = False
                        Me.chkN.Value = 0
                        Me.chkN.Enabled = False
                        Me.chkN4.Value = 1
                        Me.chkN4.Enabled = False
                    Case "Yes"
                        Me.chkY2.Value = 0
                        Me.chkY2.Enabled = False
                        Me.chkY.Value = 0
                        Me.chkY.Enabled = False
                        Me.chkN.Value = 0
                        Me.chkN.Enabled = False
                        Me.chkN4.Value = 1
                        Me.chkN4.Enabled = False
                End Select
                
                OverridePb   '' check label history
                
            End If
            
            sql99 = "select top 1 case when AutoTest is null then 'No' when AutoTest = 0 then 'No' when AutoTest = 1 then 'Yes' end as 'AutoTest' from tblH3CNew  where Part_Number ='" & Mid(txtSN.Text, 3, 8) & "'  and Part_Revision = '" & txtVer.Text & "'"
            rec1.Open sql99, conn, adOpenKeyset, adLockOptimistic
            If rec1.EOF = False Then
                If rec1.Fields(0) = "No" Then
                    Me.chkNonAutoTest.Value = 1
                    Me.chkAutoTest.Value = 0
                ElseIf rec1.Fields(0) = "Yes" Then
                    Me.chkNonAutoTest.Value = 0
                    Me.chkAutoTest.Value = 1
                End If
                Me.chkNonAutoTest.Enabled = False
                Me.chkAutoTest.Enabled = False
            End If
            rec1.Close
            Me.chkNonAutoTest.Value = 1
            Me.chkAutoTest.Value = 0
        End If
    Else
        txtVer.Text = ""
    End If
End Sub

Private Sub OverridePb()
    Dim labelHistory As New Label_History
    Dim sn As String
    sn = txtSN.Text
    If labelHistory.Init(sn) Then
    
        chkY.Value = 0
        chkY2.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
       
        Select Case labelHistory.Pb
        Case "Y*"
            chkY.Value = 1
        Case "Y2"
            chkY2.Value = 1
        Case "N*"
            chkN.Value = 1
        Case "N4"
            chkN4.Value = 1
        
        End Select
        
    End If
End Sub

