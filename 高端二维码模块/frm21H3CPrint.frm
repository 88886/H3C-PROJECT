VERSION 5.00
Begin VB.Form frm21H3CPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H3C高端模块二维码标签"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm21H3CPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10665
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox tbFirst 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1920
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton cmdGoon 
      Caption         =   "继续"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9120
      TabIndex        =   21
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7680
      TabIndex        =   20
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   5280
      TabIndex        =   14
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   2760
      TabIndex        =   13
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   10455
      Begin VB.CheckBox chkN4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N4"
         Height          =   375
         Left            =   4200
         TabIndex        =   26
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox chkN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N*"
         Height          =   375
         Left            =   3400
         TabIndex        =   25
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtWO 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   1680
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   9120
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtXH 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   17
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y*"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkY2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y2"
         Height          =   375
         Left            =   2600
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtQty1 
         Height          =   405
         Left            =   6840
         TabIndex        =   3
         Text            =   "1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   405
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
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "工单号:"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "打印数量:"
         Height          =   375
         Left            =   7800
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品型号:"
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "环保属性:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "起始条码:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "一式几份:"
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "版本:"
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5940
      Left            =   0
      Picture         =   "frm21H3CPrint.frx":13652
      ScaleHeight     =   5910
      ScaleWidth      =   10635
      TabIndex        =   6
      Top             =   0
      Width           =   10665
   End
End
Attribute VB_Name = "frm21H3CPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' ver 2.3.0 to be actived
Dim currentMonth As String
Dim currentYear As String
Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim Pb As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myFormat As LabelManager2.Format
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Dim bRun As Boolean
Dim HK_Part_Flag As Boolean

Dim DSconFTPC As New ADODB.Connection
Dim DScmdFTPC As New ADODB.Command
Dim DSrecFTPC As New ADODB.Recordset
Dim DSfieldsFTPC As ADODB.Fields
Public Global_check_ver As Boolean

Const VS_FFI_SIGNATURE = &HFEEF04BD
Const VS_FFI_STRUCVERSION = &H10000
Const VS_FFI_FILEFLAGSMASK = &H3F&
Const VS_FF_DEBUG = &H1
Const VS_FF_PRERELEASE = &H2
Const VS_FF_PATCHED = &H4
Const VS_FF_PRIVATEBUILD = &H8
Const VS_FF_INFOINFERRED = &H10
Const VS_FF_SPECIALBUILD = &H20
Const VOS_UNKNOWN = &H0
Const VOS_DOS = &H10000
Const VOS_OS216 = &H20000
Const VOS_OS232 = &H30000
Const VOS_NT = &H40000
Const VOS__BASE = &H0
Const VOS__WINDOWS16 = &H1
Const VOS__PM16 = &H2
Const VOS__PM32 = &H3
Const VOS__WINDOWS32 = &H4
Const VOS_DOS_WINDOWS16 = &H10001
Const VOS_DOS_WINDOWS32 = &H10004
Const VOS_OS216_PM16 = &H20002
Const VOS_OS232_PM32 = &H30003
Const VOS_NT_WINDOWS32 = &H40004
Const VFT_UNKNOWN = &H0
Const VFT_APP = &H1
Const VFT_DLL = &H2
Const VFT_DRV = &H3
Const VFT_FONT = &H4
Const VFT_VXD = &H5
Const VFT_STATIC_LIB = &H7
Const VFT2_UNKNOWN = &H0
Const VFT2_DRV_PRINTER = &H1
Const VFT2_DRV_KEYBOARD = &H2
Const VFT2_DRV_LANGUAGE = &H3
Const VFT2_DRV_DISPLAY = &H4
Const VFT2_DRV_MOUSE = &H5
Const VFT2_DRV_NETWORK = &H6
Const VFT2_DRV_SYSTEM = &H7
Const VFT2_DRV_INSTALLABLE = &H8
Const VFT2_DRV_SOUND = &H9
Const VFT2_DRV_COMM = &HA

Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Dim Filename As String, Directory As String, FullFileName As String
Dim StrucVer As String, FileVer As String, ProdVer As String
Dim FileFlags As String, FileOS As String, FileType As String, FileSubType As String

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
   txtSN.Text = ""
   txtVer.Text = ""
   txtCPN.Text = ""
   txtXH.Text = ""
   chkY.Value = 0
   chkY2.Value = 0
   chkN.Value = 0
   chkN4.Value = 0
   chkN.Enabled = False
   chkN4.Enabled = False
   chkY.Enabled = False
   chkY2.Enabled = False
   txtWO.Text = ""
   txtQty.Text = ""
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

    If chkY.Value = 0 And chkY2.Value = 0 And chkN.Value = 0 And chkN4.Value = 0 Then
        MsgBox "环保属性未输入,不能打印!", vbInformation + vbOKOnly, "未输入环保属性"
        txtSN.SetFocus
        Exit Sub
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


  If txtSN.Text = "" Then
      MsgBox "产品条码未输入,不能打印!", vbInformation + vbOKOnly, "未输入产品条码"
      txtSN.SetFocus
      Exit Sub
   End If
   
'   ' add by robin for txtSN  Poka-yoke
'  If UCase(tempWO) <> "TASK" Then
'
'        If Mid(txtSN.Text, 1, 2) <> 21 Then
'          MsgBox "产品条码格式错误，请重新输入"
'           txtSN.SetFocus
'           Exit Sub
'         End If
'
'        If Mid(txtSN.Text, 3, 8) <> txtCPN.Text Then
'          MsgBox "产品条码格式错误，请重新输入"
'           txtSN.SetFocus
'           Exit Sub
'         End If
'
'        If Mid(txtSN.Text, 11, 1) <> "B" Then
'          MsgBox "产品条码格式错误，请重新输入"
'           txtSN.SetFocus
'           Exit Sub
'        End If
'
'         If Mid(txtSN.Text, 15, 1) <> 0 Then
'          MsgBox "产品条码格式错误，请重新输入"
'           txtSN.SetFocus
'           Exit Sub
'        End If
'
'         currentYear = Mid(Year(Now), 3, 2)
'
'        If Mid(txtSN.Text, 12, 2) <> currentYear Then
'        MsgBox "产品条码格式错误，请重新输入"
'           txtSN.SetFocus
'           Exit Sub
'        End If
'
'            Dim con13 As ADODB.Connection
'            Dim rs13 As ADODB.Recordset
'            Dim com As ADODB.Command
'            Dim sql1 As String
'            Set con13 = New ADODB.Connection
'            Set rs13 = New ADODB.Recordset
'            strConn1 = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
'            con13.Open ConnectionString:=strConn1
'            Set com = New ADODB.Command
'            com.ActiveConnection = con13
'            sql1 = "select top 1 creation_time from UNIT where serial_number ='" & txtSN.Text & "'"
'            com.CommandText = sql1
'            rs13.Open Source:=com
'            If rs13.EOF = True Then
'
'               MsgBox "此产品序号不存在!"
'                rs13.Close
'                cmdCancel_Click
'                Exit Sub
'
'            Else
'           currentMonth = Trim(rs13.Fields(0))
'          End If
'
'     currentMonth = Split(currentMonth, "/")(1)
'
'   Select Case currentMonth
'      Case 10
'       currentMonth = A
'      Case 11
'       currentMonth = B
'      Case 12
'       currentMonth = C
'    End Select
'
'
'    If Mid(txtSN.Text, 14, 1) <> currentMonth Then
'        MsgBox "产品条码格式错误，请重新输入"
'           txtSN.SetFocus
'           Exit Sub
'        End If
'
'   End If
'
'   ' end add
   
   
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
   
    Dim con14 As ADODB.Connection
    Dim rs14 As ADODB.Recordset
    Dim com14 As ADODB.Command

    Set con14 = New ADODB.Connection
    Set rs14 = New ADODB.Recordset
    strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
    
    con14.Open ConnectionString:=strConn
    Set com14 = New ADODB.Command
    com14.ActiveConnection = con14
    
    sql = "select 1 from H3C_HaiKangPart_set where is_valid=1 AND Part_Number like '%" & Trim(txtCPN.Text) & "%'"
    com14.CommandText = sql
       
    rs14.Open Source:=com14
   
    If rs14.EOF = False Then
        MsgBox "海康机种,不能在此打印!", vbInformation + vbOKOnly, "海康机种"
        rs14.Close
        Exit Sub
    End If
    rs14.Close
   
   cmdPrint.Caption = "执行中..."
   cmdPrint.Enabled = False
   cmdStop.Enabled = True
    
   Dim i, j, qty, qty1 As Integer
   Dim leftstr, rightstr, str As String, str1 As String, str2 As String, endStr As String
   
   endStr = "XXXXXXXXXXXXXXXXXXXX"
   
   
   qty = CInt(txtQty.Text)
   qty1 = CInt(txtQty1.Text)
   leftstr = UCase(Left(txtSN.Text, 14))
   If chkY.Value = 1 Then
        Pb = "Y*"
   ElseIf chkY2.Value = 1 Then
        Pb = "Y2"
   ElseIf chkN.Value = 1 Then
        Pb = "N*"
   ElseIf chkN4.Value = 1 Then
        Pb = "N4"
   End If
   If Pb = "Y2" Then
'      rightstr = "9" + Right(txtSN.Text, 5)
      rightstr = "0" + Right(txtSN.Text, 5)
   ElseIf (Pb = "Y*" Or Pb = "N*" Or Pb = "N4") And taskOrderFlag = False Then
      rightstr = tbFirst.Text + Right(txtSN.Text, 5)
   ElseIf (Pb = "Y*" Or Pb = "N*" Or Pb = "N4") And taskOrderFlag = True Then
      rightstr = tbFirst.Text + Right(txtSN.Text, 5)
   End If
   
   OpenLppx
     
    bRun = True
    Dim k As Integer
    k = 0
    Dim strPreviousLength As Integer, strFinalLength As Integer
    Dim strFinal As String, strFinal1 As String, strFinal2 As String
    

   For i = 0 To qty - 1 Step 3
'      str = leftstr & Right("000000" & CStr(CInt(rightstr) + i), 6)
'==================edit by ben 2011-10-14 start========================
       strPreviousLength = Len(rightstr)
       strFinal = CStr(CLng(rightstr) + i)
       strFinalLength = Len(strFinal)
       For m = strprevisouslength To strFinalLength - 1
            strFinal = "0" + strFinal
       Next
       str = leftstr & Right("000000" & strFinal, 6)
       If i + 1 > qty - 1 Then
            strFinal1 = endStr
            strFinalLength = Len(strFinal1)
       Else
            strFinal1 = CStr(CLng(rightstr) + i + 1)
            strFinalLength = Len(strFinal1)
       End If
       
       For m = strprevisouslength To strFinalLength - 1
            strFinal1 = "0" + strFinal1
       Next
       str1 = leftstr & Right("000000" & strFinal1, 6)
       If i + 2 > qty - 1 Then
            strFinal2 = endStr
            strFinalLength = Len(strFinal2)
       Else
            strFinal2 = CStr(CLng(rightstr) + i + 2)
            strFinalLength = Len(strFinal2)
       End If
       
       For m = strprevisouslength To strFinalLength - 1
            strFinal2 = "0" + strFinal2
       Next
       str2 = leftstr & Right("000000" & strFinal2, 6)
       
    If UploadH3CInfo(Pb, Trim(UCase(str)), Trim(UCase(txtVer.Text)), "NA", "N/A", "CHINA", "H3C_2D_Offline") = False _
        Or UploadH3CInfo(Pb, Trim(UCase(str1)), Trim(UCase(txtVer.Text)), "NA", "N/A", "CHINA", "H3C_2D_Offline") = False _
        Or UploadH3CInfo(Pb, Trim(UCase(str2)), Trim(UCase(txtVer.Text)), "NA", "N/A", "CHINA", "H3C_2D_Offline") = False Then
        MsgBox "资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
       
    '======Add by mike 2015.3.24 for data upload to FTPC============
    If UploadH3C_PB(Pb, Trim(str), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", "H3C_2D_Offline.exe") = False _
        Or UploadH3C_PB(Pb, Trim(str1), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", "H3C_2D_Offline.exe") = False _
        Or UploadH3C_PB(Pb, Trim(str2), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", "H3C_2D_Offline.exe") = False Then
        MsgBox "PB资料保存失败不能打印!", vbInformation + vbOKOnly, "资料保存失败"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
       
'==================edit by ben 2011-10-14 end==========================
    For j = 0 To qty1 - 1
 
        If bRun = True Then
            If k > 0 And k Mod 100 = 0 Then
                Savetime = timeGetTime '记下开始时的时间
                While timeGetTime < Savetime + 30000 '循环等待
                    DoEvents '转让控制权，以便让操作系统处理其它的事件。
                Wend
            End If
keepprint:
'            myVars.Item("2D").Value = str
'            myVars.Item("2D2").Value = str1
'            myVars.Item("2D3").Value = str2
            myVars.Item("SN").Value = Mid(str, 1, 10)
            myVars.Item("SN1").Value = Mid(str, 11, Len(str) - 10)
            myVars.Item("SN2").Value = Mid(str1, 11, Len(str1) - 10)
            myVars.Item("SN3").Value = Mid(str2, 11, Len(str2) - 10)
            
            'myVars.Item("Item").Value = "03" & UCase(Left(txtSN.Text, 6))
            If txtVer.Text = "" Or txtVer.Text = "/" Then
                'myObjs("Sver").Top = 5
                myVars.Item("Rev").Value = "N/A"
            ElseIf txtVer.Text = "00" Then
                myVars.Item("Rev").Value = ""
            Else
                'myObjs("Sver").Top = 5
                myVars.Item("Rev").Value = UCase(txtVer.Text)
            End If
            myVars.Item("Type").Value = txtXH.Text
            
 
            If chkY.Value = 1 Then
                myVars.Item("Rohs").Value = "Y*"
            ElseIf chkY2.Value = 1 Then
                myVars.Item("Rohs").Value = "Y2"
            ElseIf chkN.Value = 1 Then
                myVars.Item("Rohs").Value = "N*"
            ElseIf chkN4.Value = 1 Then
                myVars.Item("Rohs").Value = "N4"
            End If
 
            'myApp.Visible = True
            myDoc.PrintLabel 1
            myDoc.FormFeed
   
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
   Me.chkY.Enabled = False
   Me.chkY2.Enabled = False
   Me.chkN.Enabled = False
   Me.chkN4.Enabled = False
   
    DSrecFTPC.CursorType = adOpenStatic
    DSrecFTPC.CursorLocation = adUseClient
    DSconFTPC.ConnectionString = "provider=SQLOLEDB;server=10.11.1.130; UID=sa; PWD=Flash123;database=afg_active_90"
    DSconFTPC.Open
  
    Filename = "H3C_2D_Offline.exe"
    Directory = App.Path + "\"
    FullFileName = Directory + Filename
    'set graphics mode to persistent
    Me.AutoRedraw = True
    'retrieve the information
    DisplayVerInfo
    'show the results
    Me.Caption = Me.Caption + " ( Version: " & FileVer & " )"
    If Len(FileVer) < 1 Then
        MsgBox "程序检测出错!", vbOKOnly, "提示"
        Unload Me
        Exit Sub
    End If
    
    Call check_ver
    
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
'     txtVer.SetFocus
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
        If Len(txtSN.Text) <> 20 Then
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
        
      sql = "select * from SingleUnit where SN='" & Mid(txtSN.Text, 3, 8) & "'"
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
          
'      HK_Part_Flag = False
'      sql = "select 1 from H3C_HaiKangPart_set where is_valid=1 and like '%" & txtCPN.Text & "%' "
'      rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
'      If rec.EOF = True Then
'        HK_Part_Flag = False
'      Else
'        HK_Part_Flag = True
'      End If
      rec.Close
      txtWO.SetFocus
   Else
      txtWO.Text = ""
      txtCPN.Text = ""
      txtVer.Text = ""
      txtXH.Text = ""
   End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
'   If HK_Part_Flag = False Then
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\打印中心\" & "高端模块二维码标签new.lab")
'   Else
'        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\打印中心\" & "海康H3C二维码标签new.lab")
'   End If
   Set myFormat = myDoc.Format
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub


Private Sub txtWO_KeyPress(KeyAscii As Integer)
Dim tempWO As String, first As String
    If (KeyAscii = 13) Then
        If Len(Trim(txtCPN.Text)) <> 8 Then
            MsgBox "产品编码长度必须为8位!"
            txtSN.SetFocus
            Exit Sub
        Else
            taskOrderFlag = False
            tempWO = txtWO.Text
            If tempWO = "" Or tempWO = Null Then Return
            If UCase(tempWO) = "TASK" Then
                txtVer.Enabled = True
                chkY.Enabled = True
                chkY2.Enabled = True
                chkN.Enabled = True
                chkN4.Enabled = True
                Exit Sub
            End If
            first = ""
            tbFirst.Text = first
            sql = "select part_revision,(select order_type_S from [10.11.1.130].afg_active_90.dbo.UDA_Order where object_key=A.order_key) order_type from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A,[10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B " & _
            "WHERE A.order_key = B.order_key AND A.order_number ='" & tempWO & "' and ( part_number like 'HWF" & txtCPN.Text & "%' " & _
            "or part_number like 'HUV" & txtCPN.Text & "%' ) "
            rec.Open sql, conn, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                MsgBox "SAP中此工单的编码号与此产品编码不一致!"
                txtWO.Text = ""
                txtVer.Text = ""
'                txtSN.Text = ""
'                txtCPN.Text = ""
'                txtXH.Text = ""
                txtWO.SetFocus
                rec.Close
                Exit Sub
            Else
                txtVer.Text = Trim(rec.Fields(0))
                If rec.Fields(1) = "PP05" Then
                    txtVer.Enabled = True
                    chkY.Enabled = True
                    chkY2.Enabled = True
                    chkN.Enabled = True
                    chkN4.Enabled = True
                    rec.Close
                    Exit Sub
                End If
                
                rec.Close
                sql = "select Type,barcodeType from tblCustomType where PartNumber = '" & Me.txtCPN.Text & "'"
                rec.Open sql, conn, adOpenKeyset, adLockReadOnly
                If rec.EOF = False Then
                    If IsNull(rec.Fields(0)) = True Or IsNull(rec.Fields(1)) = True Then
                        MsgBox "请在整机序列号品牌维护该编码的条码类型和品牌!"
                        rec.Close
                        Exit Sub
                    Else
                        If rec.Fields(1) = "1D" Then
                            MsgBox "系统中设定该编码对应的1D条码，该程序只打印2D条码"
                            rec.Close
                            Exit Sub
                        End If
                    End If
                Else
                    MsgBox "该编码没有在整机序列号品牌维护中设定"
                    rec.Close
                    Exit Sub
                End If
                Pb = Connect.getPbByPartList(tempWO, first)
                Me.tbFirst.Text = first
                Me.tbFirst.Text = "0"
                If Pb = "Y2" Then
                    Me.chkY2.Enabled = False
                    Me.chkY2.Value = 1
                    Me.chkY.Enabled = False
                    Me.chkY.Value = 0
                    Me.chkN.Enabled = False
                    Me.chkN.Value = 0
                    Me.chkN4.Enabled = False
                    Me.chkN4.Value = 0
                ElseIf Pb = "Y*" Then
                    Me.chkY2.Enabled = False
                    Me.chkY2.Value = 0
                    Me.chkY.Enabled = False
                    Me.chkY.Value = 1
                    Me.chkN.Enabled = False
                    Me.chkN.Value = 0
                    Me.chkN4.Enabled = False
                    Me.chkN4.Value = 0
                ElseIf Pb = "N*" Then
                    Me.chkY2.Enabled = False
                    Me.chkY2.Value = 0
                    Me.chkY.Enabled = False
                    Me.chkY.Value = 0
                    Me.chkN.Enabled = False
                    Me.chkN.Value = 1
                    Me.chkN4.Enabled = False
                    Me.chkN4.Value = 0
                ElseIf Pb = "N4" Then
                    Me.chkY2.Enabled = False
                    Me.chkY2.Value = 0
                    Me.chkY.Enabled = False
                    Me.chkY.Value = 0
                    Me.chkN.Enabled = False
                    Me.chkN.Value = 0
                    Me.chkN4.Enabled = False
                    Me.chkN4.Value = 1
                Else
                    MsgBox "铅属性不是Y2或者Y*或者N*或者N4,不能打印！请确认该工单的下阶的铅属性是否设定"
                    Me.tbFirst.Text = ""
                    rec.Close
                    Exit Sub
                End If
                
                OverridePb
            End If
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

Private Sub check_ver()
    strSql = "select APP_Version from Version_Control with(NOLOCK) where APP_Name=N'H3C_2D_Offline' "
    If DSrecFTPC.State <> 0 Then DSrecFTPC.Close
    
    Set DScmdFTPC.ActiveConnection = DSconFTPC
    DScmdFTPC.CommandText = strSql
    DScmdFTPC.Properties("Command Time Out") = 900
    Set DSrecFTPC = DScmdFTPC.Execute
    Set DSfieldsFTPC = DSrecFTPC.Fields
    If Not DSrecFTPC.EOF Then
        If FileVer <> Trim(DSfieldsFTPC.Item(0)) Then
            MsgBox "程序版本错误!服务器版本:[" + Trim(DSfieldsFTPC.Item(0)) + "],请更新!", vbOKOnly + vbCritical, "提示"
            Global_check_ver = False
            Unload Me
            Exit Sub
        End If
    Else
        MsgBox "数据库未检测到版本信息!", vbOKOnly, "提示"
        Global_check_ver = False
        Exit Sub
    End If
    Global_check_ver = True
End Sub

Private Sub DisplayVerInfo()
   Dim rc As Long, lDummy As Long, sBuffer() As Byte
   Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
   Dim lVerbufferLen As Long

   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
   If lBufferLen < 1 Then
      MsgBox "No Version Info available!"
      Exit Sub
   End If

   '**** Store info to udtVerBuffer struct ****
   ReDim sBuffer(lBufferLen)
   rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
   rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
   MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)

   '**** Determine Structure Version number - NOT USED ****
   StrucVer = Format$(udtVerBuffer.dwStrucVersionh) & "." & Format$(udtVerBuffer.dwStrucVersionl)

   '**** Determine File Version number ****
   FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)

   '**** Determine Product Version number ****
   ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)

   '**** Determine Boolean attributes of File ****
   FileFlags = ""
   If udtVerBuffer.dwFileFlags And VS_FF_DEBUG Then FileFlags = "Debug "
   If udtVerBuffer.dwFileFlags And VS_FF_PRERELEASE Then FileFlags = FileFlags & "PreRel "
   If udtVerBuffer.dwFileFlags And VS_FF_PATCHED Then FileFlags = FileFlags & "Patched "
   If udtVerBuffer.dwFileFlags And VS_FF_PRIVATEBUILD Then FileFlags = FileFlags & "Private "
   If udtVerBuffer.dwFileFlags And VS_FF_INFOINFERRE Then FileFlags = FileFlags & "Info "
   If udtVerBuffer.dwFileFlags And VS_FF_SPECIALBUILD Then FileFlags = FileFlags & "Special "
   If udtVerBuffer.dwFileFlags And VFT2_UNKNOWN Then FileFlags = FileFlags + "Unknown "

   '**** Determine OS for which file was designed ****
   Select Case udtVerBuffer.dwFileOS
      Case VOS_DOS_WINDOWS16
        FileOS = "DOS-Win16"
      Case VOS_DOS_WINDOWS32
        FileOS = "DOS-Win32"
      Case VOS_OS216_PM16
        FileOS = "OS/2-16 PM-16"
      Case VOS_OS232_PM32
        FileOS = "OS/2-16 PM-32"
      Case VOS_NT_WINDOWS32
        FileOS = "NT-Win32"
      Case other
        FileOS = "Unknown"
   End Select
   Select Case udtVerBuffer.dwFileType
      Case VFT_APP
         FileType = "App"
      Case VFT_DLL
         FileType = "DLL"
      Case VFT_DRV
         FileType = "Driver"
         Select Case udtVerBuffer.dwFileSubtype
            Case VFT2_DRV_PRINTER
               FileSubType = "Printer drv"
            Case VFT2_DRV_KEYBOARD
               FileSubType = "Keyboard drv"
            Case VFT2_DRV_LANGUAGE
               FileSubType = "Language drv"
            Case VFT2_DRV_DISPLAY
               FileSubType = "Display drv"
            Case VFT2_DRV_MOUSE
               FileSubType = "Mouse drv"
            Case VFT2_DRV_NETWORK
               FileSubType = "Network drv"
            Case VFT2_DRV_SYSTEM
               FileSubType = "System drv"
            Case VFT2_DRV_INSTALLABLE
               FileSubType = "Installable"
            Case VFT2_DRV_SOUND
               FileSubType = "Sound drv"
            Case VFT2_DRV_COMM
               FileSubType = "Comm drv"
            Case VFT2_UNKNOWN
               FileSubType = "Unknown"
         End Select
      Case VFT_FONT
         FileType = "Font"
         Select Case udtVerBuffer.dwFileSubtype
            Case VFT_FONT_RASTER
               FileSubType = "Raster Font"
            Case VFT_FONT_VECTOR
               FileSubType = "Vector Font"
            Case VFT_FONT_TRUETYPE
               FileSubType = "TrueType Font"
         End Select
      Case VFT_VXD
         FileType = "VxD"
      Case VFT_STATIC_LIB
         FileType = "Lib"
      Case Else
         FileType = "Unknown"
   End Select
End Sub

