VERSION 5.00
Begin VB.Form frm21H3CPrintMaterial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "转卖物料二维码标签"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12585
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
   ScaleHeight     =   9735
   ScaleWidth      =   12585
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdGoon 
      Caption         =   "继续"
      Enabled         =   0   'False
      Height          =   615
      Left            =   11280
      TabIndex        =   22
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   615
      Left            =   10080
      TabIndex        =   21
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   960
      TabIndex        =   17
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   5700
      TabIndex        =   20
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   3330
      TabIndex        =   19
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      TabIndex        =   18
      Top             =   3120
      Width           =   12255
      Begin VB.ComboBox cboVendor 
         Height          =   450
         ItemData        =   "frm21H3CPrint.frx":13652
         Left            =   2040
         List            =   "frm21H3CPrint.frx":1365C
         TabIndex        =   4
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox txtUnit 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Text            =   "PCS"
         Top             =   1272
         Width           =   3975
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8160
         TabIndex        =   14
         Top             =   2130
         Width           =   3975
      End
      Begin VB.TextBox txtStar 
         Enabled         =   0   'False
         Height          =   495
         Left            =   8160
         TabIndex        =   15
         Text            =   "/"
         Top             =   2670
         Width           =   3975
      End
      Begin VB.TextBox txtChaomin 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8160
         TabIndex        =   12
         Top             =   1035
         Width           =   3975
      End
      Begin VB.TextBox txtPid 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Left            =   8160
         TabIndex        =   13
         Text            =   "/"
         Top             =   1575
         Width           =   3975
      End
      Begin VB.TextBox txtOriginCountry 
         Enabled         =   0   'False
         Height          =   495
         Left            =   8160
         TabIndex        =   16
         Text            =   "CHN"
         Top             =   4800
         Width           =   3975
      End
      Begin VB.TextBox txtRohs 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8160
         TabIndex        =   11
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         Top             =   756
         Width           =   3975
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txtVer 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2040
         TabIndex        =   10
         Top             =   4890
         Width           =   3975
      End
      Begin VB.TextBox txtManufacture 
         Height          =   495
         Left            =   2040
         TabIndex        =   8
         Top             =   3852
         Width           =   3975
      End
      Begin VB.TextBox txtPO 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   2820
         Width           =   3975
      End
      Begin VB.TextBox txtVendor 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7800
         TabIndex        =   40
         Top             =   3480
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox txtMPN 
         Height          =   495
         Left            =   2040
         TabIndex        =   9
         Top             =   4368
         Width           =   3975
      End
      Begin VB.TextBox txtLOT 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   3336
         Width           =   3975
      End
      Begin VB.TextBox txtDC 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2040
         TabIndex        =   5
         Top             =   2304
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "UNIT:"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   1280
         Width           =   1335
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PO:"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   2840
         Width           =   720
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MPN:"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   4400
         Width           =   1455
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CPN:"
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VENDOR:"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Manufacture:"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   3880
         Width           =   1695
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LOT:"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VER:"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label lblWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "D/C:"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   2320
         Width           =   1440
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "QTY:"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   760
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PID:"
         Height          =   375
         Left            =   6480
         TabIndex        =   28
         Top             =   1695
         Width           =   720
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "防静电等级:"
         Height          =   375
         Left            =   6480
         TabIndex        =   27
         Top             =   2790
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RoHS:"
         Height          =   375
         Left            =   6480
         TabIndex        =   26
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "潮敏:"
         Height          =   375
         Left            =   6480
         TabIndex        =   25
         Top             =   1155
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Origin Country:"
         Height          =   375
         Left            =   6240
         TabIndex        =   24
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
         Height          =   375
         Left            =   6480
         TabIndex        =   23
         Top             =   2250
         Width           =   1335
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9540
      Left            =   0
      Picture         =   "frm21H3CPrint.frx":13672
      ScaleHeight     =   9510
      ScaleWidth      =   12315
      TabIndex        =   0
      Top             =   0
      Width           =   12345
      Begin VB.CommandButton cmdClear 
         Caption         =   "清空(Clear)"
         Height          =   615
         Left            =   8040
         TabIndex        =   39
         Top             =   8760
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm21H3CPrintMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' ver 2.3.0 to be actived

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

Private Sub cmdCancel_Click()
   txtCPN.Text = ""
   txtQty.Text = ""
'   cboVendor.Index = -1
   cboVendor.Text = ""
   

   txtDC.Text = ""
   txtPO.Text = ""
   txtLOT.Text = ""
   txtManufacture.Text = ""
   txtMPN.Text = ""
   txtVer.Text = ""
   txtRohs.Text = ""
   txtChaomin.Text = ""
'   txtPid.Text = ""
   txtRemark.Text = ""
'   txtStar.Text = ""
   txtVer.Text = ""
'   txtOriginCountry.Text = ""
'   txtUnit.Text = ""
   txtCPN.SetFocus
   
End Sub

Private Sub cmdClear_Click()
    cmdCancel_Click
End Sub

Private Sub cmdGoon_Click()
    bRun = True
    cmdPrint.Enabled = False
    CmdCancel.Enabled = True
    cmdReturn.Enabled = True
    cmdStop.Enabled = True
    cmdGoon.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    'sql = "select active from tblECO_Ver where PartNumber='" & Trim(txtCPN.Text) & "' and Version='" & Trim(txtVer.Text) & "'"
    'If rec.State = 1 Then
    '  rec.Close
    'End If
   
   'rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   'If rec.EOF = False Then
   '     If rec.Fields(0).Value = "False" Then
   '         MsgBox "此版本已经被禁用,不能打印!", vbInformation + vbOKOnly, "版本已经被禁用"
   '         txtSN.SetFocus
   '         Exit Sub
   '     End If
   'End If
   'rec.Close


    If Trim(txtCPN.Text) = "" Then
        MsgBox "CPN未输入,不能打印!", vbInformation + vbOKOnly, "未输入CPN"
        txtCPN.SetFocus
        Exit Sub
    End If
   
    If Len(Trim(txtCPN.Text)) <> 8 Then
        MsgBox "CPN长度不是8位,不能打印!", vbInformation + vbOKOnly, "CPN长度不是8位"
        txtCPN.SetFocus
        Exit Sub
    End If
   
   If txtQty.Text = "" Then
      MsgBox "数量未输入，不能打印！", vbInformation + vbOKOnly, "未输入数量"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If txtUnit.Text = "" Then
      MsgBox "UNIT未输入，不能打印！", vbInformation + vbOKOnly, "未输入UNIT"
      txtUnit.SetFocus
      Exit Sub
   End If
   
   If CLng(txtQty.Text) <= 0 Then
      MsgBox "请输入正确的数量！", vbInformation + vbOKOnly, "数量不对"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If Trim(cboVendor.Text) = "" Then
      MsgBox "Vendor未输入,不能打印!", vbInformation + vbOKOnly, "未输入Vendor"
      cboVendor.SetFocus
      Exit Sub
   End If
   
   If Trim(txtDC.Text) = "" Then
      MsgBox "DC未输入，不能打印！", vbInformation + vbOKOnly, "未输入DC"
      txtDC.SetFocus
      Exit Sub
   End If
   
   If Len(Trim(txtDC.Text)) <> 6 Then
      MsgBox "DC长度不是6位,不能打印！", vbInformation + vbOKOnly, "DC长度不是6位"
      txtDC.SetFocus
      Exit Sub
   End If
   
   If Left(Trim(txtDC.Text), 3) <> "201" Then
      MsgBox "DC不是201开头,不能打印！", vbInformation + vbOKOnly, "DC不是201开头"
      txtDC.SetFocus
      Exit Sub
   End If
   
   If Trim(txtPO.Text) = "" Then
      MsgBox "PO未输入,不能打印!", vbInformation + vbOKOnly, "未输入PO"
      txtPO.SetFocus
      Exit Sub
   End If
   
    If Trim(txtLOT.Text) = "" Then
      MsgBox "LOT未输入,不能打印!", vbInformation + vbOKOnly, "未输入LOT"
      txtLOT.SetFocus
      Exit Sub
   End If
   
   If Trim(txtManufacture.Text) = "" Then
      MsgBox "Manufacture未输入,不能打印!", vbInformation + vbOKOnly, "未输入Manufacture"
      txtManufacture.SetFocus
      Exit Sub
   End If
   
   If Trim(txtMPN.Text) = "" Then
      MsgBox "MPN未输入,不能打印!", vbInformation + vbOKOnly, "未输入MPN"
      txtMPN.SetFocus
      Exit Sub
   End If
   
   If Trim(txtVer.Text) = "" Then
      MsgBox "VER未输入,不能打印!", vbInformation + vbOKOnly, "未输入VER"
      txtVer.SetFocus
      Exit Sub
   End If
   
   If Trim(txtRohs.Text) = "" Then
      MsgBox "Rohs未输入,不能打印!", vbInformation + vbOKOnly, "未输入Rohs"
      txtRohs.SetFocus
      Exit Sub
   End If

   If Trim(txtChaomin.Text) = "" Then
      MsgBox "潮敏未输入,不能打印!", vbInformation + vbOKOnly, "未输入潮敏"
      txtChaomin.SetFocus
      Exit Sub
   End If
   
   If Trim(txtPid.Text) = "" Then
      MsgBox "Pid未输入,不能打印!", vbInformation + vbOKOnly, "未输入Pid"
      txtPid.SetFocus
      Exit Sub
   End If
   
   If Trim(txtRemark.Text) = "" Then
      MsgBox "备注未输入,不能打印!", vbInformation + vbOKOnly, "未输入备注"
      txtRemark.SetFocus
      Exit Sub
   End If
   
   If Trim(txtStar.Text) = "" Then
      MsgBox "防静电等级未输入,不能打印!", vbInformation + vbOKOnly, "未输入防静电等级"
      txtStar.SetFocus
      Exit Sub
   End If
   

   If Trim(txtOriginCountry.Text) = "" Then
      MsgBox "Origin Country未输入,不能打印!", vbInformation + vbOKOnly, "未输入Origin Country"
      txtOriginCountry.SetFocus
      Exit Sub
   End If
   
   cmdPrint.Caption = "执行中..."
   cmdPrint.Enabled = False
   cmdStop.Enabled = True
    
   Dim i, j, qty, qty1 As Integer
   Dim leftstr, rightstr, str As String, str1 As String, str2 As String, endStr As String
   
   endStr = "XXXXXXXXXXXXXXXXXXXX"
   
   OpenLppx
     
    bRun = True
    Dim k As Integer
    k = 0
    Dim strPreviousLength As Integer, strFinalLength As Integer
    Dim strFinal As String, strFinal1 As String, strFinal2 As String
       
    sql = " insert into tblMaterial_log(CREATE_USER,CPN,QTY,VENDOR,DC,PO,LOT,Manufacture,MPN,VER,ROHS,CHAOMIN,PID,MARK,JinDianDengJi,OriginCountry,UNIT) "
    sql = sql & " Values( '" & golUSERNAME & "','" & UCase(Trim(txtCPN.Text)) & "','" & CLng(Trim(txtQty.Text)) & "','" & cboVendor.Text & "','" & txtDC.Text & "','" & txtPO.Text & "','" & txtLOT.Text & "','" & txtManufacture.Text & "','" & txtMPN.Text & "','" & txtVer.Text & "','" & txtRohs.Text & "','" & txtChaomin.Text & "','" & txtPid.Text & "','" & txtRemark.Text & "','" & txtStar.Text & "','" & txtOriginCountry.Text & "','" & txtUnit.Text & "')"
    Status = Connect.excuteUpdate(sql)
    If Status <> "" Then
       MsgBox "保存打印记录失败，请重新打印!" & "原因是" & Status
       bRun = False
   
        UnloadLppx
        cmdCancel_Click
   
        cmdPrint.Caption = "打印(Print) &p"
        cmdPrint.Enabled = True
    End If
    'cmdCancel_Click
    
    
    If bRun = True Then
        If k > 0 And k Mod 100 = 0 Then
            Savetime = timeGetTime '记下开始时的时间
            While timeGetTime < Savetime + 30000 '循环等待
                DoEvents '转让控制权，以便让操作系统处理其它的事件。
            Wend
        End If
keepprint:
        myVars.Item("cpn").Value = UCase(Trim(txtCPN.Text))
        myVars.Item("qty").Value = Trim(txtQty.Text)
        myVars.Item("unit").Value = Trim(txtUnit.Text)
        myVars.Item("ven").Value = Trim(cboVendor.Text)
        myVars.Item("dc").Value = Trim(txtDC.Text)
        myVars.Item("po").Value = Trim(txtPO.Text)
        myVars.Item("lot").Value = Trim(txtLOT.Text)
        myVars.Item("Man").Value = Trim(txtManufacture.Text)
        myVars.Item("mpn").Value = Trim(txtMPN.Text)
        myVars.Item("ver").Value = Trim(txtVer.Text)
        myVars.Item("rohs").Value = Trim(txtRohs.Text)
        myVars.Item("cha").Value = Trim(txtChaomin.Text)
        myVars.Item("pid").Value = Trim(txtPid.Text)
        myVars.Item("mark").Value = Trim(txtRemark.Text)
        myVars.Item("sta").Value = Trim(txtStar.Text)
        myVars.Item("ori").Value = Trim(txtOriginCountry.Text)
    
        'myApp.Visible = True
        myDoc.PrintLabel 1
        myDoc.FormFeed
    
        
        DoEvents
        
    Else
        While (bRun = False)
            'sleep 1000
            DoEvents
        Wend
        
        GoTo keepprint
    End If
   
   UnloadLppx
    

'   cmdCancel_Click
   
   cmdPrint.Caption = "打印(Print) &p"
   cmdPrint.Enabled = True
   
End Sub

Private Sub cmdStop_Click()
    bRun = False
    cmdPrint.Enabled = False
    CmdCancel.Enabled = False
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
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   bRun = False
End Sub

Private Sub txtDC_Change()
If txtDC.Text <> "" Then
    If Asc(Right(txtDC.Text, 1)) > 57 Or Asc(Right(txtDC.Text, 1)) < 48 Then
       MsgBox "只能输入数字！", vbInformation + vbOKOnly, "输入不正确"
       SendKeys "{backspace}"
       txtDC.SetFocus
       Exit Sub
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
   Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\" & "原材料二维码追溯.lab")
   Set myFormat = myDoc.Format
   'Set myDoc = myApp.Documents.Open("G:\flash\标签模板\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub




