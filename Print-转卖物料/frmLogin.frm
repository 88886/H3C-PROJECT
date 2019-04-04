VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户登录(User Login)"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5130
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fmLongin 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtNewPass 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton cmdOK 
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         Picture         =   "frmLogin.frx":073E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton CmdCancel 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         Picture         =   "frmLogin.frx":1020
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lbNew 
         Caption         =   "新密码(Password):"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblUser 
         Caption         =   "用户名(User Name):"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblPassward 
         Caption         =   "密    码(Password):"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conn As New ADODB.Connection
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

Private Sub cmdChange_Click()
    If Me.lbNew.Visible = False Then
        lbNew.Visible = True
    End If
    If Me.txtNewPass.Visible = False Then
        txtNewPass.Visible = True
    End If
    If Me.txtUserName.Text <> "" And Me.txtPassword.Text <> "" And Me.txtNewPass.Text <> "" Then
        If Connect.ChangePassword(Me.txtUserName.Text, Me.txtPassword.Text, Me.txtNewPass.Text) = True Then
            MsgBox "用户名密码变更成功!"
            Unload Me
            Exit Sub
        Else
            MsgBox "用户名密码变更失败!"
            Exit Sub
        End If
    End If
    
    

End Sub

Private Sub cmdOK_Click()

  Dim sUserName As String
  Dim sPWD As String
  sPWD = Trim(txtPassword.Text)
  sUserName = Trim(txtUserName.Text)
  
  
  Dim sSQL As String
  sSQL = "select * from users where username='" & sUserName & "' and Password='" & sPWD & "' "
  Dim rsResult As Recordset
  Set rsResult = conn.Execute(sSQL)
  If rsResult.EOF = True Then
    MsgBox "用户名密码不正确!"
    Exit Sub
  End If
  golUSERID = Trim(rsResult.Fields(0).Value)
  golUSERNAME = Trim(rsResult.Fields(1).Value)
  
  If rsResult.Fields(3).Value = "reprint" Then
    reprint = True
  Else
    reprint = False
  End If
'  reprint = True
  
  frm21H3CPrintMaterial.Visible = True
   
 
  'frmMain.Visible = True
  conn.Close
  Set conn = Nothing
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    golPath = Connect.getConnectionstring
    conn.ConnectionString = golPath
    conn.Open
    
    DSrecFTPC.CursorType = adOpenStatic
    DSrecFTPC.CursorLocation = adUseClient
    DSconFTPC.ConnectionString = "provider=SQLOLEDB;server=10.11.1.130; UID=sa; PWD=Flash123;database=afg_active_90"
    DSconFTPC.Open
  
    'Filename = "单板+3COMPrint(正向).exe"
    Filename = App.EXEName
    Directory = App.Path + "\"
    FullFileName = Directory + Filename + ".exe"
    'set graphics mode to persistent
    Me.AutoRedraw = True
    'retrieve the information
    DisplayVerInfo
    'show the results
    frmLogin.Caption = Filename + " ( Version: " & FileVer & " )"
    If Len(FileVer) < 1 Then
        MsgBox "程序检测出错!", vbOKOnly, "提示"
        Unload Me
        Exit Sub
    End If
    
    Call check_ver

End Sub

Private Sub check_ver()
    'strSql = "select APP_Version from Version_Control with(NOLOCK) where APP_Name=N'单板+3COMPrint(正向)' "
    strSql = "select APP_Version from Version_Control with(NOLOCK) where APP_Name=N'" & Filename & "' "
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
        Unload Me
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


