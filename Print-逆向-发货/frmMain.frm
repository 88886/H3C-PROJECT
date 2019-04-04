VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Main Form"
   ClientHeight    =   6330
   ClientLeft      =   5130
   ClientTop       =   2595
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10935
   Begin VB.CommandButton Command16 
      Caption         =   "海康 备件良品"
      Height          =   615
      Left            =   5520
      TabIndex        =   24
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command15 
      Caption         =   "海康 生产"
      Height          =   615
      Left            =   5520
      TabIndex        =   23
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton Command14 
      Caption         =   "海康 备件"
      Height          =   615
      Left            =   5520
      TabIndex        =   22
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command13 
      Caption         =   "大华 备件良品"
      Height          =   615
      Left            =   8280
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command12 
      Caption         =   "大华 生产"
      Height          =   615
      Left            =   8280
      TabIndex        =   20
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Command11 
      Caption         =   "大华 备件"
      Height          =   615
      Left            =   8280
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command10 
      Caption         =   "UNIS 备件"
      Height          =   615
      Left            =   5640
      TabIndex        =   18
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      Caption         =   "UNIS 生产"
      Height          =   615
      Left            =   5640
      TabIndex        =   17
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "UNIS备件良品"
      Height          =   615
      Left            =   5640
      TabIndex        =   16
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "打印程序升级日志"
      Height          =   735
      Left            =   8280
      TabIndex        =   15
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "新H3C备件良品"
      Height          =   615
      Left            =   2880
      TabIndex        =   14
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "新H3C 生产"
      Height          =   615
      Left            =   2880
      TabIndex        =   13
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "新H3C 备件"
      Height          =   615
      Left            =   2880
      TabIndex        =   12
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HUV备件良品"
      Height          =   615
      Left            =   8520
      TabIndex        =   11
      Top             =   9360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HUV 生产"
      Height          =   615
      Left            =   8520
      TabIndex        =   10
      Top             =   8640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HUV 备件"
      Height          =   615
      Left            =   8520
      TabIndex        =   9
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdHW_LiangPin 
      Caption         =   "HUAWEI备件良品"
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdLiangPin 
      Caption         =   "DP备件良品"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmd3COM 
      Caption         =   "DP 备件"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdH3C 
      Caption         =   "DP 生产"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdHUAWEI 
      Caption         =   "HUAWEI 备件"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdH3C_3COM 
      Caption         =   "HUAWEI 生产"
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退  出(Exit)"
      Height          =   735
      Left            =   8280
      TabIndex        =   0
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   8760
      Picture         =   "frmMain.frx":13652
      Top             =   7080
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label2 
      Caption         =   "2014.09.11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   6480
      Width           =   855
   End
   Begin VB.Image imgH3C 
      Height          =   810
      Left            =   480
      Picture         =   "frmMain.frx":16DB4
      Top             =   3360
      Width           =   1785
   End
   Begin VB.Image imgHUAWEI 
      Height          =   1185
      Left            =   480
      Picture         =   "frmMain.frx":17807
      Top             =   4560
      Width           =   1740
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "标签打印选择(Label Printed Select)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmd3COM_Click()
   frmH3CPrint.Show
End Sub

Private Sub cmdExit_Click()
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   End
End Sub

Private Sub cmdH3C_3COM_Click()
   frmHUAWEISHENCHAN.Show
End Sub

Private Sub cmdH3C_Click()
   frmH3Cshenchan.Show
   
End Sub
Private Sub cmdHUAWEI_Click()
   frmHUAWEIPrint.Show
End Sub

Private Sub img3COM_Click()
   cmd3COM_Click
End Sub

Private Sub imgH3C_3COM_Click()
   cmdH3C_3COM_Click
End Sub

Private Sub cmdHW_LiangPin_Click()
    frmHUAWEI_LiangPin.Show
End Sub

Private Sub cmdLiangPin_Click()
    frmH3C_LiangPin.Show
End Sub

Private Sub Command1_Click()
    frmHUVPrint.Show
End Sub

Private Sub Command10_Click()
    frmUNISPrint.Show
End Sub

Private Sub Command11_Click()
  frmDaHuaPrint.Show
End Sub

Private Sub Command12_Click()
  frmDaHuashenchan.Show
End Sub

Private Sub Command13_Click()
   frmDaHuaLiangPin.Show
End Sub

Private Sub Command14_Click()
   frmChunConsenPrint.Show
End Sub

Private Sub Command15_Click()
   frmChunConsenshenchan.Show
End Sub

Private Sub Command16_Click()
  frmChunConsenLiangPin.Show
End Sub

Private Sub Command2_Click()
   frmHUVshenchan.Show
End Sub

Private Sub Command3_Click()
    frmHUV_LiangPin.Show
End Sub

Private Sub Command4_Click()
    frmNewH3CPrint.Show
End Sub

Private Sub Command5_Click()
    frmNewH3Cshenchan.Show
End Sub

Private Sub Command6_Click()
    frmNewH3C_LiangPin.Show
End Sub

Private Sub Command7_Click()
    frmUpgradeLog.Show
End Sub

Private Sub Command8_Click()
    frmUNIS_LiangPin.Show
End Sub

Private Sub Command9_Click()
    frmUNISshenchan.Show
End Sub

Private Sub Form_Load()
    DSrecFTPC.CursorType = adOpenStatic
    DSrecFTPC.CursorLocation = adUseClient
    DSconFTPC.ConnectionString = "provider=SQLOLEDB;server=10.11.1.130; UID=sa; PWD=Flash123;database=afg_active_90"
    DSconFTPC.Open
  
    'Filename = "逆向发货Print.exe"
    Filename = App.EXEName
    Directory = App.Path + "\"
    FullFileName = Directory + Filename + ".exe"
    'set graphics mode to persistent
    Me.AutoRedraw = True
    'retrieve the information
    DisplayVerInfo
    'show the results
    frmMain.Caption = Filename + "( Version: " & FileVer & " )"
    If Len(FileVer) < 1 Then
        MsgBox "程序检测出错!", vbOKOnly, "提示"
        Unload Me
    End If
    
    Call check_ver

End Sub

Private Sub imgH3C_Click()
   cmdH3C_Click
End Sub

Private Sub imgHUAWEI_Click()
   cmdHUAWEI_Click
End Sub


Private Sub check_ver()
    'strSql = "select APP_Version from Version_Control with(NOLOCK) where APP_Name=N'逆向发货Print' "
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



