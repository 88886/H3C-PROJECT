VERSION 5.00
Begin VB.Form frmH3CPrintWarrantyByPart 
   BackColor       =   &H00E0E0E0&
   Caption         =   "保修卡打印"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13680
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
   ScaleHeight     =   8145
   ScaleWidth      =   13680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      Height          =   615
      Left            =   6240
      TabIndex        =   9
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      Height          =   615
      Left            =   3840
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "打印(Print) &p"
      Height          =   615
      Left            =   1440
      TabIndex        =   7
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtQty 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   6600
      TabIndex        =   6
      Text            =   "1"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtQty1 
      Height          =   405
      Left            =   3240
      TabIndex        =   4
      Text            =   "1"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtSN 
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "打印数量:"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblDes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "一式几份:"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "序列号:"
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   " 保  修  卡  打  印"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmH3CPrintWarrantyByPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' ver 2.3.0 to be actived

Dim rec As New ADODB.Recordset
Dim sql As String
Dim SN As String
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
   txtSN.Text = ""
   txtSN.SetFocus
   
End Sub

Private Sub cmdClear_Click()
    cmdCancel_Click
End Sub


Private Sub cmdPrint_Click()

  Dim part As String
  Dim res As New ADODB.Recordset
   Dim res2 As New ADODB.Recordset
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
  
  
  
   SN = UCase(Trim(txtSN.Text))
     If SN = "" Then
         MsgBox "请扫描产品序列号!"
         txtSN.Text = ""
         txtSN.SetFocus
         Exit Sub
     End If
     
     sql = "select part_number from UNIT where serial_number = '" & SN & "'"
     If res2.State = 1 Then
      res2.Close
       End If
       res2.Open sql, connFTPC, adOpenKeyset, adLockOptimistic
       part = Trim(res2.Fields(0))
     If part = "" Then
      MsgBox "请输入正确的序列号!"
         txtSN.Text = ""
         txtSN.SetFocus
         Exit Sub
     End If

     
     sql = "select url_S from AT_afg_warrantyCard_Maintain where partNumber_S = '" & part & "'and valid_I =1"
     If res.State = 1 Then
      res.Close
       End If
     
       res.Open sql, connFTPC, adOpenKeyset, adLockOptimistic
       
        If res.EOF = True Then
         MsgBox "此机种网址未进行维护!"
         txtSN.Text = ""
         txtSN.SetFocus
         res.Close
         Exit Sub
       End If
             
        
   Dim i, j, qty, qty1 As Integer
   qty = CInt(txtQty.Text)
   qty1 = CInt(txtQty1.Text)
   
     OpenLppx
     
     bRun = True
     
   Dim k As Integer
   k = 0
   
   For i = 0 To qty - 1
    For j = 0 To qty1 - 1
        
        If bRun = True Then
        
            If k > 0 And k Mod 100 = 0 Then
                Savetime = timeGetTime '记下开始时的时间
                While timeGetTime < Savetime + 30000 '循环等待
                    DoEvents '转让控制权，以便让操作系统处理其它的事件。
                Wend
            End If
      
keepprint:
       ' myVars.Item("机种").Value = partCut
       'myVars.Item("产品代码").Value = Trim(rec.Fields(0))
        myVars.Item("资料网址").Value = Trim(res.Fields(0))
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

    sql = " insert into H3C_WarrantyPrint(partNum,printUser) "
    sql = sql & " Values( '" & part & "','" & golUSERNAME & "')"
    Status = Connect.excuteUpdate(sql)
    If Status <> "" Then
       MsgBox "保存打印记录失败，请重新打印!" & "原因是" & Status
        UnloadLppx
        cmdCancel_Click
    End If

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
Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\H3C保修卡\" & "保修卡311A01M.lab")
   Set myFormat = myDoc.Format
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub





Private Sub txtSN_()
cmdPrint_Click
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
   cmdPrint_Click
   End If
End Sub

