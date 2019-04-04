Attribute VB_Name = "modExtButton"
 Option Explicit
    
  '==================================================================
  '   modExtButton.bas
  '
  '   本模块可让你改变命令按钮的文本颜色。
  '   使用方法:
  '
  '   -   在设计时将文本的Style设为Graphical.
  '
  '   -   随意设定背景色和图象属性.
  '
  '   -   在Form_Load中调用   SetButton   :
  '   SetButton   Command1.hWnd,   vbBlue
  '   (你可以任意次的调用该过程甚至不必先调用   RemoveButton.)
  '
  '   -   在Form_Unload中调用   RemoveButton   :
  '   RemoveButton   Command1.hWnd
  '
  '==================================================================
    
  Private Type RECT
  Left   As Long
  Top   As Long
  Right   As Long
  Bottom   As Long
  End Type
    
  Private Declare Function GetParent Lib "user32" _
  (ByVal hWnd As Long) As Long
    
  Private Declare Function GetWindowLong Lib "user32" Alias _
  "GetWindowLongA" (ByVal hWnd As Long, _
  ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias _
  "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
  ByVal dwNewLong As Long) As Long
  Private Const GWL_WNDPROC = (-4)
    
  Private Declare Function GetProp Lib "user32" Alias "GetPropA" _
  (ByVal hWnd As Long, ByVal lpString As String) As Long
  Private Declare Function SetProp Lib "user32" Alias "SetPropA" _
  (ByVal hWnd As Long, ByVal lpString As String, _
  ByVal hData As Long) As Long
  Private Declare Function RemoveProp Lib "user32" Alias _
  "RemovePropA" (ByVal hWnd As Long, _
  ByVal lpString As String) As Long
    
  Private Declare Function CallWindowProc Lib "user32" Alias _
  "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
  ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
  ByVal lParam As Long) As Long
    
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (Destination As Any, Source As Any, ByVal Length As Long)
    
  'Owner   draw   constants
  Private Const ODT_BUTTON = 4
  Private Const ODS_SELECTED = &H1
  'Window   messages   we're   using
  Private Const WM_DESTROY = &H2
  Private Const WM_DRAWITEM = &H2B
    
  Private Type DRAWITEMSTRUCT
  CtlType   As Long
  CtlID   As Long
  itemID   As Long
  itemAction   As Long
  itemState   As Long
  hwndItem   As Long
  hDC   As Long
  rcItem   As RECT
  itemData   As Long
  End Type
    
  Private Declare Function GetWindowText Lib "user32" Alias _
  "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, _
  ByVal cch As Long) As Long
  'Various   GDI   painting-related   functions
  Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
  (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, _
  lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, _
  ByVal crColor As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, _
  ByVal nBkMode As Long) As Long
  Private Const TRANSPARENT = 1
    
  Private Const DT_CENTER = &H1
  Public Enum TextVAligns
  DT_VCENTER = &H4
  DT_BOTTOM = &H8
  End Enum
  Private Const DT_SINGLELINE = &H20
    
    
  Private Sub DrawButton(ByVal hWnd As Long, ByVal hDC As Long, _
  rct As RECT, ByVal nState As Long)
    
  Dim s     As String
  Dim va     As TextVAligns
    
  va = GetProp(hWnd, "VBTVAlign")
    
  'Prepare   DC   for   drawing
  SetBkMode hDC, TRANSPARENT
  SetTextColor hDC, GetProp(hWnd, "VBTForeColor")
    
  'Prepare   a   text   buffer
  s = String$(255, 0)
  'What   should   we   print   on   the   button?
  GetWindowText hWnd, s, 255
  'Trim   off   nulls
  s = Left$(s, InStr(s, Chr$(0)) - 1)
    
  If va = DT_BOTTOM Then
  'Adjust   specially   for   VB's   CommandButton   control
  rct.Bottom = rct.Bottom - 4
  End If
    
  If (nState And ODS_SELECTED) = ODS_SELECTED Then
  'Button   is   in   down   state   -   offset
  'the   text
  rct.Left = rct.Left + 1
  rct.Right = rct.Right + 1
  rct.Bottom = rct.Bottom + 1
  rct.Top = rct.Top + 1
  End If
    
  DrawText hDC, s, Len(s), rct, DT_CENTER Or DT_SINGLELINE _
  Or va
    
  End Sub
    
  Public Function ExtButtonProc(ByVal hWnd As Long, _
  ByVal wMsg As Long, ByVal wParam As Long, _
  ByVal lParam As Long) As Long
    
  Dim lOldProc     As Long
  Dim di     As DRAWITEMSTRUCT
    
  lOldProc = GetProp(hWnd, "ExtBtnProc")
    
  ExtButtonProc = CallWindowProc(lOldProc, hWnd, wMsg, wParam, lParam)
    
  If wMsg = WM_DRAWITEM Then
  CopyMemory di, ByVal lParam, Len(di)
  If di.CtlType = ODT_BUTTON Then
  If GetProp(di.hwndItem, "VBTCustom") = 1 Then
  DrawButton di.hwndItem, di.hDC, di.rcItem, _
  di.itemState
    
  End If
    
  End If
    
  ElseIf wMsg = WM_DESTROY Then
  ExtButtonUnSubclass hWnd
    
  End If
    
  End Function
    
  Public Sub ExtButtonSubclass(hWndForm As Long)
    
  Dim l     As Long
    
  l = GetProp(hWndForm, "ExtBtnProc")
  If l <> 0 Then
  'Already   subclassed
  Exit Sub
  End If
    
  SetProp hWndForm, "ExtBtnProc", _
  GetWindowLong(hWndForm, GWL_WNDPROC)
  SetWindowLong hWndForm, GWL_WNDPROC, AddressOf ExtButtonProc
    
  End Sub
    
  Public Sub ExtButtonUnSubclass(hWndForm As Long)
    
  Dim l     As Long
    
  l = GetProp(hWndForm, "ExtBtnProc")
  If l = 0 Then
  'Isn't   subclassed
  Exit Sub
  End If
    
  SetWindowLong hWndForm, GWL_WNDPROC, l
  RemoveProp hWndForm, "ExtBtnProc"
    
  End Sub
    
  Public Sub SetButton(ByVal hWnd As Long, _
  ByVal lForeColor As Long, _
  Optional ByVal VAlign As TextVAligns = DT_VCENTER)
    
  Dim hWndParent     As Long
    
  hWndParent = GetParent(hWnd)
  If GetProp(hWndParent, "ExtBtnProc") = 0 Then
  ExtButtonSubclass hWndParent
  End If
    
  SetProp hWnd, "VBTCustom", 1
  SetProp hWnd, "VBTForeColor", lForeColor
  SetProp hWnd, "VBTVAlign", VAlign
    
  End Sub
    
  Public Sub RemoveButton(ByVal hWnd As Long)
    
  RemoveProp hWnd, "VBTCustom"
  RemoveProp hWnd, "VBTForeColor"
  RemoveProp hWnd, "VBTVAlign"
    
  End Sub

