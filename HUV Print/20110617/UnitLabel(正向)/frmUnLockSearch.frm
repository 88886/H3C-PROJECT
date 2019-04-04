VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmUnLockSearch 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "箱号解锁"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUnLockSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmUnLockSearch.frx":073E
   ScaleHeight     =   8505
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton cmdPrintBefore 
         Height          =   495
         Left            =   8760
         Picture         =   "frmUnLockSearch.frx":A21B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   495
         Left            =   6600
         Picture         =   "frmUnLockSearch.frx":A879
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtSN 
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
         Left            =   3120
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtBOXID 
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
         Left            =   3120
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblCodeNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "箱号(Box ID):"
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
         Left            =   360
         TabIndex        =   0
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblArea 
         BackColor       =   &H00FFFFFF&
         Caption         =   "流水号(SN Number):"
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
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   7335
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   11415
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridResult 
         Height          =   5055
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8916
         _Version        =   393216
         Rows            =   10
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmUnLockSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As New ADODB.Recordset
Dim sql As String

Private Sub cmdPrintBefore_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()

    If rec.State = 1 Then
        rec.Close
    End If
   
    'sql = "Select no=Identity(int,1,1),* Into #UNIT_temptable From tblUNIT_Unlock where BOXID in(Select BoxID From tblUNIT where sn='" & Trim(txtSN.Text) & "') or boxid='" & Trim(txtBOXID.Text) & "'; Select * From #UNIT_temptable; Drop Table #UNIT_temptable "
    sql = "select SN, BarCodeNum, VendorCode, VendorName, ContpactNo, PDate, Model, Description, BoxID, Rev, Quality, case RoHS when 1 then 'RoHS' else 'Non-RoHS' end as RoHS, UserID, CONVERT(varchar(100), PrintTime, 20) as PrintTime,UnlockUser,CONVERT(varchar(100), UnLockTime, 20) as UnLockTime from tblUNIT_Unlock where BOXID in(Select BoxID From tblUNIT_Unlock where sn='" & Trim(txtSN.Text) & "') or boxid='" & Trim(txtBOXID.Text) & "'"
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic

     Set gridResult.DataSource = rec
    
    'Call AdjustColWidth(frmUnLockSearch, gridResult)

    
    With gridResult
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 400
        .ColWidth(1) = 3000
        .ColWidth(2) = 700
        .ColWidth(3) = 1200
        .ColWidth(4) = 3000
        .ColWidth(5) = 1000
        .ColWidth(6) = 2000
        .ColWidth(7) = 2000
        .ColWidth(8) = 3000
        .ColWidth(9) = 2000
        .ColWidth(10) = 1500
        .ColWidth(11) = 1200
        .ColWidth(12) = 1500
        .ColWidth(13) = 1500
        .ColWidth(14) = 2000
        .ColWidth(15) = 1500
        .ColWidth(16) = 2000
        
        .TextMatrix(0, 1) = "产品编码"
        .TextMatrix(0, 2) = "个数"
        .TextMatrix(0, 3) = "厂商代码"
        .TextMatrix(0, 4) = "厂商名称"
        .TextMatrix(0, 5) = "合同号"
        .TextMatrix(0, 6) = "日期"
        .TextMatrix(0, 7) = "机种编码"
        .TextMatrix(0, 8) = "描述"
        .TextMatrix(0, 9) = "箱号"
        .TextMatrix(0, 10) = "软件版本"
        .TextMatrix(0, 11) = "内含数量"
        .TextMatrix(0, 12) = "RoHS"
        .TextMatrix(0, 13) = "用户"
        .TextMatrix(0, 14) = "打印时间"
        .TextMatrix(0, 15) = "解锁用户"
        .TextMatrix(0, 16) = "解锁时间"
        
   End With
   
    Me.txtBOXID.Text = ""
    Me.txtSN.Text = ""
    
End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
End Sub

Private Sub gridResult_DblClick()
    Dim i As Integer, j As Integer, nLen As Integer, nMaxLen As Integer
        With gridResult
                For i = 0 To .Rows - 1
                        nLen = LenB(StrConv(.TextMatrix(i, .Col), vbFromUnicode))
                        If nMaxLen < nLen Then
                                nMaxLen = nLen
                                j = i
                        End If
                Next i
                If nMaxLen = 0 Then Exit Sub
                Call ColWidthByCell(j, .Col)
        End With

End Sub

Private Sub ColWidthByCell(ByVal Row As Long, ByVal Col As Long)
        Dim lWidth As Long
        lWidth = (LenB(StrConv(gridResult.TextMatrix(Row, Col), vbFromUnicode)) + 1) * gridResult.FontSize * 16                         '16可按你具体情况调整
        If Row = 0 Then
                gridResult.ColWidth(Col) = lWidth
        ElseIf gridResult.ColWidth(Col) < lWidth Then
                gridResult.ColWidth(Col) = lWidth
        End If
End Sub


