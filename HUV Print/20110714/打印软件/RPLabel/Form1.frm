VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "仓库出货 RP"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11235
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   11175
      Begin VB.CheckBox chkRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "有"
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
         Left            =   7800
         TabIndex        =   28
         Top             =   2640
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkNonRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "无"
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
         Left            =   9120
         TabIndex        =   27
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmdReturn 
         Caption         =   "关闭(Close)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7560
         TabIndex        =   25
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(Cancel)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         TabIndex        =   24
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "打印(Print) &p"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   23
         Top             =   3720
         Width           =   1815
      End
      Begin VB.ComboBox ComboNote 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Form1.frx":073E
         Left            =   2280
         List            =   "Form1.frx":075D
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7560
         TabIndex        =   20
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtNumber 
         BackColor       =   &H00E0E0E0&
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
         Height          =   405
         Left            =   2280
         TabIndex        =   19
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox txtHetonghao 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7560
         MaxLength       =   8
         TabIndex        =   18
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtSupplierModel 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   17
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtDC 
         BackColor       =   &H00E0E0E0&
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
         Height          =   405
         Left            =   7560
         TabIndex        =   16
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtHW3COM 
         BackColor       =   &H00E0E0E0&
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
         Height          =   405
         Left            =   2280
         TabIndex        =   15
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtSupplierName 
         BackColor       =   &H00E0E0E0&
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
         Height          =   405
         Left            =   7560
         TabIndex        =   14
         Text            =   "飞旭电子(苏州)有限公司"
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtSupplierNo 
         BackColor       =   &H00E0E0E0&
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
         Height          =   405
         Left            =   2280
         TabIndex        =   13
         Text            =   "141078"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "认证信息RoHS:"
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
         Left            =   5520
         TabIndex        =   26
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备    注:"
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
         TabIndex        =   12
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "日    期:"
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
         Left            =   5520
         TabIndex        =   11
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "内装数量:"
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
         TabIndex        =   10
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "合同号:"
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
         Left            =   5520
         TabIndex        =   9
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "厂商规格:"
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
         TabIndex        =   8
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "厂商D/C:"
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
         Left            =   5520
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "H3C编码:"
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
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lbl2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "供应商名称:"
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
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "供应商代码:"
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
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmdNoRP 
         Caption         =   "无RP打印"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8880
         TabIndex        =   29
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtRPsn 
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
         Left            =   4320
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "仓库出货标签打印"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   5280
         TabIndex        =   22
         Top             =   120
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   405
         Left            =   2400
         Picture         =   "Form1.frx":07C7
         Top             =   120
         Width           =   2595
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RP编码:"
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
         Left            =   2880
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects

Public cnDB2  As New ADODB.Connection
Public rsDB2  As New ADODB.Recordset

Public strCommand       As String
Public strOraCon        As String
Public strDB2Con        As String
Public ITEM_KEY As String

Private Function fCntoDB2() As Integer
On Error GoTo Err_fCntoDB2

    fCntoDB2 = -1
    cnDB2.CursorLocation = adUseClient
    cnDB2.Open strDB2Con
    fCntoDB2 = 0
Exit Function
Err_fCntoDB2:
    fCntoDB2 = -1
End Function

Private Sub cmdCancel_Click()

    txtHW3COM.Text = ""
    txtDC.Text = ""
    txtSupplierModel.Text = ""
    txtHetonghao.Text = ""
    txtNumber.Text = ""
    txtDate.Text = ""
    ComboNote.ListIndex = -1
    chkRoHS.Value = 1
    txtRPsn.Text = ""
    
    txtRPsn.SetFocus
    
End Sub

Private Sub cmdNoRP_Click()
    frmNoRP.Show 1
End Sub

Private Sub cmdPrint_Click()
    If txtSupplierModel.Text = "" Then
      MsgBox "未输入厂商规格！", vbInformation + vbOKOnly, "未输入厂商规格"
      txtSupplierModel.SetFocus
      Exit Sub
    End If
    
    If txtHetonghao.Text = "" Then
      MsgBox "未输入合同号！", vbInformation + vbOKOnly, "未输入合同号"
      txtHetonghao.SetFocus
      Exit Sub
    End If
    
    If ComboNote.Text = "" Then
      MsgBox "未输入备注！", vbInformation + vbOKOnly, "未输入备注"
      ComboNote.SetFocus
      Exit Sub
    End If
    
    If Len(txtHetonghao.Text) < 8 Then
        MsgBox "合同号必须输入8位！", vbInformation + vbOKOnly, "合同号长度不对"
        txtHetonghao.SetFocus
        Exit Sub
    End If
    
    
    OpenLppx
    
    myVars.Item("Item").Value = UCase(txtHW3COM.Text)
    myVars.Item("DC").Value = UCase(txtDC.Text)
    myVars.Item("Specification").Value = UCase(Trim(txtSupplierModel.Text))
    myVars.Item("Contract").Value = UCase(Trim(txtHetonghao.Text))
    myVars.Item("Number").Value = UCase(txtNumber.Text)
    myVars.Item("Remark").Value = ComboNote.Text
    
    If chkNonRoHS.Value = 1 Then
       myObjs("ROHS").Top = 10000
       myVars.Item("sNull").Value = "NULL"
    Else
       myObjs("sNull").Top = 10000
    End If
    
    myDoc.PrintLabel 1
    myDoc.FormFeed
    UnloadLppx
    
    cmdCancel_Click
    
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass

   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "仓库出货标签.lab")
   
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub ComboNote_Click()
    chkRoHS.SetFocus
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
End Sub

Private Sub txtRPsn_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Dim strSqlDB2 As String
        Dim strSqlTmp As String
        Dim item_id As String
    
        If Trim(txtRPsn.Text) = "" Then
            MsgBox "Please enter RP Number!", vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
        
        If Len(txtRPsn.Text) < 5 Then
            MsgBox "RP条码不正确!"
            txtSN.SetFocus
            Exit Sub
        End If
        
        If Left(UCase(Trim(txtRPsn.Text)), 4) = "ITEM" Then
            item_id = Mid(Trim(txtRPsn.Text), 5)
        Else
            item_id = Trim(txtRPsn.Text)
        End If
        
        strDB2Con = "Provider=IBMDADB2.1;Password=T0mcat4Fun;Persist Security Info=True;User ID=cgsapp;Data Source=CGS"
        
        If fCntoDB2 <> 0 Then
            MsgBox "Connect to DB2 Error!", vbOKOnly, "Error"
            Exit Sub
        End If
        
        strSqlDB2 = "SELECT A.PART_NUMBER,B.ITEM_ID,C.SUPPLIER_PN as SUPPLIER_PN,B.QUANTITY " & _
        " FROM CGS.PART_NUMBER A right join CGS.ITEM B " & _
        " ON A.PART_NUMBER_KEY =B.PART_NUMBER_KEY " & _
        " left join CGS.PN_SUPPLIER C " & _
        " ON B.PN_SUPPLIER_KEY =C.PN_SUPPLIER_KEY " & _
        " where B.ITEM_KEY=" & item_id & " With UR "

    
        Set rsDB2 = cnDB2.Execute(strSqlDB2)
        
        If rsDB2.EOF = False Then
            
            txtHW3COM.Text = rsDB2("PART_NUMBER") & ""
            Dim st1 As String
            st1 = UCase(rsDB2("ITEM_ID") & "")
            txtDC.Text = get_DC(st1)
            txtSupplierModel.Text = rsDB2("SUPPLIER_PN") & ""
            txtNumber.Text = rsDB2("QUANTITY") & ""
            txtDate.Text = Date
        Else
            MsgBox "ITEM ID: " & Trim(item_id) & " not Find!"
            Exit Sub
        End If
    
        cnDB2.Close
        Set rsDB2 = Nothing
        
    End If
End Sub

Public Function get_DC(strDC As String) As String

    If InStr(strDC, ".S") > 1 Then
        get_DC = Mid(strDC, 1, InStr(strDC, ".S") - 1)
    Else
        get_DC = strDC
    End If
    

End Function

Private Sub chkNonRoHS_Click()
    If chkNonRoHS.Value = 1 Then
        chkRoHS.Value = 0
    Else
        chkRoHS.Value = 1
    End If
End Sub

Private Sub chkRoHS_Click()
    If chkRoHS.Value = 1 Then
        chkNonRoHS.Value = 0
    Else
        chkNonRoHS.Value = 1
    End If
End Sub
