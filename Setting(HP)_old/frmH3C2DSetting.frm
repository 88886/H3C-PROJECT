VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmH3C2DSetting 
   Caption         =   "H3C 2D �����趨"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   12525
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fmH3C 
      Height          =   5415
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12015
      Begin VB.TextBox txtEC 
         Height          =   375
         Left            =   9000
         TabIndex        =   51
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CheckBox chkNoPrintSV 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   30
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox chkSVPrint 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   29
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtRemark 
         Height          =   495
         Left            =   7560
         TabIndex        =   28
         Top             =   4680
         Width           =   4335
      End
      Begin VB.TextBox txtHV 
         Height          =   495
         Left            =   2400
         TabIndex        =   27
         Top             =   4680
         Width           =   2895
      End
      Begin VB.TextBox txtMS 
         Height          =   450
         Left            =   9000
         TabIndex        =   26
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox txtNAL 
         Height          =   435
         Left            =   2400
         TabIndex        =   25
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox txtGW 
         Height          =   450
         Left            =   9000
         TabIndex        =   24
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtOS 
         Height          =   450
         Left            =   2400
         TabIndex        =   23
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtDes 
         Height          =   450
         Left            =   9000
         TabIndex        =   22
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtEPN 
         Height          =   450
         Left            =   2400
         TabIndex        =   21
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtCPN 
         Height          =   450
         Left            =   9000
         TabIndex        =   20
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtSN 
         Height          =   450
         Left            =   2400
         TabIndex        =   19
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkPowerYes 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox chkPowerNo 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox cbPicture 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmH3C2DSetting.frx":0000
         Left            =   2400
         List            =   "frmH3C2DSetting.frx":0002
         TabIndex        =   16
         Top             =   2640
         Width           =   2895
      End
      Begin VB.CheckBox chkROHSYes 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10200
         TabIndex        =   15
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox chkROHSNo 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11040
         TabIndex        =   14
         Top             =   2160
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpMSValidFrom 
         Height          =   495
         Left            =   9000
         TabIndex        =   31
         Top             =   4080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16646145
         CurrentDate     =   40425
      End
      Begin MSComCtl2.DTPicker dtpValidFrom 
         Height          =   495
         Left            =   2400
         TabIndex        =   32
         Top             =   4080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16646145
         CurrentDate     =   39757
      End
      Begin VB.Label Label5 
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   50
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblCombination 
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   3000
         Width           =   11775
      End
      Begin VB.Label lblPrintSV 
         Caption         =   "�Ƿ��ӡ����汾:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   48
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label lblMSValidFrom 
         Caption         =   "�����׼��Ч��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   47
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label lblOS 
         Caption         =   "��ߴ�(MM):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblNAL 
         Caption         =   "������ɺ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label lblRemark 
         Caption         =   "��ע:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   44
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label lblHV 
         Caption         =   "Ӳ���汾:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   43
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label lblMS 
         Caption         =   "�����׼:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   42
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblValidFrom 
         Caption         =   "������Ч��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label lblGW 
         Caption         =   "ë��(KG):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   40
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblDes 
         Caption         =   "��Ʒ����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   39
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblEPN 
         Caption         =   "��Ʒ����(Ӣ��):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblCPN 
         Caption         =   "��Ʒ����(����):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   37
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblSN 
         Caption         =   "��Ʒ����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "�Ƿ�ֱ����Դ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "ͼ���:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   34
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "�Ƿ���ROHS:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   33
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(Save)"
      Height          =   495
      Left            =   10800
      TabIndex        =   12
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "���(Clear)"
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   8280
      Width           =   2895
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ѡ��(Select)"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "����(Import)"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "����(Export)"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "��ѯ(Query)"
      Height          =   495
      Left            =   9360
      TabIndex        =   3
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "����(Insert)"
      Height          =   495
      Left            =   8040
      TabIndex        =   2
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�޸�(Update)"
      Height          =   495
      Left            =   9360
      TabIndex        =   1
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(Delete)"
      Height          =   495
      Left            =   10800
      TabIndex        =   0
      Top             =   8880
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdSelect 
      Left            =   4200
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgH3C 
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4683
      _Version        =   393216
      SelectionMode   =   1
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
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "�Ƿ��ӡ����汾:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblPath 
      Caption         =   "����/����·��:"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   8400
      Width           =   1335
   End
End
Attribute VB_Name = "frmH3C2DSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim op As String
Dim xlApp As New Excel.Application
Dim xlBook As New Excel.Workbook
Dim xlSheet As New Excel.Worksheet
Dim query As Boolean

'Private Sub cbPicture_Change()
'     sql = "select Combination from [tblH3C_PictureMapping] where PictureID = '" & cbPicture.SelText & "'"
'    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
'    If rec.EOF = False Then
'        Me.lblCombination.Caption = rec.Fields(0)
'    End If
'    rec.Close
'End Sub

Private Sub cbPicture_Click()
    sql = "select Combination from [tblH3C_PictureMapping] where PictureID = '" & cbPicture.List(cbPicture.ListIndex) & "'"
    rec.Open sql, conn, adOpenKeyset, adLockReadOnly
    If rec.EOF = False Then
        Me.lblCombination.Caption = rec.Fields(0)
    End If
    rec.Close
End Sub

Private Sub cmdClear_Click()
    Dim objControl As Control
    Dim sTemp As String
    For Each objControl In Me.Controls
        If TypeOf objControl Is TextBox Then
        ' clear the text
        objControl.Text = ""
        ElseIf TypeOf objControl Is ComboBox Then
        ' reset the listindex
        objControl.ListIndex = -1
        ElseIf TypeOf objControl Is Label Then
        ' leave labels as is
        ElseIf TypeOf objControl Is CommandButton Then
        ' leave commandbuttons as is
        ElseIf TypeOf objControl Is CheckBox Then
            objControl.Value = vbUnchecked
        Else
        ' leave any other control alone
        End If
    Next

End Sub
Private Sub renovate()
   sql = "SELECT [ID],[HV],[SN],[CPN],[EPN],[Des],[OS],[GW],[ROHS],[Combination],[PictureID],[Power],[MS],[MSValidFrom],[NAL],[ValidFrom],[PrintSV],[Remark] FROM [Print].[dbo].[tblH3C_2D] order by ID,SN"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgH3C.DataSource = rec
   With mfgH3C
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 3000
        .ColWidth(4) = 2500
        .ColWidth(5) = 3000
        .ColWidth(6) = 3500
        .ColWidth(7) = 3000
        .ColWidth(8) = 1500
        .ColWidth(9) = 1000
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        .ColWidth(13) = 1500
        .ColWidth(14) = 1500
        .ColWidth(15) = 2000
        .ColWidth(16) = 2000
        .ColWidth(17) = 1500
        .ColWidth(18) = 1000
        .ColWidth(19) = 2000
        
        .TextMatrix(0, 1) = "���(ID)"
        .TextMatrix(0, 2) = "Ӳ���汾(Hardware Version)"
        .TextMatrix(0, 3) = "��Ʒ����(Model Number)"
        .TextMatrix(0, 4) = "��Ʒ����(����)(Chinese Product Name)"
        .TextMatrix(0, 5) = "��Ʒ����(Ӣ��)(English Product Name)"
        .TextMatrix(0, 6) = "��Ʒ����(Description)"
        .TextMatrix(0, 7) = "����ߴ�(Outside Size)"
        .TextMatrix(0, 8) = "ë��(Gross Weight)"
        .TextMatrix(0, 9) = "ROHS"
        .TextMatrix(0, 10) = "ͼ�����"
        .TextMatrix(0, 11) = "ͼ��ID"
        .TextMatrix(0, 12) = "ֱ����Դ"
        .TextMatrix(0, 13) = "�����׼(China MFG Standards)"
        .TextMatrix(0, 14) = "�����׼��Ч��(Valid From)"
        .TextMatrix(0, 15) = "������ɺ�(China N.A.L.)"
        .TextMatrix(0, 16) = "������Ч��(Valid From)"
        .TextMatrix(0, 17) = "�Ƿ��ӡ����汾"
        .TextMatrix(0, 18) = "��ע(Remark)"
   End With
   rec.Close
End Sub

Private Sub cmdDelete_Click()
     If mfgH3C.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫɾ������!", vbInformation + vbOKOnly, "δѡ����"
      Exit Sub
   End If
   sql = "delete from tblH3C_2d where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and SN='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 3) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "ɾ��H3C�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "ɾ��ʧ��"
   End If
   MsgBox "ɾ��H3C�趨���ϳɹ�!", vbInformation + vbOKOnly, "ɾ���ɹ�"
   renovate
End Sub

Private Sub cmdInsert_Click() '
    op = "Insert"
    cmdClear_Click
    txtSN.SetFocus
End Sub

Private Sub cmdQuery_Click()
      If txtSN.Enabled = False Then
      MsgBox "�밴������ť��վͿ������ѯ����!", vbOKOnly + vbInformation, "�����ѯ����"
    End If
    If rec.State = 1 Then
        rec.Close
     End If
     sql = "select * from tblH3C_2D Where 1=1"
     If txtSN.Text <> "" Then
        sql = sql & " and SN like '%" & txtSN.Text & "%'"
     End If
     If txtCPN.Text <> "" Then
        sql = sql & " and CPN like '%" & txtCPN.Text & "%'"
     End If
     If txtEPN.Text <> "" Then
        sql = sql & " and EPN='%" & txtEPN.Text & "%'"
     End If
     If txtDes.Text <> "" Then
        sql = sql & " and Des like '%" & txtDes.Text & "%'"
     End If
     If txtOS.Text <> "" Then
        sql = sql & " and OS like '%" & txtOS.Text & "%'"
     End If
     If txtGW.Text <> "" Then
        sql = sql & " and GW like '%" & txtGW.Text & "%'"
     End If
'     Dim CE, WEEE, ChinaRoHS, RoHS As String
'     If chkCE.Value = 1 Then
'        CE = "CE"
'     ElseIf chkNonCE.Value = 1 Then
'        CE = "/"
'     End If
'     If chkWEEE.Value = 1 Then
'        WEEE = "WEEE"
'     ElseIf chkNonWEEE.Value = 1 Then
'        WEEE = "/"
'     End If
'     If chkChinaRoHS.Value = 1 Then
'        ChinaRoHS = "China RoHS"
'     ElseIf chkNonChinaRoHS.Value = True Then
'        ChinaRoHS = "/"
'     End If
'     If optH3CRoHS.Value = 1 Then
'        RoHS = "H3C RoHS"
'     ElseIf opt3COMRoHS.Value = 1 Then
'        RoHS = "3COM RoHS"
'     ElseIf optNonRoHS.Value = 1 Then
'        RoHS = "/"
'     End If
'     If CE <> "" Then
'        sql = sql & " and CE='" & CE & "'"
'     End If
'     If WEEE <> "" Then
'        sql = sql & " and WEEE='" & WEEE & "'"
'     End If
'     If ChinaRoHS <> "" Then
'        sql = sql & " and ChinaRoHS='" & ChinaRoHS & "'"
'     End If
'     If RoHS <> "" Then
'        sql = sql & " and RoHS='" & RoHS & "'"
'     End If
'     If txtMS.Text <> "" Then
'        sql = sql & " and MS='" & txtMS.Text & "'"
'     End If
'     If txtNAL.Text <> "" Then
'        sql = sql & " and NAL='" & txtNAL.Text & "'"
'     End If
'     If txtChangNAL.Text <> "" Then
'        sql = sql & " and ChangNAL='" & txtChangNAL.Text & "'"
'     End If
'      If txtHV.Text <> "" Then
'        sql = sql & " and HV='" & txtHV.Text & "'"
'     End If
'     If txtRemark.Text <> "" Then
'        sql = sql & " and Remark='" & txtRemark.Text & "'"
'     End If
     sql = sql & " order by ID,SN"
     rec.Open sql, conn, adOpenKeyset, adLockOptimistic
     Set mfgH3C.DataSource = rec
End Sub

Private Sub cmdSave_Click()
   If Trim(txtSN.Text) = "" Then
      MsgBox "��Ʒ���벻��Ϊ��!!", vbExclamation + vbOKOnly, "��Ʒ�����"
      txtSN.SetFocus
      Exit Sub
   End If
   If txtCPN.Text = "" Then
       MsgBox "��Ʒ����(����)����Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ����(����)��"
       txtCPN.SetFocus
       Exit Sub
   End If
   If txtEPN.Text = "" Then
      MsgBox "��Ʒ����(Ӣ��)����Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ����(Ӣ��)��"
      txtEPN.SetFocus
      Exit Sub
   End If
   If txtDes.Text = "" Then
      MsgBox "��Ʒ��������Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ������"
      txtDes.SetFocus
      Exit Sub
   End If
   If txtOS.Text = "" Then
      MsgBox "��ߴ粻��Ϊ��!", vbExclamation + vbOKOnly, "��ߴ��"
      txtOS.SetFocus
      Exit Sub
   End If
   If txtOS.Text = "/" Then
      MsgBox "����ߴ���ά��N/A!", vbExclamation + vbOKOnly, "����ߴ�"
      txtOS.SetFocus
      Exit Sub
   End If
   If txtOS.Text = "n/a" Then
      txtOS.Text = UCase(txtOS.Text)
   End If

   If txtOS.Text <> "N/A" Then
   
        txtOS.Text = LTrim(RTrim(txtOS.Text))
      
        If Right(txtOS.Text, 2) <> "mm" Then
            MsgBox "��ߴ��ʽ����!", vbExclamation + vbOKOnly, "��ߴ����"
            txtOS.SetFocus
            Exit Sub
        End If
        
        If InStr(txtOS.Text, "mmm") > 0 Then
            MsgBox "��ߴ��ʽ����!", vbExclamation + vbOKOnly, "��ߴ����"
            txtOS.SetFocus
            Exit Sub
        End If
   End If
   
   
   
   'If txtGW.Text = "" Then
   '   MsgBox "ë�ز���Ϊ��!", vbExclamation + vbOKOnly, "ë�ؿ�"
   '   txtGW.SetFocus
   '   Exit Sub
   'End If
    If Trim(txtGW.Text) <> "" Then
        If UCase(Right(Trim(txtGW.Text), 2)) <> "KG" Then
           MsgBox "ë�ر�����ϵ�λkg!", vbExclamation + vbOKOnly, "ë�ص�λ��"
           txtGW.SetFocus
           Exit Sub
        End If
   End If
   If txtMS.Text = "" Then
      MsgBox "�����׼����Ϊ��!", vbExclamation + vbOKOnly, "�����׼��"
      txtMS.SetFocus
      Exit Sub
   End If
   If txtNAL.Text = "" Then
      MsgBox "������ɺŲ���Ϊ��!", vbExclamation + vbOKOnly, "������ɺſ�"
      txtNAL.SetFocus
      Exit Sub
   End If
   If txtHV.Text = "" Then
      MsgBox "Ӳ���汾����Ϊ��!", vbExclamation + vbOKOnly, "Ӳ���汾��"
      txtHV.SetFocus
      Exit Sub
   End If
   If chkSVPrint.Value = 0 And chkNoPrintSV.Value = 0 Then
      MsgBox "�Ƿ��ӡ����汾����Ϊ��!", vbExclamation + vbOKOnly, "������汾��"
      txtHV.SetFocus
      Exit Sub
   End If
   
   If Me.cbPicture.ListIndex = -1 Then
      MsgBox "ͼ�α��ѡ����ܲ�ѡ!", vbExclamation + vbOKOnly, "ͼ�α�Ų���Ϊ��"
      cbPicture.SetFocus
      Exit Sub
   End If
   If Me.lblCombination.Caption = "" Then
      MsgBox "ͼ�α������Ӧ����ϲ�����!", vbExclamation + vbOKOnly, "ͼ��������ݲ���Ϊ��"
      cbPicture.SetFocus
      Exit Sub
   End If
   If Trim(Me.txtEC.Text) = "" Then
      MsgBox "������Ϣ����Ϊ��", vbExclamation + vbOKOnly, "������Ϣ����Ϊ��"
      Me.txtEC.SetFocus
      Exit Sub
   End If
   
   
   Dim RoHS, SVPrint, Power As String
   If chkSVPrint.Value = 1 Then
      SVPrint = "Y"
   ElseIf chkNoPrintSV.Value = 1 Then
      SVPrint = "N"
   End If
   
   If chkROHSYes.Value = 1 Then
      RoHS = "Y"
   ElseIf Me.chkROHSNo.Value = 1 Then
      RoHS = "N"
   End If
   
   If chkPowerYes.Value = 1 Then
      Power = "Y"
   ElseIf chkPowerNo.Value = 1 Then
      Power = "N"
   End If

  ' If optH3CRoHS.Value = True Then
  '    RoHS = "H3C RoHS"
  ' ElseIf opt3COMRoHS.Value = True Then
  '    RoHS = "3COM RoHS"
  ' ElseIf optNonRoHS.Value = True Then
  '    RoHS = "/"
  ' End If
  
  txtGW.Text = LCase(Trim(txtGW.Text))
   
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from tblH3C where SN='" & Trim(txtSN.Text) & "' and HV='" & Trim(txtHV.Text) & "' "
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "��Ʒ����&�汾�Ѵ���!", vbExclamation + vbOKOnly, "��Ʒ����ظ�"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close

      sql = "Insert into tblH3C_2D(ID, HV, SN, CPN, EPN, Des, OS, GW,[ROHS],[Combination],[PictureID],[Power] MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark, EC) " & _
            "Values(" & getmaxID("tblH3C") & ",'" & Trim(txtHV.Text) & "','" & Trim(txtSN.Text) & "','" & Trim(txtCPN.Text) & "','" & Trim(txtEPN.Text) & "','" & Trim(txtDes.Text) & "','" & Trim(txtOS.Text) & "','" & Trim(txtGW.Text) & "','" & RoHS & "','" & Me.lblCombination.Caption & "','" & Me.cbPicture.Text & "','" & Power & "'," & _
            "'" & txtMS.Text & "','" & dtpMSValidFrom.Value & "','" & txtNAL.Text & "','" & dtpValidFrom.Value & "','" & SVPrint & "','" & txtRemark.Text & "','" & Me.txtEC.Text & "')"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "����H3C�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "����ʧ��"
      End If
      MsgBox "����H3C�趨���ϳɹ�!", vbInformation + vbOKOnly, "�����ɹ�"
      renovate
      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update tblH3C_2D set CPN='" & Trim(txtCPN.Text) & "',EPN='" & Trim(txtEPN.Text) & "',Combination='" & Trim(Me.lblCombination.Caption) & "',PictureID='" & Trim(Me.cbPicture.Text) & "',Power='" & Power & "',RoHS='" & RoHS & "'," & _
            "MS='" & txtMS.Text & "',MSValidFrom='" & dtpMSValidFrom.Value & "',NAL='" & txtNAL.Text & "',ValidFrom='" & dtpValidFrom.Value & "',HV='" & txtHV.Text & "',PrintSV='" & SVPrint & "',Remark='" & txtRemark.Text & "'" & ",EC = '" & Me.txtEC.Text & "'" & _
            " where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and SN='" & Trim(txtSN.Text) & "'"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "�޸�H3C�趨����ʧ��!" & "ԭ����" & status, vbExclamation + vbOKOnly, "�޸�ʧ��"
      End If
      MsgBox "�޸�H3C�趨���ϳɹ�!", vbInformation + vbOKOnly, "�޸ĳɹ�"
      renovate
      cmdClear_Click
   End If
   renovate
End Sub

Private Sub cmdUpdate_Click()
    op = "Update"
    txtSN.SetFocus
End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   renovate
   GetPictureMapping
End Sub

Private Sub mfgH3C_Click()
     If mfgH3C.RowSel > 0 Then
      txtHV.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 2)
      txtSN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 3)
      txtCPN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 4)
      txtEPN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 5)
      txtDes.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 6)

      txtOS.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 7)
      txtGW.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 8)
      If UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 9)) = "Y" Then
        chkROHSYes.Value = 1
        chkROHSNo.Value = 0
      ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 9)) = "N" Then
        chkROHSYes.Value = 0
        chkROHSNo.Value = 1
      End If
     
      Me.lblCombination.Caption = UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 10))
      
      If IsNull(mfgH3C.TextMatrix(mfgH3C.RowSel, 11)) = False And Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 11)) <> "" Then
        Dim pictureID As String
        pictureID = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 11))
        Me.cbPicture.ListIndex = CheckIfExistInCombo(cbPicture, pictureID)
      End If
      
      If IsNull(mfgH3C.TextMatrix(mfgH3C.RowSel, 12)) And Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 12)) <> "" Then
        If Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 12)) = "Y" Then
            chkPowerYes.Value = 1
            chkPowerNo.Value = 0
        Else
            chkPowerYes.Value = 0
            chkPowerNo.Value = 1
        End If
      Else
        chkPowerYes.Value = 0
        chkPowerNo.Value = 0
      End If
      
      txtMS.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 13)
      dtpMSValidFrom.Value = mfgH3C.TextMatrix(mfgH3C.RowSel, 14)
      
      txtNAL.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 15)
      dtpValidFrom.Value = mfgH3C.TextMatrix(mfgH3C.RowSel, 16)
      
      If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 17))) = "Y" Then
        chkSVPrint.Value = 1
        chkNoPrintSV.Value = 0
      ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 18))) = "N" Then
        chkSVPrint.Value = 0
        chkNoPrintSV.Value = 1
      End If
      txtRemark.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 18)
   End If
End Sub

Public Function CheckIfExistInCombo(objCombo As Object, TextToFind As String) As Integer
    Dim NumOfItems As Variant 'The Number Of Items In ComboBox
    Dim IndexNum As Integer 'Index
    NumOfItems = objCombo.ListCount
    For IndexNum = 0 To NumOfItems - 1
        If objCombo.List(IndexNum) = TextToFind Then
            CheckIfExistInCombo = IdexNum
            Exit Function
        End If
    Next IndexNum
    
    CheckIfExistInCombo = -1
End Function

Public Function GetPictureMapping()
    Dim rs As New Recordset
    sql = "select PictureID,Combination from [tblH3C_PictureMapping]"
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    With rs
    Do While Not .EOF
        Me.cbPicture.AddItem (rs!pictureID)
'        Me.cbPicture.ItemData(cbPicture.NewIndex) = rs!Combination
        .MoveNext
    Loop
    End With
    Me.cbPicture.Refresh
    rs.Close
End Function
