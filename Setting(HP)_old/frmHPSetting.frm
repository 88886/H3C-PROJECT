VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmHPSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HP Setting"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHPSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdSelect 
      Left            =   2520
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgHP 
      Height          =   2775
      Left            =   0
      TabIndex        =   36
      Top             =   5400
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   4895
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10740
      TabIndex        =   31
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9300
      TabIndex        =   30
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "ȷ��(Confirm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7620
      TabIndex        =   29
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(Delete)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10740
      TabIndex        =   28
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�޸�(Update)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9300
      TabIndex        =   27
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "����(Insert)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7620
      TabIndex        =   26
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "��ѯ(Query)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5940
      TabIndex        =   25
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "����(Export)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3420
      TabIndex        =   24
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "����(Import)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3420
      TabIndex        =   23
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ѡ��(Select)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1380
      TabIndex        =   22
      Top             =   9240
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
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
      Height          =   495
      Left            =   180
      TabIndex        =   21
      Top             =   8640
      Width           =   3015
   End
   Begin VB.Frame fmH3C 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.CheckBox chkNo 
         Caption         =   "��"
         Height          =   345
         Left            =   10680
         TabIndex        =   45
         Top             =   4680
         Width           =   615
      End
      Begin VB.CheckBox chkYes 
         Caption         =   "��"
         Height          =   375
         Left            =   9120
         TabIndex        =   44
         Top             =   4680
         Width           =   615
      End
      Begin VB.ComboBox cb5000 
         Height          =   465
         ItemData        =   "frmHPSetting.frx":2E1A
         Left            =   9120
         List            =   "frmHPSetting.frx":2E2A
         TabIndex        =   42
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox txtHPSNProduct 
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
         Left            =   2280
         TabIndex        =   40
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox txtHPdesc2 
         Height          =   465
         Left            =   2280
         TabIndex        =   38
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtHPdesc1 
         Height          =   495
         Left            =   2280
         TabIndex        =   37
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtPack 
         Height          =   495
         Left            =   9120
         TabIndex        =   35
         Top             =   3360
         Width           =   2775
      End
      Begin VB.TextBox txtRN 
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
         Height          =   450
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtCode 
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
         Left            =   9120
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtH3CSNIII 
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
         Left            =   2280
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtHPPN 
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
         Left            =   9120
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtHPP 
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
         Left            =   9120
         TabIndex        =   5
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtHPSNIII 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9120
         TabIndex        =   4
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtHPGtinNum 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9120
         TabIndex        =   3
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox txtHPGtin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtH3CType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "�ϴ���Դ����:"
         Height          =   495
         Left            =   6840
         TabIndex        =   43
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "5000״̬:"
         Height          =   495
         Left            =   7200
         TabIndex        =   41
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "HP SN Product"
         Height          =   735
         Left            =   120
         TabIndex        =   39
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label LabPack 
         Caption         =   "Pack Label:"
         Height          =   495
         Left            =   6840
         TabIndex        =   34
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lblHPSNIII 
         Caption         =   "HP SN-III:"
         Height          =   375
         Left            =   6840
         TabIndex        =   32
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblrn 
         Caption         =   "Row Number:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblcode 
         Caption         =   "H3C BOM code"
         Height          =   375
         Left            =   6840
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblH3CSN3 
         Caption         =   "H3C SN-III"
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblHPPN 
         Caption         =   "HP PN(text):"
         Height          =   375
         Left            =   6840
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblHPP 
         Caption         =   "HP Product(text):"
         Height          =   735
         Left            =   6840
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblValidFrom 
         Caption         =   "HP desc2:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lbldesc1 
         Caption         =   "HP desc 1:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblHPGtinN 
         Caption         =   "HP GTIN Number:"
         Height          =   375
         Left            =   6360
         TabIndex        =   12
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label lblHPGTIN 
         Caption         =   "HP GTIN:"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label lblH3CType 
         Caption         =   "H3C Type:"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Caption         =   "HP desc 1:"
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblPath 
      Caption         =   "����/����·��:"
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
      Left            =   180
      TabIndex        =   20
      Top             =   8160
      Width           =   2175
   End
End
Attribute VB_Name = "frmHPSetting"
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

Private Sub enable()
   txtRN.Enabled = False
   txtRN.BackColor = &HFFFFFF
   txtCode.Enabled = True
   txtCode.BackColor = &HFFFFFF
   txtH3CSNIII.Enabled = True
   txtH3CSNIII.BackColor = &HFFFFFF
   txtHPSNProduct.Enabled = True
   txtHPSNProduct.BackColor = &HFFFFFF
   txtHPPN.Enabled = True
   txtHPPN.BackColor = &HFFFFFF
   
   txtHPP.Enabled = True
   txtHPP.BackColor = &HFFFFFF
   
   txtHPSNIII.Enabled = True
   txtHPSNIII.BackColor = &HFFFFFF
   txtHPdesc1.Enabled = True
   txtHPdesc1.BackColor = &HFFFFFF
   txtHPdesc2.Enabled = True
   txtHPdesc2.BackColor = &HFFFFFF
   txtHPGtin.Enabled = True
   txtHPGtin.BackColor = &HFFFFFF
   txtHPGtinNum.Enabled = True
   txtHPGtinNum.BackColor = &HFFFFFF
   txtH3CType.Enabled = True
   txtH3CType.BackColor = &HFFFFFF
   txtPack.Enabled = True
   txtPack.BackColor = &HFFFFFF
   
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = False
   cmdUpdate.Enabled = False
   cmdDelete.Enabled = False
   cmdConfirm.Enabled = True
   cmdCancel.Enabled = True
End Sub

Private Sub unable()
   txtRN.Enabled = False
   txtRN.BackColor = &HE0E0E0
   txtCode.Enabled = False
   txtCode.BackColor = &HE0E0E0
   txtH3CSNIII.Enabled = False
   txtH3CSNIII.BackColor = &HE0E0E0
   txtHPSNProduct.Enabled = False
   txtHPSNProduct.BackColor = &HE0E0E0
   txtHPPN.Enabled = False
   txtHPPN.BackColor = &HE0E0E0

   txtHPP.Enabled = False
   txtHPP.BackColor = &HE0E0E0
   txtHPSNIII.Enabled = False
   txtHPSNIII.BackColor = &HE0E0E0
   txtHPdesc1.Enabled = False
   txtHPdesc1.BackColor = &HE0E0E0
   txtHPdesc2.Enabled = False
   txtHPdesc2.BackColor = &HE0E0E0
   txtHPGtin.Enabled = False
   txtHPGtin.BackColor = &HE0E0E0
   txtHPGtinNum.Enabled = False
   txtHPGtinNum.BackColor = &HE0E0E0
   txtH3CType.Enabled = False
   txtH3CType.BackColor = &HE0E0E0
   txtPack.Enabled = False
   txtPack.BackColor = &HE0E0E0
   
   cmdSelect.Enabled = True
   cmdImport.Enabled = True
   cmdExport.Enabled = True
   cmdQuery.Enabled = True
   cmdInsert.Enabled = True
   cmdUpdate.Enabled = True
   cmdDelete.Enabled = True
   cmdConfirm.Enabled = False
   cmdCancel.Enabled = False
End Sub

Private Sub chkNo_Click()
    If chkNo.Value = 1 Then
        chkYes.Value = 0
    End If
End Sub

Private Sub chkYes_Click()
    If chkYes.Value = 1 Then
        chkNo.Value = 0
    End If
End Sub

Private Sub cmdCancel_Click()
   unable
   op = ""
End Sub

Private Sub cmdConfirm_Click()
   Dim ftStatus, uploadPW As String
   If chkYes.Value + chkNo.Value = 2 Or chkYes.Value + chkNo.Value = 0 Then
      MsgBox "��ѡ���Ƿ��ϴ���Դ����!", vbExclamation + vbOKOnly, "�Ƿ��ϴ���Դ����δѡ��"
      Exit Sub
   End If
   
'   If txtRN.Text = "" Then
'      MsgBox "Row Number����Ϊ��!!", vbExclamation + vbOKOnly, "Row Number��"
'      txtRN.SetFocus
'      Exit Sub
'   End If
   If txtCode.Text = "" Then
       MsgBox "H3C BOM Code����Ϊ��!", vbExclamation + vbOKOnly, "H3C BOM Code��"
       txtCode.SetFocus
       Exit Sub
   End If
   If txtH3CSNIII.Text = "" Then
      MsgBox "H3C Serial Number-III����Ϊ��!", vbExclamation + vbOKOnly, "H3C Serial Number-III��"
      txtH3CSNIII.SetFocus
      Exit Sub
   End If
   'If txtHPPN.Text = "" Then
     ' MsgBox "HP Part Number(text)����Ϊ��!", vbExclamation + vbOKOnly, "HP Part Number(text)��"
    '  txtHPPN.SetFocus
   '   Exit Sub
  ' End If
   'If txtHPP.Text = "" Then
   '   MsgBox "HP Product(text)����Ϊ��!", vbExclamation + vbOKOnly, "HP Product(text)��"
   '   txtHPP.SetFocus
   '   Exit Sub
  ' End If
   If txtHPSNIII.Text = "" Then
      MsgBox "HP Serial Number-III����Ϊ��!", vbExclamation + vbOKOnly, "HP Serial Number-III��"
      txtHPSNIII.SetFocus
      Exit Sub
   End If
   If txtHPdesc1.Text = "" Then
      MsgBox "HP Product decription(1)����Ϊ��!", vbExclamation + vbOKOnly, "HP Product decription(1)��"
      txtHPdesc1.SetFocus
      Exit Sub
   End If
   If txtHPGtin.Text = "" Then
      MsgBox "HP-GTIN����Ϊ��!", vbExclamation + vbOKOnly, "HP-GTIN��"
      txtHPGtin.SetFocus
      Exit Sub
   End If

   'If txtHPGtinNum.Text = "" Then
    '  MsgBox "HP-GTIN Number����Ϊ��!", vbExclamation + vbOKOnly, "HP-GTIN Number��"
   '   txtHPGtinNum.SetFocus
   '   Exit Sub
  ' End If

   If txtH3CType.Text = "" Then
      MsgBox "H3C Type����Ϊ��!", vbExclamation + vbOKOnly, "H3C Type��"
      txtH3CType.SetFocus
      Exit Sub
   End If
  
  If txtPack.Text = "" Then
      MsgBox "Pack Label����Ϊ��!", vbExclamation + vbOKOnly, "Pack Label��"
      txtPack.SetFocus
      Exit Sub
   End If
    
  If Len(Trim(txtHPP.Text)) <= 6 Then
      txtHPSNProduct.Text = Trim(txtHPP.Text)
  End If
  
   If Me.cb5000.ListIndex <= -1 Then
        MsgBox "5000״̬û��ѡ��!", vbExclamation + vbOKOnly, "��ѡ��5000״̬��һ��ѡ��"
         Me.cb5000.SetFocus
         Exit Sub
    End If
    
'    Y��N��NA,TBD
    If Me.cb5000.ListIndex = 0 Then
        ftStatus = "Y"
    ElseIf Me.cb5000.ListIndex = 1 Then
        ftStatus = "N"
    ElseIf Me.cb5000.ListIndex = 2 Then
        ftStatus = "NA"
    ElseIf Me.cb5000.ListIndex = 3 Then
        ftStatus = "TBD"
    End If
    
    If chkYes.Value = 1 Then
        uploadPC = "1"
    ElseIf chkNo.Value = 1 Then
        uploadPC = "0"
    End If
  
  
  
   If op = "Insert" Then
   
      '====================================
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from hp where hp_sn_iii='" & txtHPSNIII.Text & "'"
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) >= 2 Then
         MsgBox "HP Serial Number-III�Ѵ��ڶ��!"
         txtHPPN.SetFocus
         Exit Sub
      End If
      rcd.Close
      '====================================
      
      '====================================
'      Dim rcd2 As New ADODB.Recordset
'      sql = "select Count(*) from hp where row_number='" & Trim(txtRN.Text) & "'"
'      rcd2.Open sql, conn, adOpenKeyset, adLockOptimistic
'      If rcd2.Fields(0) >= 1 Then
'         MsgBox "Row Number �ظ�,�����������µ�Row Number!"
'         txtHPPN.SetFocus
'         Exit Sub
'      End If
'      rcd2.Close
      '====================================
      
      
      sql = "insert hp(h3c_bom_code,h3c_sn_iii,hp_pn,hp_product,hp_sn_iii,hp_desc1,hp_desc2,hp_gtin,hp_gtin_number,h3c_type,pack_label,hpsnproduct,[5000_status],upload_power_code)" & _
            "Values('" & txtCode.Text & "','" & txtH3CSNIII.Text & "','" & txtHPPN.Text & "','" & txtHPP.Text & "','" & txtHPSNIII.Text & "'," & _
            "'" & txtHPdesc1.Text & "','" & txtHPdesc2.Text & "','" & txtHPGtin.Text & "','" & Trim(txtHPGtinNum.Text) & "','" & txtH3CType.Text & "','" & txtPack.Text & "','" & txtHPSNProduct.Text & "','" & ftStatus & "'," & uploadPC & ")"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "����HP�趨����ʧ��!" & "ԭ����" & status
      Else
         MsgBox "����HP�趨���ϳɹ�!"
      End If
     
      renovate
      cmdInsert_Click
   ElseIf op = "Update" Then
    
   
      sql = "Update hp set h3c_sn_iii='" & txtH3CSNIII.Text & "',hp_product='" & txtHPP.Text & "',[5000_status] = '" & ftStatus & "',upload_power_code = " & uploadPC & "," & _
            "hp_pn='" & txtHPPN.Text & "', hp_desc1='" & txtHPdesc1.Text & "',hp_desc2='" & txtHPdesc2.Text & "',hp_gtin='" & txtHPGtin.Text & "',hp_gtin_number='" & Trim(txtHPGtinNum.Text) & "',h3c_type='" & txtH3CType.Text & "',pack_label='" & txtPack.Text & "',hpsnproduct='" & Trim(txtHPSNProduct.Text) & "'" & _
            " where hp_sn_iii='" & txtHPSNIII.Text & "' and  h3c_bom_code='" & txtCode.Text & "' "
     ' MsgBox (sql)
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "�޸�HP�趨����ʧ��!" & "ԭ����" & status
      Else
         MsgBox "�޸�HP�趨���ϳɹ�!"
      End If
      renovate
      cmdCancel_Click
   End If
   renovate
End Sub

Private Sub cmdDelete_Click()
   If mfgHP.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫɾ������!"
      Exit Sub
   End If
   If MsgBox("ȷ��Ҫɾ����ǰѡ�е�������", vbYesNo, "����ɾ��ȷ��") = vbYes Then
    'sql = "delete from hp where hp_pn='" & mfgHP.TextMatrix(mfgHP.RowSel, 4) & "'"
    sql = "delete from hp where hp_sn_iii='" & mfgHP.TextMatrix(mfgHP.RowSel, 6) & "' and h3c_bom_code='" & mfgHP.TextMatrix(mfgHP.RowSel, 2) & "'"
    status = Connect.excuteUpdate(sql)
    If status <> "" Then
       MsgBox "ɾ��HP�趨����ʧ��!" & "ԭ����" & status
    End If
    MsgBox "ɾ��HP�趨���ϳɹ�!"
    renovate
   End If
   
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If mfgHP.Rows = 0 Then
      MsgBox "�����Ͽɻ��"
      Exit Sub
   End If
   If txtPath.Text <> "" Then
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets.Item(1)
       For i = 0 To mfgHP.Rows - 1
         For j = 1 To mfgHP.Cols - 1
          xlSheet.Cells(i + 1, j) = mfgHP.TextMatrix(i, j)
       Next j
      Next i
      xlBook.SaveAs (txtPath.Text)
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "�����EXCEL���ϳɹ�!!"
    End If
End Sub

Private Sub cmdImport_Click()

MsgBox "�˹����Ѿ�ȡ��!"
Exit Sub
      
   If txtPath.Text = "" Then
      MsgBox "����·������Ϊ��!"
      Exit Sub
   End If
   Dim action As Integer
   Dim info As Boolean
   info = True
   Set xlBook = xlApp.Workbooks.Open(txtPath.Text)
      For i = 1 To xlBook.Sheets.Count
       Set xlSheet = xlBook.Sheets.Item(i)
       For j = 2 To xlSheet.Rows.Count
        r = xlSheet.Cells(j, 1)
        If r = "" Then
           Exit For
        Else
          Dim cellValue As String
          Dim isexist As Boolean
          If xlSheet.Cells(j, 12) = "" Then
             MsgBox "�������ϸ�ʽ����ȷ!"
             Exit Sub
          End If
          isexist = False
          For K = 1 To 12
           If K = 6 Then
             cellValue = xlSheet.Cells(j, K)
             If cellValue = "" Then
                MsgBox "�������ϸ�ʽ����ȷ!"
                Exit Sub
             End If
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from hp where hp_sn_iii='" & cellValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("HP Serial Number-III�Ѵ���!", vbAbortRetryIgnore + vbExclamation, "�����ظ�")
                End If
                
                If action = vbAbort Then
                   MsgBox "���ϵ�������ֹ!!"
                   rcd.Close
                   Exit Sub
                ElseIf action = vbIgnore And info = True Then
                   MsgBox "�ظ���Ʒ������ϲ��ᵼ��,���Ե�..!!"
                   rcd.Close
                   info = False
                   Exit For
                ElseIf action = vbRetry And info = True Then
                   MsgBox "�ظ���Ʒ������ϻ��Զ�����,���Ե�..!!"
                   info = False
                End If
                isexist = True
             Else
                isexist = False
             End If
             rcd.Close
            End If
            
            
            If K = 12 Then
               If action = vbRetry Then
                   sql = "Update hp set row_number='" & xlSheet.Cells(j, 1) & "',h3c_bom_code='" & xlSheet.Cells(j, 2) & "',h3c_sn_iii='" & xlSheet.Cells(j, 3) & "',hp_product='" & xlSheet.Cells(j, 5) & "'," & _
                        "hp_pn='" & xlSheet.Cells(j, 4) & "',hp_desc1='" & xlSheet.Cells(j, 7) & "',hp_desc2='" & xlSheet.Cells(j, 8) & "',hp_gtin='" & xlSheet.Cells(j, 9) & "',hp_gtin_number='" & xlSheet.Cells(j, 10) & "',h3c_type='" & xlSheet.Cells(j, 11) & "',pack_label='" & xlSheet.Cells(j, 12) & "',hpsnproduct='" & xlSheet.Cells(j, 13) & "'" & _
                        " where hp_sn_iii='" & xlSheet.Cells(j, 6) & "'"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "�޸�HP�趨����ʧ��!" & "ԭ����" & status
                   End If
'                   MsgBox "�޸�HUAWEI�趨���ϳɹ�!"
               ElseIf isexist = False Then
               'MsgBox (xlSheet.Cells(j, 11))
                   sql = "Insert into hp(row_number,h3c_bom_code,h3c_sn_iii,hp_pn,hp_product,hp_sn_iii,hp_desc1,hp_desc2,hp_gtin,hp_gtin_number,h3c_type,pack_label,hpsnproduct) Values('" & xlSheet.Cells(j, 1) & "','" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "','" & xlSheet.Cells(j, 4) & "','" & xlSheet.Cells(j, 5) & "','" & xlSheet.Cells(j, 6) & "','" & xlSheet.Cells(j, 7) & "','" & xlSheet.Cells(j, 8) & "','" & xlSheet.Cells(j, 9) & "','" & xlSheet.Cells(j, 10) & "','" & xlSheet.Cells(j, 11) & "','" & xlSheet.Cells(j, 12) & "','" & xlSheet.Cells(j, 13) & "')"
                  'MsgBox (sql)
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "����HP�趨����ʧ��!" & "ԭ����" & status
                   End If
'                   MsgBox "����HUAWEI�趨���ϳɹ�!"
               End If
           End If
         Next K
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "HP�趨���ϵ���ɹ�!"
      renovate
End Sub

Private Sub cmdInsert_Click()
   enable
   txtRN.Text = ""
   txtCode.Text = ""
   txtH3CSNIII.Text = "21"
   txtHPSNProduct.Text = ""
   txtHPPN.Text = ""

   txtHPP.Text = ""

   txtHPSNIII.Text = ""
   txtHPdesc1.Text = ""
   txtHPdesc2.Text = ""
   txtHPGtin.Text = ""
   txtHPGtinNum.Text = ""
   txtH3CType.Text = ""
   txtPack.Text = ""
   cb5000.ListIndex = -1
   chkYes.Value = 0
   chkNo.Value = 0
   
   op = "Insert"
End Sub

Private Sub cmdQuery_Click()
    If txtCode.Enabled = False Then
        MsgBox "�밴������ť��վͿ������ѯ����!", vbOKOnly + vbInformation, "�����ѯ����"
    End If
    
   If rec.State = 1 Then
      rec.Close
   End If
   sql = "select * from hp Where 1=1"
'   If txtRN.Text <> "" Then
'      sql = sql & " and row_number like '%" & txtRN.Text & "%'"
'   End If
   If txtCode.Text <> "" Then
      sql = sql & " and h3c_bom_code like '%" & txtCode.Text & "%'"
   End If
   If txtH3CSNIII.Text <> "" Then
      sql = sql & " and h3c_sn_iii like '%" & txtH3CSNIII.Text & "%'"
   End If
   If txtHPSNProduct.Text <> "" Then
      sql = sql & " and hpsnproduct like '%" & txtHPSNProduct.Text & "%'"
   End If
   If txtHPPN.Text <> "" Then
      sql = sql & " and hp_pn like '%" & txtHPPN.Text & "%'"
   End If
   If txtHPSNIII.Text <> "" Then
      sql = sql & " and hp_sn_iii like '%" & txtHPSNIII.Text & "%'"
   End If

   sql = sql & " order by row_number"
   'MsgBox (sql)
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgHP.DataSource = rec
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdSelect_Click()
   On Error Resume Next
   cdSelect.CancelError = True
   cdSelect.Filter = "*.xls|*.xls"
   cdSelect.action = 1
   If cdSelect.FileName <> "" Then txtPath.Text = cdSelect.FileName
End Sub

Private Sub cmdUpdate_Click()
   If mfgHP.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫ�޸ĵ���!"
      Exit Sub
   End If
   mfgHP_Click
   enable
   txtHPSNIII.Enabled = False
   txtHPSNIII.BackColor = &HE0E0E0
   
   txtCode.Enabled = False
   txtCode.BackColor = &HE0E0E0
   
   
   op = "Update"
End Sub

Private Sub Form_Load()
   unable
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   renovate
End Sub

Private Sub renovate()
   sql = "select * from hp order by row_number"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set mfgHP.DataSource = rec
   With mfgHP
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 3500
        .ColWidth(3) = 3500
        .ColWidth(4) = 3500
        .ColWidth(5) = 3500
        .ColWidth(6) = 3500
        .ColWidth(7) = 3500
        .ColWidth(8) = 3500
        .ColWidth(9) = 3500
        .ColWidth(10) = 3500
        .ColWidth(11) = 3500
        .ColWidth(12) = 3500
        .ColWidth(13) = 3500

        
        .TextMatrix(0, 1) = "Row Number"
        .TextMatrix(0, 2) = "H3C Bom Code"
        .TextMatrix(0, 3) = "H3C Serial Number-III"
        .TextMatrix(0, 4) = "HP Part Number(text)"
        .TextMatrix(0, 5) = "HP Product(text)"
        .TextMatrix(0, 6) = "HP Serial Number-III"
        .TextMatrix(0, 7) = "HP Product Description(1)"
        .TextMatrix(0, 8) = "HP Product Description(2)"
        .TextMatrix(0, 9) = "HP-GTIN"
        .TextMatrix(0, 10) = "HP-GTIN Number"
        .TextMatrix(0, 11) = "H3C Type"
        .TextMatrix(0, 12) = "Pack Label(Y/N)"
        .TextMatrix(0, 13) = "HP SN Product"
         
   End With
   rec.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rec.State = 1 Then
      rec.Close
      Set rec = Nothing
   End If
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub



Private Sub mfgHP_Click()
    If mfgHP.RowSel > 0 Then
        txtRN.Text = mfgHP.TextMatrix(mfgHP.RowSel, 1)
        txtCode.Text = mfgHP.TextMatrix(mfgHP.RowSel, 2)
        txtH3CSNIII.Text = mfgHP.TextMatrix(mfgHP.RowSel, 3)
        txtHPPN.Text = mfgHP.TextMatrix(mfgHP.RowSel, 4)
        txtHPP.Text = mfgHP.TextMatrix(mfgHP.RowSel, 5)
        txtHPSNIII.Text = mfgHP.TextMatrix(mfgHP.RowSel, 6)
        txtHPdesc1.Text = mfgHP.TextMatrix(mfgHP.RowSel, 7)
        txtHPdesc2.Text = mfgHP.TextMatrix(mfgHP.RowSel, 8)
        
        txtHPGtin.Text = mfgHP.TextMatrix(mfgHP.RowSel, 9)
        txtHPGtinNum.Text = mfgHP.TextMatrix(mfgHP.RowSel, 10)
        txtH3CType.Text = mfgHP.TextMatrix(mfgHP.RowSel, 11)
        txtPack.Text = mfgHP.TextMatrix(mfgHP.RowSel, 12)
        txtHPSNProduct.Text = mfgHP.TextMatrix(mfgHP.RowSel, 13)
        
        If Trim(mfgHP.TextMatrix(mfgHP.RowSel, 14)) = "Y" Then
            cb5000.ListIndex = 0
        ElseIf Trim(mfgHP.TextMatrix(mfgHP.RowSel, 14)) = "N" Then
            cb5000.ListIndex = 1
        ElseIf Trim(mfgHP.TextMatrix(mfgHP.RowSel, 14)) = "NA" Then
            cb5000.ListIndex = 2
        ElseIf Trim(mfgHP.TextMatrix(mfgHP.RowSel, 14)) = "TBD" Then
            cb5000.ListIndex = 3
        End If
        
        If UCase(Trim(mfgHP.TextMatrix(mfgHP.RowSel, 15))) = "YES" Or UCase(Trim(mfgHP.TextMatrix(mfgHP.RowSel, 15))) = "TRUE" Then
            chkYes.Value = 1
        ElseIf UCase(Trim(mfgHP.TextMatrix(mfgHP.RowSel, 15))) = "NO" Or UCase(Trim(mfgHP.TextMatrix(mfgHP.RowSel, 15))) = "FALSE" Then
            chkNo.Value = 1
        End If
    End If
End Sub

Private Sub mfgHP_SelChange()
   mfgHP_Click
End Sub

Private Sub txtHPdesc1_KeyPress(KeyAscii As Integer)
'    Debug.Print KeyAscii
    If KeyAscii = 22 Or KeyAscii = 8 Then
    Else
         KeyAscii = 0
    End If
End Sub

'Private Sub txtHPdesc1_KeyUp(KeyCode As Integer, Shift As Integer)
'     If KeyAscii <> 13 Then
'        If (Len(Me.txtHPdesc1.Text) > 2) Then
'        Else
'            MsgBox "�����ֶ�����Desc"
'            Me.txtHPdesc1.Text = ""
'            Exit Sub
'        End If
'    End If
'End Sub

Private Sub txtHPdesc2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

'Private Sub txtHPdesc2_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyAscii <> 13 Then
'        If (Len(Me.txtHPdesc2.Text) > 2) Then
'        Else
'            MsgBox "�����ֶ�����Desc"
'            Me.txtHPdesc2.Text = ""
'            Exit Sub
'        End If
'    End If
'End Sub
