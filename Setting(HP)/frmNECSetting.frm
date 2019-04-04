VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmNewNECSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New NEC Setting"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   17325
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
   MaxButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   17325
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgH3C 
      Height          =   3495
      Left            =   120
      TabIndex        =   42
      Top             =   4200
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   6165
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fmH3C 
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   16935
      Begin VB.TextBox txtRemark 
         Height          =   495
         Left            =   5040
         TabIndex        =   49
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtRCM 
         Height          =   495
         Left            =   4080
         TabIndex        =   47
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CheckBox chkNonRCM 
         BackColor       =   &H0000C000&
         Caption         =   "无"
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
         Left            =   1800
         TabIndex        =   46
         Top             =   3240
         Width           =   735
      End
      Begin VB.CheckBox chkRCM 
         Caption         =   "有"
         Height          =   375
         Left            =   1080
         TabIndex        =   45
         Top             =   3240
         Width           =   735
      End
      Begin VB.PictureBox Picture10 
         Height          =   495
         Left            =   120
         Picture         =   "frmNECSetting.frx":0000
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   43
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox chkNonCE 
         BackColor       =   &H0000C000&
         Caption         =   "无 CE"
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
         Left            =   1560
         TabIndex        =   31
         Top             =   2040
         Width           =   1215
      End
      Begin VB.PictureBox Picture4 
         Height          =   735
         Left            =   9600
         Picture         =   "frmNECSetting.frx":0BA6
         ScaleHeight     =   675
         ScaleWidth      =   3555
         TabIndex        =   30
         Top             =   2160
         Width           =   3615
      End
      Begin VB.PictureBox Picture3 
         Height          =   615
         Left            =   6360
         Picture         =   "frmNECSetting.frx":8834
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   29
         Top             =   2160
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         Height          =   615
         Left            =   3480
         Picture         =   "frmNECSetting.frx":A1C6
         ScaleHeight     =   555
         ScaleWidth      =   675
         TabIndex        =   28
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox chkNonChinaRoHS 
         BackColor       =   &H0000C000&
         Caption         =   "无"
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
         Left            =   5280
         TabIndex        =   27
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox chkNonTurkeyRohs 
         BackColor       =   &H0000C000&
         Caption         =   "无"
         Height          =   375
         Left            =   14520
         TabIndex        =   26
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox chkNonWEEE 
         BackColor       =   &H0000C000&
         Caption         =   "无"
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
         Left            =   8160
         TabIndex        =   25
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   120
         Picture         =   "frmNECSetting.frx":B72C
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   24
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtProductDescribe1 
         Height          =   495
         Left            =   8880
         TabIndex        =   23
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtSN 
         Height          =   450
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtProductDescribe2 
         Height          =   450
         Left            =   13320
         TabIndex        =   21
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtProductDescribe3 
         Height          =   450
         Left            =   1560
         TabIndex        =   20
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtEAN 
         Height          =   450
         Left            =   5040
         TabIndex        =   19
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtGW 
         Height          =   450
         Left            =   1560
         TabIndex        =   18
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkCE 
         Caption         =   "CE"
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   2280
         Width           =   855
      End
      Begin VB.CheckBox chkWEEE 
         Caption         =   "有"
         Height          =   375
         Left            =   7320
         TabIndex        =   16
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtHV 
         Height          =   495
         Left            =   5040
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txXXXXXX 
         Height          =   495
         Left            =   25200
         TabIndex        =   14
         Text            =   "sdgfdsfadsfadsf"
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox chkChinaRoHS 
         Caption         =   "有"
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Top             =   2280
         Width           =   615
      End
      Begin VB.CheckBox chkTurkeyRohs 
         Caption         =   "有"
         Height          =   375
         Left            =   13560
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox chkSVPrint 
         Caption         =   "是"
         Height          =   495
         Left            =   13920
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox chkNonSVPrint 
         BackColor       =   &H0000C000&
         Caption         =   "否"
         Height          =   495
         Left            =   14880
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox chkCEAddr 
         Caption         =   "NEC Addr"
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
         Left            =   1560
         TabIndex        =   9
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtG15Revision 
         Height          =   450
         Left            =   8880
         TabIndex        =   8
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label16 
         Caption         =   "备注:"
         Height          =   495
         Left            =   3720
         TabIndex        =   48
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "RCM ID"
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
         Left            =   3000
         TabIndex        =   44
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "项目描述1:"
         Height          =   495
         Left            =   7200
         TabIndex        =   41
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSN 
         Caption         =   "产品编码:"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblCPN 
         Caption         =   "项目描述2:"
         Height          =   375
         Left            =   11880
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblEPN 
         Caption         =   "项目描述3:"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblGW 
         Caption         =   "毛重(kg):"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblHV 
         Caption         =   "硬件版本:"
         Height          =   495
         Left            =   3720
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblRemark 
         Caption         =   "备注:"
         Height          =   495
         Left            =   24120
         TabIndex        =   35
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblOS 
         Caption         =   "EAN码:"
         Height          =   375
         Left            =   3720
         TabIndex        =   34
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblPrintSV 
         Caption         =   "打印软件版本:"
         Height          =   495
         Left            =   11880
         TabIndex        =   33
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "G15版本:"
         Height          =   375
         Left            =   7200
         TabIndex        =   32
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(Return)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14760
      TabIndex        =   6
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13200
      TabIndex        =   5
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "确定(Confirm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11520
      TabIndex        =   4
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(Delete)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14760
      TabIndex        =   3
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修改(Update)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13200
      TabIndex        =   2
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "新增(Insert)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11520
      TabIndex        =   1
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查询(Query)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9840
      TabIndex        =   0
      Top             =   8400
      Width           =   1455
   End
End
Attribute VB_Name = "frmNewNECSetting"
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

Private Sub Reset()
    Dim ctr As Control

    If op = "Insert" Then
        cmdQuery.Enabled = True
        cmdInsert.Enabled = False
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
        cmdConfirm.Enabled = True
        cmdCancel.Enabled = True
        For Each ctr In Me.Controls
            If TypeOf ctr Is TextBox Then
                    ctr.Text = ""
                    ctr.Enabled = True
                    ctr.BackColor = &HFFFFFF
                 ElseIf TypeOf ctr Is ComboBox Then
                    If ctr.Style = 2 Then
                       ctr.ListIndex = -1
                    Else
                       ctr.Text = ""
                    End If
                 ElseIf TypeOf ctr Is CheckBox Then
                    ctr.Value = 0
            End If
        Next
        Me.txtRemark = "N/A"
        Me.chkChinaRoHS.Value = 1
        Me.chkTurkeyRohs.Value = 1
        Me.chkWEEE.Value = 1
        Me.txtSN.Enabled = True
        Me.txtSN.SetFocus
    ElseIf op = "Cancel" Then
        For Each ctr In Me.Controls
        If TypeOf ctr Is TextBox Then
                ctr.Enabled = False
                ctr.BackColor = &HFFFFFF
             ElseIf TypeOf ctr Is ComboBox Then
                ctr.Enabled = False
                ctr.BackColor = &HFFFFFF
             ElseIf TypeOf ctr Is CheckBox Then
                ctr.Enabled = True
        End If
    Next
        cmdQuery.Enabled = True
        cmdInsert.Enabled = True
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
        cmdConfirm.Enabled = True
        cmdCancel.Enabled = False
    ElseIf op = "Update" Then
        For Each ctr In Me.Controls
        If TypeOf ctr Is TextBox Then
                ctr.Enabled = True
             ElseIf TypeOf ctr Is ComboBox Then
                ctr.Enabled = True
                End If
    Next
        txtSN.Enabled = False
        txtSN.BackColor = &HE0E0E0
    End If
End Sub
Private Sub enable()
   txtSN.Enabled = True
   txtSN.BackColor = &HFFFFFF
   txtCPN.Enabled = True
   txtCPN.BackColor = &HFFFFFF
   txtEPN.Enabled = True
   txtEPN.BackColor = &HFFFFFF
   txtDes.Enabled = True
   txtDes.BackColor = &HFFFFFF
   txtSize.Enabled = True
   txtSize.BackColor = &HFFFFFF
   txtGW.Enabled = True
   txtGW.BackColor = &HFFFFFF
   
   chkCE.Enabled = True
   chkCEAddr.Enabled = True
   chkHPEAddr.Enabled = True
      
   chkNonCE.Enabled = True
   chkWEEE.Enabled = True
   chkNonWEEE.Enabled = True
   chkEAC.Enabled = True
   chkNonEAC.Enabled = True
   chkChinaRoHS.Enabled = True
   chkNonChinaRoHS.Enabled = True
   chkTurkeyRohs.Enabled = True
   chkNonTurkeyRohs.Enabled = True
   
   chkRoHS.Enabled = True
   chkNonRoHS.Enabled = True
   'optH3CRoHS.Enabled = True
   'opt3COMRoHS.Enabled = True
   'optNonRoHS.Enabled = True
   
   
   txtMS.Enabled = True
   txtMS.BackColor = &HFFFFFF

   
   chkSVPrint.Enabled = True
   chkNonSVPrint.Enabled = True
   
   txtHV.Enabled = True
   txtHV.BackColor = &HFFFFFF
   txtRemark.Enabled = True
   txtRemark.BackColor = &HFFFFFF
   
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
   
   

End Sub




Private Sub chkNonCE_Click()
   If chkNonCE.Value = 1 Then
      chkCE.Value = 0
      chkCEAddr.Value = 0
   Else
      chkCE.Value = 1
   End If
End Sub

Private Sub chkChinaROHS_Click()
   If chkChinaRoHS.Value = 1 Then
      chkNonChinaRoHS.Value = 0
   End If
End Sub

Private Sub chkNonChinaRoHS_Click()
   If chkNonChinaRoHS.Value = 1 Then
      chkChinaRoHS.Value = 0
   End If
End Sub

Private Sub chkCEAddr_Click()
If chkCEAddr.Value = 1 Then
 chkNonCE.Value = 0
 chkCE.Value = 1
 End If
 
End Sub

Private Sub chkNonRCM_Click()
    If Me.chkNonRCM.Value = 1 Then
        Me.chkRCM.Value = 0
        Me.txtRCM.Text = ""
        Me.txtRCM.Enabled = False
    End If
End Sub

Private Sub chkNonSVPrint_Click()
    If Me.chkNonSVPrint.Value = 1 Then
        Me.chkSVPrint.Value = 0
    End If
End Sub

Private Sub chkRCM_Click()
    If Me.chkRCM.Value = 1 Then
        Me.chkNonRCM.Value = 0
        Me.txtRCM.Enabled = True
       ' Me.txtRCM.Text = "N279"
    End If
End Sub


Private Sub chkSVPrint_Click()
   If chkSVPrint.Value = 1 Then
      chkNonSVPrint.Value = 0
   End If
End Sub


Private Sub chkWEEE_Click()
   If chkWEEE.Value = 1 Then
      chkNonWEEE.Value = 0
   Else
      chkNonWEEE.Value = 1
   End If
End Sub

Private Sub chkNonWEEE_Click()
   If chkNonWEEE.Value = 1 Then
      chkWEEE.Value = 0
   Else
      chkWEEE.Value = 1
   End If
End Sub


Private Sub chkTurkeyROHS_Click()
    If chkTurkeyRohs.Value = 1 Then
        chkNonTurkeyRohs.Value = 0
    Else
        chkNonTurkeyRohs.Value = 1
    End If
End Sub

Private Sub chkNonTurkeyRohs_Click()
    If chkNonTurkeyRohs.Value = 1 Then
        chkTurkeyRohs.Value = 0
    Else
        chkTurkeyRohs.Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
   op = "Cancel"
   Reset
End Sub

Private Sub cmdConfirm_Click()
   If Trim(txtSN.Text) = "" Then
      MsgBox "产品编码不能为空!!", vbExclamation + vbOKOnly, "产品编码空"
      txtSN.SetFocus
      Exit Sub
   End If

 
   If Trim(txtGW.Text) <> "" Then
        If UCase(Right(Trim(txtGW.Text), 2)) <> "KG" Then
           MsgBox "毛重必须加上单位kg!", vbExclamation + vbOKOnly, "毛重单位空"
           txtGW.SetFocus
           Exit Sub
        End If
        If Len(Me.txtGW.Text) < 6 Or Mid(Right(Trim(Me.txtGW.Text), 5), 1, 1) <> "." Then
            MsgBox "毛重数据长度应大于6位并且包含小数点，如x.xxkg!", vbExclamation + vbOKOnly, "毛重格式不正确"
            txtGW.SetFocus
            Exit Sub
        End If
   End If
   

   If txtHV.Text = "" Then
      MsgBox "硬件版本不能为空!", vbExclamation + vbOKOnly, "硬件版本空"
      txtHV.SetFocus
      Exit Sub
   End If
   If chkSVPrint.Value = 0 And chkNonSVPrint.Value = 0 Then
      MsgBox "是否打印软件版本不能为空!", vbExclamation + vbOKOnly, "软件件版本空"
      txtHV.SetFocus
      Exit Sub
   End If
   

   
   
   Dim CE, WEEE, ChinaRoHS, TurkeyRoHS, SVPrint, RCM, RCM_ID As String
  

   
   If chkCE.Value = 1 Then
'      CE = "CE"
      If chkCEAddr.Value = 1 Then
        CE = "2"
      Else
        CE = "1"
      End If
   ElseIf chkNonCE.Value = 1 Then
      CE = "0"
   End If
   If chkWEEE.Value = 1 Then
      WEEE = "1"
   ElseIf chkNonWEEE.Value = 1 Then
      WEEE = "0"
   End If

   If chkChinaRoHS.Value = 1 Then
      ChinaRoHS = "1"
   Else
      ChinaRoHS = "0"
   End If
   If chkTurkeyRohs.Value = 1 Then
      TurkeyRoHS = "1"
   Else
      TurkeyRoHS = "0"
   End If
   If chkSVPrint.Value = 1 Then
      SVPrint = "1"
   Else
      SVPrint = "0"
   End If


    If (Me.chkRCM.Value = 1 And Len(Trim(Me.txtRCM.Text)) = 0) Then
         MsgBox "RCM 文本框没有值，请确认!", vbExclamation + vbOKOnly, "RCM没有值"
         Exit Sub
    End If
        
    If Me.chkRCM.Value = 1 Then
        RCM = "1"
        If Trim(Me.txtRCM.Text) = "/" Or UCase(Trim(Me.txtRCM.Text)) = "N/A" Or UCase(Trim(Me.txtRCM.Text)) = "NA" Then
            RCM_ID = ""
        Else
            RCM_ID = Trim(Me.txtRCM.Text)
        End If
    Else
        RCM = "0"
        RCM_ID = ""
    End If
    
  
  txtGW.Text = LCase(Trim(txtGW.Text))
   
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from tblNEC where PartNum ='" & Trim(txtSN.Text) & "' and HardWare_Revision ='" & Trim(txtHV.Text) & "' "
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "产品编码&版本已存在!", vbExclamation + vbOKOnly, "产品编号重复"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close

      sql = "insert [tblNEC]([PartNum],[HardWare_Revision],[Product_Describe1],[Product_Describe2],[Product_Describe3],[EAN],[G15_Revision],[GW],[Print_SV],[CE],[ChinaRoHS],[WEEE],[TurkeyRoHS],[RCM],[RCM_ID],[Remark]) " & _
            "Values('" & Trim(txtSN.Text) & "','" & Trim(txtHV.Text) & "','" & Trim(txtProductDescribe1.Text) & "','" & Trim(txtProductDescribe2.Text) & "','" & Trim(txtProductDescribe3.Text) & "','" & Trim(txtEAN.Text) & "','" & Trim(txtG15Revision.Text) & "','" & Trim(txtGW.Text) & "'," & SVPrint & _
            ", " & CE & " , " & ChinaRoHS & "  ," & WEEE & " ," & TurkeyRoHS & " , " & RCM & " ,'" & RCM_ID & "','" & Trim(txtRemark.Text) & "')"
    Debug.Print (sql)

     
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "新增NEC设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "新增失败"
      Else
        MsgBox "新增NEC设定资料成功!", vbInformation + vbOKOnly, "新增成功"
      End If
      renovate ("")
      cmdInsert_Click
   ElseIf op = "Update" Then
      
      
   sql = "Update tblNEC set Product_Describe1='" & Trim(txtProductDescribe1.Text) & "',Product_Describe2='" & Trim(txtProductDescribe2.Text) & "',Product_Describe3 ='" & Trim(txtProductDescribe3.Text) & "',EAN='" & Trim(txtEAN.Text) & "',G15_Revision='" & Trim(txtG15Revision.Text) & "',GW='" & Trim(txtGW.Text) & "',Print_SV=" & SVPrint & ",CE=" & CE & ",WEEE=" & WEEE & ",ChinaRoHS=" & ChinaRoHS & _
            ",TurkeyRoHS=" & TurkeyRoHS & ",RCM_ID = '" & RCM_ID & "',RCM = " & RCM & _
            ",Remark='" & txtRemark.Text & _
           "' where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and PartNum ='" & Trim(txtSN.Text) & "'"
        
    Debug.Print (sql)

    
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "修改NEC设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "修改失败"
      Else
         MsgBox "修改NEC设定资料成功!", vbInformation + vbOKOnly, "修改成功"
      End If
      renovate ("")
      cmdCancel_Click
   End If
   renovate ("")
End Sub

Private Sub cmdDelete_Click()
   If mfgH3C.RowSel <= 0 Then
      MsgBox "请选择要删除的行!", vbInformation + vbOKOnly, "未选择行"
      Exit Sub
   End If
   

   sql = sql & " delete from tblNEC where ID=" & mfgH3C.TextMatrix(mfgH3C.RowSel, 1) & " and PartNum ='" & mfgH3C.TextMatrix(mfgH3C.RowSel, 2) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "删除NEC设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "删除失败"
   End If
   MsgBox "删除NEC设定资料成功!", vbInformation + vbOKOnly, "删除成功"
   renovate ("")
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If mfgH3C.Rows = 0 Then
      MsgBox "无资料可汇出", vbExclamation + vbOKOnly, "无资料"
      Exit Sub
   End If
   If txtPath.Text <> "" Then
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets.Item(1)
       For i = 0 To mfgH3C.Rows - 1
         For j = 1 To mfgH3C.Cols - 1
          xlSheet.Cells(i + 1, j) = mfgH3C.TextMatrix(i, j)
       Next j
      Next i
      xlBook.SaveAs (txtPath.Text)
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "汇出到EXCEL资料成功!!", vbInformation + vbOKOnly, "汇出成功"
    End If
End Sub

Private Sub cmdImport_Click()
   If txtPath.Text = "" Then
      MsgBox "导入路径不能为空!", vbExclamation + vbOKOnly, "导入路径空"
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
          Dim cellhvValue As String
          
          Dim isexist As Boolean
          If xlSheet.Cells(j, 19) = "" Then
             MsgBox "导入资料格式不正确!", vbExclamation + vbOKOnly, "格式错误"
             Exit Sub
          End If
          If Not ((xlSheet.Cells(j, 18) = "N") Or (xlSheet.Cells(j, 18) = "Y")) Then
             MsgBox "导入资料格式不正确!", vbExclamation + vbOKOnly, "格式错误"
             Exit Sub
          End If
          isexist = False
          For K = 1 To 19
          '======================================================
           If K = 3 Then
             cellValue = xlSheet.Cells(j, K)
             cellhvValue = xlSheet.Cells(j, 2)
             
             If cellValue = "" Or cellhvValue = "" Then
                MsgBox "导入资料格式不正确!", vbExclamation + vbOKOnly, "格式错误"
                Exit Sub
             End If
             
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from tblH3C where SN='" & cellValue & "' and HV='" & cellhvValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("产品编码&版本已存在!", vbAbortRetryIgnore + vbExclamation, "资料重复")
                End If
                
                If action = vbAbort Then
                   MsgBox "资料导入已终止!!", vbInformation + vbOKOnly, "导入终止"
                   rcd.Close
                   Exit Sub
                ElseIf action = vbIgnore And info = True Then
                   MsgBox "重复产品编号资料不会导入,请稍等..!!", vbInformation + vbOKOnly, "重复不会导入"
                   rcd.Close
                   info = False
                   Exit For
                ElseIf action = vbRetry And info = True Then
                   MsgBox "重复产品编号资料会自动更新,请稍等..!!", vbInformation + vbOKOnly, "重复会自动更新"
                   info = False
                End If
                isexist = True
             Else
                isexist = False
             End If
             rcd.Close
            End If
            '==================================================
            
            If K = 19 Then
               If action = vbRetry Then
                   sql = "Update tblH3C set CPN='" & xlSheet.Cells(j, 4) & "',EPN='" & xlSheet.Cells(j, 5) & "',Des='" & xlSheet.Cells(j, 6) & "',OS='" & xlSheet.Cells(j, 7) & "',GW='" & xlSheet.Cells(j, 8) & "',CE='" & xlSheet.Cells(j, 9) & "',WEEE='" & xlSheet.Cells(j, 10) & "',ChinaRoHS='" & xlSheet.Cells(j, 11) & "'," & _
                        "RoHS='" & xlSheet.Cells(j, 12) & "',TurkeyRohs='" & xlSheet.Cells(j, 13) & "',MS='" & xlSheet.Cells(j, 14) & "',MSValidFrom='" & xlSheet.Cells(j, 15) & "',NAL='" & xlSheet.Cells(j, 16) & "',ValidFrom='" & xlSheet.Cells(j, 17) & "',PrintSV='" & xlSheet.Cells(j, 18) & "',Remark='" & xlSheet.Cells(j, 19) & "'" & _
                        " where SN='" & xlSheet.Cells(j, 3) & "' and HV='" & xlSheet.Cells(j, 2) & "' "
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "修改H3C设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "修改失败"
                   End If
'                   MsgBox "修改H3C设定资料成功!"
               ElseIf isexist = False Then
                   sql = "Insert into tblH3C(ID, HV, SN, CPN, EPN, Des, OS, GW, CE, WEEE, ChinaRoHS, RoHS, TurkeyRohs, MS, MSValidFrom, NAL, ValidFrom, PrintSV, Remark) " & _
                        " Values(" & getmaxID("tblH3C") & ",'" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "','" & xlSheet.Cells(j, 4) & "','" & xlSheet.Cells(j, 5) & "','" & xlSheet.Cells(j, 6) & "','" & xlSheet.Cells(j, 7) & "','" & xlSheet.Cells(j, 8) & "','" & xlSheet.Cells(j, 9) & "','" & xlSheet.Cells(j, 10) & "','" & xlSheet.Cells(j, 11) & "'," & _
                        "'" & xlSheet.Cells(j, 12) & "','" & xlSheet.Cells(j, 13) & "','" & xlSheet.Cells(j, 14) & "','" & xlSheet.Cells(j, 15) & "','" & xlSheet.Cells(j, 16) & "','" & xlSheet.Cells(j, 17) & "','" & xlSheet.Cells(j, 18) & "','" & xlSheet.Cells(j, 19) & "')"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "新增H3C设定资料失败!" & "原因是" & status, vbExclamation + vbOKOnly, "修改失败"
                   End If
'                   MsgBox "新增H3C设定资料成功!"
               End If
           End If
         Next K
         
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "H3C设定资料导入成功!"
      renovate ("")
End Sub

Private Sub cmdInsert_Click()
    op = "Insert"
    Reset
End Sub

Private Sub cmdQuery_Click()
    If txtSN.Enabled = False Then
      MsgBox "请按新增按钮清空就可输入查询内容!", vbOKOnly + vbInformation, "输入查询内容"
    End If
    If rec.State = 1 Then
        rec.Close
     End If
       sql = "SELECT [ID],[PartNum],[HardWare_Revision],[Product_Describe1],[Product_Describe2],[Product_Describe3],[EAN],[G15_Revision],[GW] " & _
        ",case when Print_SV is null then 'N/A' when Print_SV = 0 then 'No' when Print_SV = 1 then 'Yes' end as 'Print_SV'" & _
        ",case when [CE] = 0 then 'Non CE' when CE = 1 then 'CE' when CE = 2 then 'NEC Addr' end as 'CE'" & _
        ",case when ChinaRoHS is null then 'N/A' when ChinaRoHS = 0 then 'No' when ChinaRoHS = 1 then 'Yes' end as 'ChinaRoHS'" & _
        ",case when WEEE is null then 'N/A' when WEEE = 0 then 'No' when WEEE = 1 then 'Yes' end as 'WEEE'" & _
        ",case when [TurkeyRoHS] is null then 'N/A' when [TurkeyRoHS] = 0 then 'No' when TurkeyRoHS = 1 then 'Yes' end as '[TurkeyRoHS]'" & _
        ",case when RCM is null then 'N/A' when RCM = 0 then 'No' when RCM = 1 then 'Yes' end as 'RCM'" & _
        ",[RCM_ID],[Remark]" & _
        " FROM [Print].[dbo].[tblNEC] where 1 = 1"
     
     If txtSN.Text <> "" Then
        sql = sql & " and PartNum like '%" & txtSN.Text & "%'"
     End If
    ' If txtCPN.Text <> "" Then
      '  sql = sql & " and CPN like '%" & txtCPN.Text & "%'"
    ' End If
    ' If txtEPN.Text <> "" Then
    '    sql = sql & " and EPN='%" & txtEPN.Text & "%'"
    ' End If
     'If Me.txtProductID.Text <> "" Then
    '    sql = sql & " and ProductID like '%" & Me.txtProductID.Text & "%'"
    ' End If
'     If txtSize.Text <> "" Then
'        sql = sql & " and Size like '%" & txtSize.Text & "%'"
'     End If
'     If txtGW.Text <> "" Then
'        sql = sql & " and GW like '%" & txtGW.Text & "%'"
'     End If
     sql = sql & " order by PartNum,HardWare_Revision"
    renovate (sql)
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub


Private Sub cmdUpdate_Click()
   If mfgH3C.RowSel <= 0 Then
      MsgBox "请选择要修改的行!", vbInformation + vbOKOnly, "未选择行"
      Exit Sub
   End If
   mfgH3C_Click
   op = "Update"
   Reset
End Sub

Private Sub Form_Load()
   unable
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   

   renovate ("")
End Sub

Private Sub renovate(sql As String)
    Set mfgH3C.DataSource = Nothing
    If sql = "" Then
        sql = "SELECT [ID],[PartNum],[HardWare_Revision],[Product_Describe1],[Product_Describe2],[Product_Describe3],[EAN],[G15_Revision],[GW] " & _
        ",case when Print_SV is null then 'N/A' when Print_SV = 0 then 'No' when Print_SV = 1 then 'Yes' end as 'Print_SV'" & _
        ",case when [CE] = 0 then 'Non CE' when CE = 1 then 'CE' when CE = 2 then 'NEC Addr' end as 'CE'" & _
        ",case when ChinaRoHS is null then 'N/A' when ChinaRoHS = 0 then 'No' when ChinaRoHS = 1 then 'Yes' end as 'ChinaRoHS'" & _
        ",case when WEEE is null then 'N/A' when WEEE = 0 then 'No' when WEEE = 1 then 'Yes' end as 'WEEE'" & _
        ",case when [TurkeyRoHS] is null then 'N/A' when [TurkeyRoHS] = 0 then 'No' when TurkeyRoHS = 1 then 'Yes' end as '[TurkeyRoHS]'" & _
        ",case when RCM is null then 'N/A' when RCM = 0 then 'No' when RCM = 1 then 'Yes' end as 'RCM'" & _
        ",[RCM_ID],[Remark]" & _
        " FROM [Print].[dbo].[tblNEC] order by PartNum,HardWare_Revision"
    End If
    If rec.State = 1 Then
    rec.Close
    End If
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    Set mfgH3C.DataSource = rec
    With mfgH3C
      .Cols = rec.Fields.Count + 1
      .ColWidth(0) = 400
      .ColWidth(1) = 1000
      .ColWidth(2) = 1600
      .ColWidth(3) = 1300
      .ColWidth(4) = 1500
      .ColWidth(5) = 1500
      .ColWidth(6) = 1500
      .ColWidth(7) = 1000
      .ColWidth(8) = 1500
      .ColWidth(9) = 1000
      .ColWidth(10) = 1500
      .ColWidth(11) = 1000
      .ColWidth(12) = 1600
      .ColWidth(13) = 1000
      .ColWidth(14) = 1800
      .ColWidth(15) = 1000
      .ColWidth(16) = 1300
      .ColWidth(17) = 1000


      .TextMatrix(0, 1) = "ID"
      .TextMatrix(0, 2) = "产品编码"
      .TextMatrix(0, 3) = "硬件版本"
      .TextMatrix(0, 4) = "项目描述1"
      .TextMatrix(0, 5) = "项目描述2"
      .TextMatrix(0, 6) = "项目描述3"
      .TextMatrix(0, 7) = "EAN码"
      .TextMatrix(0, 8) = "G15版本"
      .TextMatrix(0, 9) = "毛重"
      .TextMatrix(0, 10) = "打印版本"
      .TextMatrix(0, 11) = "CE"
      .TextMatrix(0, 12) = "ChinaRoHS"
      .TextMatrix(0, 13) = "WEEE"
      .TextMatrix(0, 14) = "TurkeyRoHS"
      .TextMatrix(0, 15) = "RCM"
      .TextMatrix(0, 16) = "RCM ID"
      .TextMatrix(0, 17) = "备注"
  
    End With
    rec.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rec.State = 1 Then
      rec.Close
      Set conn = Nothing
   End If
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub

Private Sub Label5_Click()

End Sub

Private Sub mfgH3C_Click()

   If mfgH3C.RowSel > 0 Then
      txtHV.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 3)
      txtSN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 2)
      txtProductDescribe1.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 4)
      txtProductDescribe2.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 5)
      txtProductDescribe3.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 6)
      txtEAN.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 7)
      txtGW.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 9)
      txtG15Revision.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 8)
      txtRemark.Text = mfgH3C.TextMatrix(mfgH3C.RowSel, 17)
      
    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 11))) = "CE" Then
     chkCE.Value = 1
     chkNonCE.Value = 0
     chkCEAddr.Value = 0
    ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 11) = "Non CE" Then
     chkCE.Value = 0
     chkNonCE.Value = 1
     chkCEAddr.Value = 0
    ElseIf mfgH3C.TextMatrix(mfgH3C.RowSel, 11) = "NEC Addr" Then
     chkCE.Value = 1
     chkNonCE.Value = 0
     chkCEAddr.Value = 1
    End If
    
    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 13))) = "YES" Then
     chkWEEE.Value = 1
    Else
     chkNonWEEE.Value = 1
    End If
    
    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 12))) = "YES" Then
     chkChinaRoHS.Value = 1
    ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 12)) = "NO" Then
     chkNonChinaRoHS.Value = 1
    End If
    
    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 14))) = "YES" Then
     chkTurkeyRohs.Value = 1
    ElseIf UCase(mfgH3C.TextMatrix(mfgH3C.RowSel, 14)) = "NO" Then
     chkNonTurkeyRohs.Value = 1
    End If
    
    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 10))) = "YES" Then
        Me.chkSVPrint.Value = 1
    ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 10))) = "NO" Then
        Me.chkNonSVPrint.Value = 1
    End If
    
    If UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 15))) = "YES" Then
        Me.chkRCM.Value = 1
        Me.txtRCM.Text = Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 16))
    ElseIf UCase(Trim(mfgH3C.TextMatrix(mfgH3C.RowSel, 15))) = "NO" Then
        Me.chkNonRCM.Value = 1
        Me.txtRCM.Text = ""
    End If
    
   End If
End Sub

Private Sub mfgH3C_SelChange()
   mfgH3C_Click
End Sub


Private Sub MSHFlexGrid1_Click()

End Sub

