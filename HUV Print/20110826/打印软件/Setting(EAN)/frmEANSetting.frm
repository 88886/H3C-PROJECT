VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEANSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EAN�� Setting"
   ClientHeight    =   5940
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
   Icon            =   "frmEANSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdSelect 
      Left            =   2400
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   10620
      TabIndex        =   17
      Top             =   5040
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
      Left            =   9180
      TabIndex        =   16
      Top             =   5040
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
      Left            =   7500
      TabIndex        =   15
      Top             =   5040
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
      Left            =   10620
      TabIndex        =   14
      Top             =   4200
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
      Left            =   9180
      TabIndex        =   13
      Top             =   4200
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
      Left            =   7500
      TabIndex        =   12
      Top             =   4200
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
      Left            =   5820
      TabIndex        =   11
      Top             =   4680
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
      Left            =   3300
      TabIndex        =   10
      Top             =   5040
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
      Left            =   3300
      TabIndex        =   9
      Top             =   4200
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
      Left            =   1260
      TabIndex        =   8
      Top             =   5280
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
      Left            =   60
      TabIndex        =   7
      Top             =   4680
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
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.TextBox txtIsPrint 
         Height          =   465
         Left            =   10320
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtSN 
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
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtXH 
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
         Left            =   6120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblisPrint 
         Caption         =   "�Ƿ��ӡ:"
         Height          =   375
         Left            =   8880
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblSN 
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblCPN 
         Caption         =   "��Ʒ�ͺ�:"
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.PictureBox mfgHUAWEI 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   11835
      TabIndex        =   5
      Top             =   1080
      Width           =   11895
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MfgEAN 
         Height          =   2775
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   4895
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
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
      Left            =   60
      TabIndex        =   6
      Top             =   4200
      Width           =   2175
   End
End
Attribute VB_Name = "frmEANSetting"
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
   txtSN.Enabled = True
   txtSN.BackColor = &HFFFFFF
   txtXH.Enabled = True
   txtXH.BackColor = &HFFFFFF
   txtIsPrint.Enabled = True
   txtIsPrint.BackColor = &HFFFFFF
  
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
   txtSN.Enabled = False
   txtSN.BackColor = &HE0E0E0
   txtXH.Enabled = False
   txtXH.BackColor = &HE0E0E0
   txtIsPrint.Enabled = False
   txtIsPrint.BackColor = &HE0E0E0
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


Private Sub cmdCancel_Click()
   unable
   op = ""
End Sub

Private Sub cmdConfirm_Click()
   If txtSN.Text = "" Then
      MsgBox "��Ʒ���벻��Ϊ��!!", vbExclamation + vbOKOnly, "��Ʒ�����"
      txtSN.SetFocus
      Exit Sub
   End If
   If txtXH.Text = "" Then
       MsgBox "EAN Number����Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ�ͺſ�"
       txtXH.SetFocus
       Exit Sub
   End If
   
   If txtIsPrint.Text = "" Then
       MsgBox "�Ƿ���Ϊ��!", vbExclamation + vbOKOnly, "��Ʒ�ͺſ�"
       txtIsPrint.SetFocus
       Exit Sub
   End If
   
   
   If op = "Insert" Then
      Dim rcd As New ADODB.Recordset
      sql = "select Count(*) from ean where SN='" & txtSN.Text & "'"
      rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rcd.Fields(0) > 0 Then
         MsgBox "��Ʒ�����Ѵ���!"
         txtSN.SetFocus
         Exit Sub
      End If
      rcd.Close
      sql = "Insert into ean(SN,type,isprint) " & _
            "Values('" & txtSN.Text & "','" & txtXH.Text & "','" & txtIsPrint.Text & "')"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "��������ʧ��!" & "ԭ����" & status
      End If
      MsgBox "�������ϳɹ�!"
      renovate
      cmdInsert_Click
   ElseIf op = "Update" Then
      sql = "Update EAN set type='" & txtXH.Text & "',isprint='" & txtIsPrint.Text & "'" & _
            " where ID=" & MfgEAN.TextMatrix(MfgEAN.RowSel, 1) & " and SN='" & txtSN.Text & "'"
      status = Connect.excuteUpdate(sql)
      If status <> "" Then
         MsgBox "�޸�����ʧ��!" & "ԭ����" & status
      End If
      MsgBox "�޸����ϳɹ�!"
      renovate
      cmdCancel_Click
   End If
   renovate
End Sub

Private Sub cmdDelete_Click()
   If MfgEAN.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫɾ������!"
      Exit Sub
   End If
   sql = "delete from ean where ID=" & MfgEAN.TextMatrix(MfgEAN.RowSel, 1) & " and SN='" & MfgEAN.TextMatrix(MfgEAN.RowSel, 2) & "'"
   status = Connect.excuteUpdate(sql)
   If status <> "" Then
      MsgBox "ɾ������ʧ��!" & "ԭ����" & status
   End If
   MsgBox "ɾ�����ϳɹ�!"
   renovate
End Sub

Private Sub cmdExport_Click()
   On Error Resume Next
   If MfgEAN.Rows = 0 Then
      MsgBox "�����Ͽɻ��"
      Exit Sub
   End If
   If txtPath.Text <> "" Then
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets.Item(1)
       For i = 0 To MfgEAN.Rows - 1
         For j = 1 To MfgEAN.Cols - 1
          xlSheet.Cells(i + 1, j) = MfgEAN.TextMatrix(i, j)
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
          If xlSheet.Cells(j, 2) = "" Then
             MsgBox "�������ϸ�ʽ����ȷ!"
             Exit Sub
          End If
          isexist = False
          For K = 1 To 2
           If K = 2 Then
             cellValue = xlSheet.Cells(j, K)
             If cellValue = "" Then
                MsgBox "�������ϸ�ʽ����ȷ!"
                Exit Sub
             End If
             Dim rcd As New ADODB.Recordset
             sql = "select Count(*) from ean where SN='" & cellValue & "'"
             rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
             If rcd.Fields(0) > 0 Then
                If action = 0 Then
                   action = MsgBox("��Ʒ�����Ѵ���!", vbAbortRetryIgnore + vbExclamation, "�����ظ�")
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

            If K = 2 Then
               If action = vbRetry Then
                   sql = "Update ean set type='" & xlSheet.Cells(j, 3) & "' isprint='" & xlSheet.Cells(j, 4) & "'" & _
                        " where SN='" & xlSheet.Cells(j, 2) & "'"
                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                     MsgBox "�޸�����ʧ��!" & "ԭ����" & status
                   End If
'                   MsgBox "�޸�EAN�趨���ϳɹ�!"
               ElseIf isexist = False Then
                   sql = "Insert into ean(SN,type,isprint) " & _
                        "Values('" & xlSheet.Cells(j, 2) & "','" & xlSheet.Cells(j, 3) & "','" & xlSheet.Cells(j, 4) & "')"

                   status = Connect.excuteUpdate(sql)
                   If status <> "" Then
                      MsgBox "��������ʧ��!" & "ԭ����" & status
                   End If
'                   MsgBox "����EAN�趨���ϳɹ�!"
               End If
           End If
         Next K
        End If
       Next j
      Next i
      xlBook.Close
      Set xlBook = Nothing
      xlApp.Quit
      MsgBox "���ϵ���ɹ�!"
      renovate
End Sub

Private Sub cmdInsert_Click()
   enable
   txtSN.Text = ""
   txtXH.Text = ""
   txtIsPrint.Text = ""
   
   op = "Insert"
End Sub

Private Sub cmdQuery_Click()
   MsgBox "�밴������ť��վͿ������ѯ����!", vbOKOnly + vbInformation, "�����ѯ����"
   If rec.State = 1 Then
      rec.Close
   End If
   sql = "select * from ean Where 1=1"
   If txtSN.Text <> "" Then
      sql = sql & " and SN like '%" & txtSN.Text & "%'"
   End If
   If txtXH.Text <> "" Then
      sql = sql & " and type like '%" & txtXH.Text & "%'"
   End If
     sql = sql & " order by ID,SN"
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set MfgEAN.DataSource = rec
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
   If MfgEAN.RowSel <= 0 Then
      MsgBox "��ѡ��Ҫ�޸ĵ���!"
      Exit Sub
   End If
   mfgEAN_Click
   enable
   txtSN.Enabled = False
   txtSN.BackColor = &HE0E0E0
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
   sql = "select * from ean order by ID,SN"
   If rec.State = 1 Then
      rec.Close
   End If
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   Set MfgEAN.DataSource = rec
   With MfgEAN
        .Cols = rec.Fields.Count + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 3500
        .ColWidth(3) = 3500
        .ColWidth(4) = 3500
        .TextMatrix(0, 1) = "���(ID)"
        .TextMatrix(0, 2) = "��Ʒ����(Model Number)"
        .TextMatrix(0, 3) = "EAN Number"
        .TextMatrix(0, 4) = "�Ƿ��ӡ3COM��ַ"
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

Private Sub mfgEAN_Click()
   If MfgEAN.RowSel > 0 Then
      txtSN.Text = MfgEAN.TextMatrix(MfgEAN.RowSel, 2)
      txtXH.Text = MfgEAN.TextMatrix(MfgEAN.RowSel, 3)
      txtIsPrint.Text = MfgEAN.TextMatrix(MfgEAN.RowSel, 4)
    
   End If
End Sub

Private Sub mfgEAN_SelChange()
   mfgEAN_Click
End Sub


