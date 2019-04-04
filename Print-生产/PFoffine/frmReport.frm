VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   Caption         =   "��ӡ��¼��ѯ����"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9405
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdPathSet 
      Caption         =   "�趨����·��"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog cdSelect 
      Left            =   600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dsMain 
      Height          =   3495
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "����Excel"
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "��ѯ"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox tbOrder 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "����·��:"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "�����빤��:"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xlApp As New Excel.Application
Dim xlBook As New Excel.Workbook
Dim xlSheet As New Excel.Worksheet
Private Sub cmdExport_Click()
    On Error Resume Next
       If Me.txtPath.Text = "" Then
          MsgBox "δ�趨����·�������趨", vbExclamation + vbOKOnly, "û�е���·��"
          Exit Sub
       End If
       
       If dsMain.Rows = 0 Then
          MsgBox "�����Ͽɻ��", vbExclamation + vbOKOnly, "������"
          Exit Sub
       End If
       If txtPath.Text <> "" Then
          Set xlBook = xlApp.Workbooks.Add
          Set xlSheet = xlBook.Sheets.Item(1)
           For i = 0 To dsMain.Rows - 1
             For j = 1 To dsMain.Cols - 1
              xlSheet.Cells(i + 1, j) = dsMain.TextMatrix(i, j)
           Next j
          Next i
          xlBook.SaveAs (txtPath.Text)
          xlBook.Close
          Set xlBook = Nothing
          xlApp.Quit
          MsgBox "�����EXCEL���ϳɹ�!!", vbInformation + vbOKOnly, "����ɹ�"
        End If
End Sub

Private Sub cmdPathSet_Click()
    On Error Resume Next
   cdSelect.CancelError = True
   cdSelect.Filter = "*.xls|*.xls"
   cdSelect.Action = 1
   If cdSelect.FileName <> "" Then
        txtPath.Text = cdSelect.FileName
   End If
   
End Sub

Private Sub cmdQuery_Click()
    Dim order As String
    Dim sql As String
    Dim conn As New ADODB.Connection
    Dim rec As New ADODB.Recordset
    Me.cmdQuery.Enabled = False
    order = Trim(Me.tbOrder.Text)
    If order = "" Then
        MsgBox "��������Ϊ��!", vbOKOnly + vbExclamation, "��������Ϊ�գ�"
        Exit Sub
    End If
    
   
    conn.ConnectionString = getConnectionstring()
    sql = "select * from PacketFrontRecords where wo = '" & Me.tbOrder.Text & "'"
    conn.Open
'    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    rec.CursorLocation = adUseClient
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    
    
   If rec.RecordCount > 0 Then
        Set Me.dsMain.DataSource = rec
          With dsMain
            .Cols = rec.Fields.count + 1
             .ColWidth(0) = 400
             .ColWidth(1) = 1000
             .ColWidth(2) = 2500
             .ColWidth(3) = 1000
             
             .TextMatrix(0, 1) = "����"
             .TextMatrix(0, 2) = "SN"
             .TextMatrix(0, 3) = "MAC"
             .TextMatrix(0, 4) = "�޸�ʱ��"
        End With
    Else
         MsgBox "ϵͳ��û������!", vbOKOnly + vbExclamation, "��ѯ���û�м�¼,��ȷ�������Ƿ���ȷ��"
    End If
    rec.Close
    Me.cmdQuery.Enabled = True
End Sub
