VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmUnLock 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ž���"
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
   Icon            =   "frmUnLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmUnLock.frx":073E
   ScaleHeight     =   8505
   ScaleWidth      =   11415
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.TextBox txtHidBoxid 
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
         Left            =   10560
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.CommandButton cmdDeleteBoxID 
         Height          =   495
         Left            =   8400
         Picture         =   "frmUnLock.frx":A21B
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   495
         Left            =   6600
         Picture         =   "frmUnLock.frx":AB0C
         Style           =   1  'Graphical
         TabIndex        =   5
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
         TabIndex        =   4
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
         Height          =   390
         Left            =   3120
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblCodeNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���(Box ID):"
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
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblArea 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ˮ��(SN Number):"
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
         Index           =   1
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      Height          =   7335
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   11415
      Begin VB.CommandButton cmdPrintBefore 
         Height          =   495
         Left            =   6720
         Picture         =   "frmUnLock.frx":B17E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearchUnLock 
         Height          =   495
         Left            =   8520
         Picture         =   "frmUnLock.frx":B7DC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5520
         Width           =   2775
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridResult 
         Height          =   5055
         Left            =   120
         TabIndex        =   7
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
Attribute VB_Name = "frmUnLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As New ADODB.Recordset
Dim sql As String

Private Sub cmdDeleteBoxID_Click()
    If Me.txtHidBoxid.Text = "" Then
    Else
         Dim intR As Integer
        intR = MsgBox("ȷ�������������?", vbOKCancel, "����ȷ��")

        If intR = 1 Then
            Dim recTT As New ADODB.Recordset
            sql = "select *,getdate(),'" & golUSERNAME & "' from tblUNIT where boxid='" & Me.txtHidBoxid.Text & "'"
            recTT.Open sql, conn, adOpenKeyset, adLockOptimistic
            Do While Not recTT.EOF
                Dim str As String
                'str = "insert into tblUNIT_Unlock values('" & recTT.Fields(0) & "'," & recTT.Fields(1) & ",'" & recTT.Fields(2) & "','" & recTT.Fields(3) & "','" & recTT.Fields(4) & "','" & recTT.Fields(5) & "','" & recTT.Fields(6) & "','" & recTT.Fields(7) & "','" & recTT.Fields(8) & "','" & recTT.Fields(9) & "'," & recTT.Fields(10) & "," & IIf(recTT.Fields(11) = "False", 0, 1) & ",'" & recTT.Fields(12) & "',convert(datetime,'" & recTT.Fields(13) & "'),getdate(),'" & golUSERNAME & "')"
                str = "insert into tblUNIT_Unlock values('" & recTT.Fields(0) & "'," & recTT.Fields(1) & ",'" & recTT.Fields(2) & "','" & recTT.Fields(3) & "','" & recTT.Fields(4) & "','" & recTT.Fields(5) & "','" & recTT.Fields(6) & "','" & recTT.Fields(7) & "','" & recTT.Fields(8) & "','" & recTT.Fields(9) & "'," & recTT.Fields(10) & ",'" & recTT.Fields(11) & "','" & recTT.Fields(12) & "',convert(datetime,'" & recTT.Fields(13) & "'),getdate(),'" & golUSERNAME & "')"
                conn.Execute str
            
            recTT.MoveNext
            Loop
            recTT.Close
            
        
            sql = "Delete from tblUNIT where boxid='" & Me.txtHidBoxid.Text & "' "
            conn.Execute sql
            
            sql = "Delete from tblUNIT_Tmp where boxid='" & Me.txtHidBoxid.Text & "' "
            
            conn.Execute sql
            
        
            MsgBox "������Ѿ�����"
            txtSN.Text = ""
            txtBOXID.Text = ""
            
            Me.gridResult.Clear
            
        End If
    End If
    
   
    
End Sub

Private Sub cmdPrintBefore_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()

    If rec.State = 1 Then
        rec.Close
    End If
   
    'sql = "Select no=Identity(int,1,1),* Into #UNIT_temptable From tblUNIT where BOXID in(Select BoxID From tblUNIT where sn='" & Trim(txtSN.Text) & "') or boxid='" & Trim(txtBOXID.Text) & "';Select * From #UNIT_temptable; Drop Table #UNIT_temptable"
    
    If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   
    'sql = " select SN, BarCodeNum, VendorCode, VendorName, ContpactNo, PDate, Model, Description, BoxID, Rev, Quality, case RoHS when 1 then 'RoHS' else 'Non-RoHS' end as RoHS, UserID, CONVERT(varchar(100), PrintTime, 20) as PrintTime from tblUNIT where BOXID in(Select BoxID From tblUNIT where sn='" & Trim(txtSN.Text) & "') or boxid='" & Trim(txtBOXID.Text) & "' "
    sql = " select SN, BarCodeNum, VendorCode, VendorName, ContpactNo, PDate, Model, Description, BoxID, Rev, Quality, RoHS, UserID, CONVERT(varchar(100), PrintTime, 20) as PrintTime from tblUNIT where BOXID in(Select BoxID From tblUNIT where sn='" & Trim(txtSN.Text) & "') or boxid='" & Trim(txtBOXID.Text) & "' "
    rec.Open sql, conn, adOpenKeyset, adLockOptimistic
    If rec.EOF = False Then
         Me.txtHidBoxid.Text = rec.Fields(8)
    End If
   
     Set gridResult.DataSource = rec
    
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
        
        .TextMatrix(0, 1) = "��Ʒ����"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "���̴���"
        .TextMatrix(0, 4) = "��������"
        .TextMatrix(0, 5) = "��ͬ��"
        .TextMatrix(0, 6) = "����"
        .TextMatrix(0, 7) = "���ֱ���"
        .TextMatrix(0, 8) = "����"
        .TextMatrix(0, 9) = "���"
        .TextMatrix(0, 10) = "����汾"
        .TextMatrix(0, 11) = "�ں�����"
        .TextMatrix(0, 12) = "RoHS"
        .TextMatrix(0, 13) = "�û�"
        .TextMatrix(0, 14) = "��ӡʱ��"
   End With
   
    Me.txtBOXID.Text = ""
    Me.txtSN.Text = ""
    
End Sub

Private Sub cmdSearchUnLock_Click()
    frmUnLockSearch.Show
End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   Me.Show
   
   txtSN.SetFocus
   
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
        lWidth = (LenB(StrConv(gridResult.TextMatrix(Row, Col), vbFromUnicode)) + 1) * gridResult.FontSize * 16                         '16�ɰ�������������
        If Row = 0 Then
                gridResult.ColWidth(Col) = lWidth
        ElseIf gridResult.ColWidth(Col) < lWidth Then
                gridResult.ColWidth(Col) = lWidth
        End If
End Sub
