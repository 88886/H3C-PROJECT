VERSION 5.00
Begin VB.Form frmBatchPrint 
   Caption         =   "������ӡ"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   LinkTopic       =   "������ӡ"
   ScaleHeight     =   6705
   ScaleWidth      =   11670
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chkNonChinaRoHS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Non-Rohs"
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CheckBox chkChinaRoHS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rohs"
      Height          =   375
      Left            =   6840
      TabIndex        =   15
      Top             =   3240
      Width           =   1095
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      Picture         =   "frmBatchPrint.frx":0000
      ScaleHeight     =   1545
      ScaleWidth      =   10425
      TabIndex        =   14
      Top             =   0
      Width           =   10455
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   10455
      Begin VB.TextBox txtVer 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1680
         TabIndex        =   8
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   1605
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtQty1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6800
         TabIndex        =   5
         Text            =   "1"
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox txtXH 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   4
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��������:"
         Height          =   375
         Left            =   5160
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�汾:"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "һʽ����:"
         Height          =   375
         Left            =   5160
         TabIndex        =   12
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������ӡ:"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ�ͺ�:"
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   615
      Left            =   4800
      TabIndex        =   2
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   615
      Left            =   7080
      TabIndex        =   1
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ӡ(Print) &p"
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblChinaRoHS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��������:"
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "frmBatchPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Private Sub cmdCancel_Click()
   txtSN.Text = ""
   txtVer.Text = ""
   txtCPN.Text = ""
   txtXH.Text = ""
   txtSN.SetFocus
End Sub

Private Sub cmdReturn_Click()
    Unload Me
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
'Private Sub txtQty1_Change()
'If txtQty1.Text <> "" Then
'    If Asc(Right(txtQty1.Text, 1)) > 57 Or Asc(Right(txtQty1.Text, 1)) < 48 Then
'       MsgBox "ֻ���������֣�", vbInformation + vbOKOnly, "���벻��ȷ"
'       SendKeys "{backspace}"
'       txtQty1.SetFocus
'       Exit Sub
'    End If
'End If
'End Sub
Private Sub cmdPrint_Click()
    sql = "select active from tblECO_Ver where PartNumber='" & Trim(txtCPN.Text) & "' and Version='" & Trim(txtVer.Text) & "'"
    If rec.State = 1 Then
      rec.Close
    End If
   
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   If rec.EOF = False Then
        If rec.Fields(0).Value = "False" Then
'            MsgBox "�˰汾�Ѿ�������,���ܴ�ӡ!", vbInformation + vbOKOnly, "�汾�Ѿ�������"
'            txtSN.SetFocus
'            Exit Sub
        End If
   End If
   rec.Close

   If txtSN.Text = "" Then
      MsgBox "��Ʒ����δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ�����Ʒ����"
      txtSN.SetFocus
      Exit Sub
   End If
   
'   If txtQty.Text = "" Then
'      MsgBox "����δ���룬���ܴ�ӡ��", vbInformation + vbOKOnly, "δ��������"
'      txtQty.SetFocus
'      Exit Sub
'   End If
'
'   If CInt(txtQty.Text) = 0 Then
'      MsgBox "��������ȷ��������", vbInformation + vbOKOnly, "��������"
'      txtQty.SetFocus
'      Exit Sub
'   End If
'
'      If txtQty1.Text = "" Then
'      MsgBox "һʽ����δ���룬���ܴ�ӡ��", vbInformation + vbOKOnly, "δ��������"
'      txtQty1.SetFocus
'      Exit Sub
'   End If
   
   If CInt(txtQty1.Text) = 0 Then
      MsgBox "��������ȷ��������", vbInformation + vbOKOnly, "��������"
      txtQty1.SetFocus
      Exit Sub
   End If
   
   
   If txtVer.Text = "" Then
      MsgBox "�汾δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ����汾"
      txtVer.SetFocus
      Exit Sub
   End If
   
   If txtXH.Text = "" Then
      MsgBox "�ͺ�δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ�����ͺ�"
      txtXH.SetFocus
      Exit Sub
   End If
   
   If Me.chkChinaRoHS = 0 And Me.chkNonChinaRoHS.Value = 0 Then
        MsgBox "��������δѡ��,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ�����ͺ�"
        Exit Sub
   End If
   
   Dim barcode() As String
   Dim str As String
   barcode = Split(Trim(txtSN.Text), vbCrLf, , vbTextCompare)
   OpenLppx
   For i = 0 To UBound(barcode)
      str = barcode(i)
      If Len(str) = 16 Or Len(str) = 20 Then
            myVars.Item("sn").Value = Trim(str)
            If txtVer.Text = "" Or txtVer.Text = "/" Then
               myVars.Item("ver").Value = "N/A"
            Else
              myVars.Item("ver").Value = Trim(UCase(txtVer.Text))
            End If
            myVars.Item("type").Value = Trim(txtXH.Text)
          
           If Me.chkChinaRoHS.Value = 1 Then
              myVars.Item("rohs").Value = "Y*"
           End If
            
           If Me.chkNonChinaRoHS.Value = 1 Then
              myVars.Item("rohs").Value = "N*"
           End If
          
        
         'myApp.Visible = True
        
         myDoc.PrintLabel 1
         myDoc.FormFeed
    End If
   Next
   UnloadLppx
   cmdCancel_Click
End Sub


Private Sub txtSN_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
   Dim barcode() As String
   Dim str As String
   barcode = Split(Trim(txtSN.Text), vbCrLf, , vbTextCompare)
   If UBound(barcode) = 0 Then
      txtSN.Text = ""
      txtSN.SetFocus
      Exit Sub
   End If
   str = barcode(0)
   

If Left(str, 2) = "21" Then
   sql = "select * from SingleUnit where SN= '" & Mid(str, 3, 8) & "'"
Else
   sql = "select * from SingleUnit where SN='03" & Mid(str, 1, 6) & "'"
End If


      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "�˲�Ʒ����δ��������!"
         txtVer.Text = ""
         txtSN.Text = ""
         txtSN.SetFocus
         rec.Close
         Exit Sub
      Else
        txtCPN.Text = Trim(rec.Fields(1))
        txtXH.Text = Trim(rec.Fields(2))

'        If UCase(Trim(rec.Fields(3))) = "ROHS" Then
'           chkChinaRoHS.Value = 1
'           chkNonChinaRoHS.Value = 0
'        ElseIf rec.Fields(3) = "/" Then
'           chkChinaRoHS.Value = 0
'           chkNonChinaRoHS.Value = 1
'        End If


      End If
      rec.Close
      txtVer.SetFocus
   End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\21H3C.lab")
   'Set myDoc = myApp.Documents.Open("G:\21H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

