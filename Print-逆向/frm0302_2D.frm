VERSION 5.00
Begin VB.Form frm0302_2D 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "0302�׶�ά���ǩ"
   ClientHeight    =   6432
   ClientLeft      =   48
   ClientTop       =   408
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6432
   ScaleWidth      =   10920
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdGoon 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9120
      TabIndex        =   21
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "��ͣ"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7680
      TabIndex        =   20
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ӡ(Print) &p"
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   615
      Left            =   5280
      TabIndex        =   14
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   615
      Left            =   2760
      TabIndex        =   13
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   10455
      Begin VB.CheckBox chkN4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N4"
         Height          =   375
         Left            =   4200
         TabIndex        =   27
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkNx 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N*"
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkN1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N1"
         Height          =   375
         Left            =   1680
         TabIndex        =   25
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkN3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N3"
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkY4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y4"
         Height          =   375
         Left            =   4200
         TabIndex        =   23
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox chkY3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y3"
         Height          =   375
         Left            =   3360
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   9120
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtXH 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   17
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y1"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkY2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y2"
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtQty1 
         Height          =   405
         Left            =   6840
         TabIndex        =   3
         Text            =   "1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtCPN 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtVer 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ӡ����:"
         Height          =   375
         Left            =   7800
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ�ͺ�:"
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��������:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ʼ����:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "һʽ����:"
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�汾:"
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5940
      Left            =   0
      Picture         =   "frm0302_2D.frx":0000
      ScaleHeight     =   5916
      ScaleWidth      =   10884
      TabIndex        =   6
      Top             =   0
      Width           =   10905
   End
End
Attribute VB_Name = "frm0302_2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''TBT

Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim Pb As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myFormat As LabelManager2.Format
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Dim bRun As Boolean

Private Sub chkCE_Click()
   If chkCE.Value = 1 Then
      chkNonCE.Value = 0
   Else
      chkNonCE.Value = 1
   End If
End Sub

Private Sub chkNonCE_Click()
   If chkNonCE.Value = 1 Then
      chkCE.Value = 0
   Else
      chkCE.Value = 1
   End If
End Sub

Private Sub chkNonChinaRoHS_Click()
   If chkNonChinaRoHS.Value = 1 Then
      chkChinaRoHS.Value = 0
   'Else
   '   chkChinaRoHS.Value = 1
   End If
End Sub

Private Sub chkNonWEEE_Click()
   If chkNonWEEE.Value = 1 Then
      chkWEEE.Value = 0
   Else
      chkWEEE.Value = 1
   End If
End Sub

Private Sub chkOS_Click()
   If chkOS.Value = 1 Then
      txtOS.Enabled = True
      txtOS.BackColor = &H80000005
   Else
      txtOS.Enabled = False
      txtOS.BackColor = &HC0C0C0
   End If
End Sub

Private Sub chkChinaRoHS_Click()
   If chkChinaRoHS.Value = 1 Then
      chkNonChinaRoHS.Value = 0
   'Else
   '   chkNonChinaRoHS.Value = 1
   End If
End Sub

Private Sub chkWEEE_Click()
   If chkWEEE.Value = 1 Then
      chkNonWEEE.Value = 0
   Else
      chkNonWEEE.Value = 1
   End If
End Sub

Private Sub cmdCancel_Click()
   txtSN.Text = ""
   txtVer.Text = ""
   txtCPN.Text = ""
   txtXH.Text = ""
   chkY.Value = 0
   chkY2.Value = 0
   chkY4.Value = 0
   chkN1.Value = 0
   chkN3.Value = 0
   chkN4.Value = 0
   chkNx.Value = 0
   txtQty.Text = ""
   txtSN.SetFocus
End Sub

Private Sub cmdGoon_Click()
    bRun = True
    cmdPrint.Enabled = False
    cmdCancel.Enabled = True
    cmdReturn.Enabled = True
    cmdStop.Enabled = True
    cmdGoon.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    If chkY.Value + Me.chkY2.Value + chkY3.Value + chkY4.Value + chkN1.Value + chkN3.Value + chkN4.Value + chkNx.Value <> 1 Then
        MsgBox "��������δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ���뻷������"
        txtSN.SetFocus
        Exit Sub
    End If
    
    sql = "select active from tblECO_Ver where PartNumber='" & Trim(txtCPN.Text) & "' and Version='" & Trim(txtVer.Text) & "'"
    If rec.State = 1 Then
      rec.Close
    End If
   
   rec.Open sql, conn, adOpenKeyset, adLockOptimistic
   If rec.EOF = False Then
        If rec.Fields(0).Value = "False" Then
            MsgBox "�˰汾�Ѿ�������,���ܴ�ӡ!", vbInformation + vbOKOnly, "�汾�Ѿ�������"
            txtSN.SetFocus
            Exit Sub
        End If
   End If
   rec.Close


  If txtSN.Text = "" Then
      MsgBox "��Ʒ����δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ�����Ʒ����"
      txtSN.SetFocus
      Exit Sub
   End If
   
   If txtQty.Text = "" Then
      MsgBox "����δ���룬���ܴ�ӡ��", vbInformation + vbOKOnly, "δ��������"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If CInt(txtQty.Text) = 0 Then
      MsgBox "��������ȷ��������", vbInformation + vbOKOnly, "��������"
      txtQty.SetFocus
      Exit Sub
   End If
   
   If txtQty1.Text = "" Then
      MsgBox "һʽ����δ���룬���ܴ�ӡ��", vbInformation + vbOKOnly, "δ��������"
      txtQty1.SetFocus
      Exit Sub
   End If
   
   If CInt(txtQty1.Text) = 0 Then
      MsgBox "��������ȷ��������", vbInformation + vbOKOnly, "��������"
      txtQty1.SetFocus
      Exit Sub
   End If
   
   
   If txtVer.Text = "" Then
      MsgBox "�汾δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ����汾"
      txtWO.SetFocus
      Exit Sub
   End If
   
   If txtXH.Text = "" Then
      MsgBox "�ͺ�δ����,���ܴ�ӡ!", vbInformation + vbOKOnly, "δ�����ͺ�"
      txtXH.SetFocus
      Exit Sub
   End If
   
   
   cmdPrint.Caption = "ִ����..."
   cmdPrint.Enabled = False
   cmdStop.Enabled = True
    
   Dim i, j, qty, qty1 As Integer
   Dim leftstr, rightstr, str As String, str1 As String, str2 As String, str3 As String, str4 As String, endStr As String
   
   endStr = "XXXXXXXXXXXXXXXXXXXX"
   
   
   qty = CInt(txtQty.Text)
   qty1 = CInt(txtQty1.Text)
   leftstr = UCase(Left(txtSN.Text, Len(txtSN.Text) - 6))
   
   If chkY2.Value = 1 Then
        Pb = "Y2"
   ElseIf chkY.Value = 1 Then
        Pb = "Y1"
   ElseIf chkY3.Value = 1 Then
        Pb = "Y3"
   ElseIf chkY4.Value = 1 Then
        Pb = "Y4"
   ElseIf chkN1.Value = 1 Then
        Pb = "N1"
   ElseIf chkN3.Value = 1 Then
        Pb = "N3"
   ElseIf chkN4.Value = 1 Then
        Pb = "N4"
   ElseIf chkNx.Value = 1 Then
        Pb = "N*"
   End If
   
    '' To double check the value of pb
'    If (MsgBox("��������Ϊ<" & pb & ">,�Ƿ������ӡ", vbYesNo, "ȷ����Ϣ") <> vbYes) Then
'        cmdPrint.Caption = "��ӡ(Print) &p"
'        cmdPrint.Enabled = True
'        Exit Sub
'    End If
    
   
'   If Pb = "Y2" Then
''      rightstr = "9" + Right(txtSN.Text, 5)
'      'rightstr = "0" + Right(txtSN.Text, 5)
'
'   Else
'      rightstr = "0" + Right(txtSN.Text, 5)
'   End If
   rightstr = Right(txtSN.Text, 6)
   OpenLppx
     
    bRun = True
    Dim k As Integer
    k = 0
    Dim strPreviousLength As Integer, strFinalLength As Integer
    Dim strFinal As String, strFinal1 As String, strFinal2 As String, strFinal3 As String, strFinal4 As String

    

   For i = 0 To qty - 1 Step 1
'      str = leftstr & Right("000000" & CStr(CInt(rightstr) + i), 6)
'==================edit by ben 2011-10-14 start========================
''20180122       strPreviousLength = Len(rightstr)
''       strFinal = CStr(CLng(rightstr) + i)
''       strFinalLength = Len(strFinal)
''       For m = strprevisouslength To strFinalLength - 1
''            strFinal = "0" + strFinal
''       Next
''       str = leftstr & Right("000000" & strFinal, 6)
''       If i + 1 > qty - 1 Then
''            strFinal1 = endStr
''            strFinalLength = Len(strFinal1)
''       Else
''            strFinal1 = CStr(CLng(rightstr) + i + 1)
''            strFinalLength = Len(strFinal1)
''       End If
''
''       For m = strprevisouslength To strFinalLength - 1
''            strFinal1 = "0" + strFinal1
''       Next
''20180122       str1 = leftstr & Right("000000" & strFinal1, 6)
        str = leftstr & rightstr
        str1 = str
       
       '======Add by mike 2015.3.24 for data upload to FTPC============
       If UploadH3CInfo(Pb, Trim(str), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
            Or UploadH3CInfo(Pb, Trim(str1), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False Then
           MsgBox "���ϱ���ʧ�ܲ��ܴ�ӡ!", vbInformation + vbOKOnly, "���ϱ���ʧ��"
           txtSN.SetFocus
           UnloadLppx
           Exit Sub
       End If
       
       If UploadH3C_PB(Pb, Trim(str), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False _
            Or UploadH3C_PB(Pb, Trim(str1), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False Then
           MsgBox "PB���ϱ���ʧ�ܲ��ܴ�ӡ!", vbInformation + vbOKOnly, "���ϱ���ʧ��"
           txtSN.SetFocus
           UnloadLppx
           Exit Sub
       End If
      
       '======Add by mike 2015.3.24 for data upload to FTPC============
       
'==================edit by ben 2011-10-14 end==========================
    For j = 0 To qty1 - 1
 
        If bRun = True Then
            If k > 0 And k Mod 100 = 0 Then
                Savetime = timeGetTime '���¿�ʼʱ��ʱ��
                While timeGetTime < Savetime + 30000 'ѭ���ȴ�
                    DoEvents 'ת�ÿ���Ȩ���Ա��ò���ϵͳ�����������¼���
                Wend
            End If
keepprint:
'            myVars.Item("2D").Value = str
'            myVars.Item("2D2").Value = str1
'            myVars.Item("2D3").Value = str2
            myVars.Item("SN1").Value = str
            myVars.Item("SN2").Value = str1
            
            
            'myVars.Item("Item").Value = "03" & UCase(Left(txtSN.Text, 6))
            If txtVer.Text = "" Or txtVer.Text = "/" Then
                'myObjs("Sver").Top = 5
                myVars.Item("Rev").Value = "N/A"
            ElseIf txtVer.Text = "00" Then
                myVars.Item("Rev").Value = ""
            Else
                'myObjs("Sver").Top = 5
                myVars.Item("Rev").Value = UCase(txtVer.Text)
            End If
            myVars.Item("Type").Value = txtXH.Text
            
 
            myVars.Item("Rohs").Value = Pb
 
            'myApp.Visible = True
            myDoc.PrintLabel 1
            myDoc.FormFeed
   
            k = k + 1
            
            DoEvents
            
        Else
            While (bRun = False)
                'sleep 1000
                DoEvents
            Wend
            
            GoTo keepprint
        End If
   Next
   Next
   
   UnloadLppx
    

   cmdCancel_Click
   
   cmdPrint.Caption = "��ӡ(Print) &p"
   cmdPrint.Enabled = True
   
End Sub

Private Sub cmdStop_Click()
    bRun = False
    cmdPrint.Enabled = False
    cmdCancel.Enabled = False
    cmdReturn.Enabled = False
    cmdStop.Enabled = False
    cmdGoon.Enabled = True
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If
   If connFTPC.State = 0 Then
      connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
      connFTPC.Open
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
   bRun = False
End Sub

Private Sub txtHV_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 13) Then
     txtMS.SetFocus
  End If
End Sub

Private Sub txtMS_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtNAL.SetFocus
  End If
End Sub



Private Sub txtNAL_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtRemark.SetFocus
  End If
End Sub


Private Sub txtQty_Change()
    If txtQty.Text <> "" Then
        If Asc(Right(txtQty.Text, 1)) > 57 Or Asc(Right(txtQty.Text, 1)) < 48 Then
           MsgBox "ֻ���������֣�", vbInformation + vbOKOnly, "���벻��ȷ"
           SendKeys "{backspace}"
           txtQty.SetFocus
           Exit Sub
        End If
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     txtVer.SetFocus
  End If
End Sub



Private Sub txtRemark_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
     cmdPrint_Click
  End If
End Sub

Private Sub txtQty1_Change()
    If txtQty1.Text <> "" Then
        If Asc(Right(txtQty1.Text, 1)) > 57 Or Asc(Right(txtQty1.Text, 1)) < 48 Then
           MsgBox "ֻ���������֣�", vbInformation + vbOKOnly, "���벻��ȷ"
           SendKeys "{backspace}"
           txtQty1.SetFocus
           Exit Sub
        End If
    End If
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Me.txtSN.Text = Trim(Me.txtSN.Text)
        If Len(Trim(txtSN.Text)) <> 16 Then
           MsgBox "��Ʒ��ų��ȱ���Ϊ16λ��02�׵���!"
           txtSN.SetFocus
           Exit Sub
        End If
        
        If Left(Trim(txtSN.Text), 2) <> "02" Then
           MsgBox "��Ʒ�����ʼ����02!"
           txtSN.SetFocus
           Exit Sub
        End If
        
        If Left(txtSN.Text, 2) = "21" Then
             sql = "select ID,SN,TYPE,CASE PB WHEN 1 THEN 'Yes' when 0 then 'No' else 'Non' end from SingleUnit where SN= '" & Mid(txtSN.Text, 3, 8) & "'"
        Else
             sql = "select ID,SN,TYPE,CASE PB WHEN 1 THEN 'Yes' when 0 then 'No' else 'Non' end from SingleUnit where SN='03" & Mid(txtSN.Text, 1, 6) & "'"
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
        
        End If
        rec.Close
        txtVer.SetFocus
    Else
         txtVer.Text = ""
         txtCPN.Text = ""
         txtXH.Text = ""
    End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\�����ǩģ��\" & "0302�����ά���ǩ.Lab")
   Set myFormat = myDoc.Format
   'Set myDoc = myApp.Documents.Open("G:\flash\��ǩģ��\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

