VERSION 5.00
Begin VB.Form frm21huaweiPrint_H3C 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H3C ����ģ�����ǩ��ӡ"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm21huaweiPrint_H3C.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   10665
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox tbFirst 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2160
      TabIndex        =   25
      Top             =   6120
      Width           =   255
   End
   Begin VB.CommandButton cmdGoon 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9360
      TabIndex        =   19
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "��ͣ"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7920
      TabIndex        =   18
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ӡ(Print) &p"
      Height          =   615
      Left            =   480
      TabIndex        =   13
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Height          =   615
      Left            =   5400
      TabIndex        =   12
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Frame fmVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   10455
      Begin VB.CheckBox chkN_ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N*"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3840
         TabIndex        =   24
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkY2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y2"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y*"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtWO 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   1680
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   9120
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtXH 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         TabIndex        =   15
         Top             =   720
         Width           =   3135
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
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   450
         Left            =   3840
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������:"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ӡ����:"
         Height          =   375
         Left            =   7800
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ�ͺ�:"
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblChinaRoHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��������:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblCPN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ʒ����:"
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ʼ����:"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblDes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "һʽ����:"
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblVer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�汾:"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.PictureBox picH3C 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      Picture         =   "frm21huaweiPrint_H3C.frx":13652
      ScaleHeight     =   1545
      ScaleWidth      =   10425
      TabIndex        =   4
      Top             =   480
      Width           =   10455
   End
End
Attribute VB_Name = "frm21huaweiPrint_H3C"
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
  Dim Ctr As Control
    For Each Ctr In Me.Controls
        If TypeOf Ctr Is TextBox Then
            Ctr.Text = ""
        End If
        If TypeOf Ctr Is CheckBox Then
            Ctr.Value = 0
        End If
    Next
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
    
    If Me.chkY2.Value = 0 And Me.chkY.Value = 0 Then
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
   Dim leftstr, rightstr, str As String
   qty = CInt(txtQty.Text)
   qty1 = CInt(txtQty1.Text)
   leftstr = UCase(Left(txtSN.Text, 14))
'   rightstr = Right(txtSN.Text, 6)
   rightstr = tbFirst.Text + Right(txtSN.Text, 5)
   
   
   '' To double check the value of pb
    Dim Pb As String
    If chkY2.Value = 1 Then
         Pb = "Y2"
    ElseIf chkY.Value = 1 Then
         Pb = "Y*"
    End If
'    If (MsgBox("��������Ϊ<" & pb & ">,�Ƿ������ӡ", vbYesNo, "ȷ����Ϣ") <> vbYes) Then
'        cmdPrint.Caption = "��ӡ(Print) &p"
'        cmdPrint.Enabled = True
'        Exit Sub
'    End If
   
   
   
     OpenLppx
     
     bRun = True
     
   Dim k As Integer
   k = 0
   
   For i = 0 To qty - 1
      str = leftstr & "0" & Right("000000" & CStr(CLng(rightstr) + i), 5)
      
    '======Add by mike 2015.3.24 for data upload to FTPC============
    If UploadH3C_PB(Pb, Trim(str), Trim(Me.txtVer.Text), "NA", "N/A", "CHINA", golUSERNAME) = False Then
        MsgBox "PB���ϱ���ʧ�ܲ��ܴ�ӡ!", vbInformation + vbOKOnly, "���ϱ���ʧ��"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
    
      '===============add by Carson 2015-12-15 start===============
    If reprint = False Then
       If Connect.isPrintedLabel(Trim(str), Me.Name) = True Then
           MsgBox ("�����к��Ѵ�ӡ��")
           txtSN.SetFocus
           UnloadLppx
           cmdCancel_Click
           cmdPrint.Caption = "��ӡ(Print) &p"
           cmdPrint.Enabled = True
           Exit Sub
       End If
               
    End If
    '===============add by Carson 2015-12-15 end=================
    
    For j = 0 To qty1 - 1
        
        If bRun = True Then
        
            If k > 0 And k Mod 100 = 0 Then
                Savetime = timeGetTime '���¿�ʼʱ��ʱ��
                While timeGetTime < Savetime + 30000 'ѭ���ȴ�
                    DoEvents 'ת�ÿ���Ȩ���Ա��ò���ϵͳ�����������¼���
                Wend
            End If
      
keepprint:
            myVars.Item("sn").Value = str
            'myVars.Item("Item").Value = "03" & UCase(Left(txtSN.Text, 6))
            If txtVer.Text = "" Or txtVer.Text = "/" Then
                'myObjs("Sver").Top = 5
                myVars.Item("ver").Value = "N/A"
            ElseIf Me.txtVer.Text <> "" Then
                'modified by noel.zhou 2014/4/16
                myVars.Item("ver").Value = Trim(txtVer.Text)
            Else
                'myObjs("Sver").Top = 5
                myVars.Item("ver").Value = UCase(txtVer.Text)
            End If
            myVars.Item("Type").Value = txtXH.Text
            
            If Me.chkY.Value = 1 Then
                myVars.Item("Rohs").Value = "Y*"
            End If
            If Me.chkY2.Value = 1 Then
                myVars.Item("Rohs").Value = "Y2"
            End If
   
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

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub cmdStop_Click()
    bRun = False
    cmdPrint.Enabled = False
    cmdCancel.Enabled = False
    cmdReturn.Enabled = False
    cmdStop.Enabled = False
    cmdGoon.Enabled = True
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
      txtVer.Enabled = False
      If Len(txtSN.Text) <> 20 Then
         MsgBox "��Ʒ��ų��ȱ���Ϊ20λ!"
         txtSN.SetFocus
         Exit Sub
      End If

        Dim rcd As New ADODB.Recordset
        sql = "select * from tblCustomType where PartNumber='" & Mid(txtSN.Text, 3, 8) & "'"
        rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rcd.EOF = True Then
           MsgBox "Ʒ��δά��!"
           txtSN.Text = ""
           txtSN.SetFocus
           rcd.Close
           Exit Sub
        Else
            If rcd.Fields(1) = "H3C" Then
                MsgBox "��ʹ��[H3C����ģ�����ǩ��ӡ]��ӡ!"
                txtSN.Text = ""
                txtSN.SetFocus
                rcd.Close
                Exit Sub
            End If
        End If
        rcd.Close
      
        
      sql = "select ID,SN,TYPE,CASE PB WHEN 1 THEN 'Yes' when 0 then 'No' else 'Non' end from SingleUnit where SN='" & Mid(txtSN.Text, 3, 8) & "'"
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
      txtWO.SetFocus
   Else
    chkN_.Value = 0
    chkY.Value = 0
    chkY2.Value = 0
      txtCPN.Text = ""
      txtXH.Text = ""
      txtWO.Text = ""
      txtVer.Text = ""
   End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   Set myDoc = myApp.Documents.Open("\\10.11.1.25\Public\Manufacture\��ǩģ��\" & "21HUAWEI.lab")
   'Set myDoc = myApp.Documents.Open("G:\flash\��ǩģ��\" & "H3C.lab")
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub


Private Sub txtWO_KeyPress(KeyAscii As Integer)
    Dim tempWO As String
    If (KeyAscii = 13) Then
        If Len(Trim(txtCPN.Text)) <> 8 Then
            MsgBox "��Ʒ���볤�ȱ���Ϊ8λ!"
            txtSN.SetFocus
            Exit Sub
        Else
            tempWO = txtWO.Text
            If tempWO = "" Or tempWO = Null Then Return
            If UCase(tempWO) = "TASK" Then
                txtVer.Enabled = True
                chkY.Enabled = True
                chkY2.Enabled = True
                chkN_.Enabled = True
                Exit Sub
            End If
'            While (Len(tempWO) < 12)
'                tempWO = "0" & tempWO
'            Wend
'            sql = "select MaterialRevision from [10.11.1.17].dsActive.dbo.SAP_WO " & _
'                "where WorkOrderNumber = '" & tempWO & "' and ( MaterialNumber like 'HWF" & txtCPN.Text & "%' " & _
'                "or MaterialNumber like 'HUV" & txtCPN.Text & "%' ) "
            sql = "select part_revision,(select order_type_S from [10.11.1.130].afg_active_90.dbo.UDA_Order where object_key=A.order_key) order_type from [10.11.1.130].afg_active_90.dbo.WORK_ORDER A,[10.11.1.130].afg_active_90.dbo.WORK_ORDER_ITEMS B " & _
                "WHERE A.order_key = B.order_key AND A.order_number ='" & tempWO & "' and ( part_number like 'HWF" & txtCPN.Text & "%')"
'                 "or part_number like 'HUV" & txtCPN.Text & "%' ) "
            rec.Open sql, conn, adOpenForwardOnly, adLockReadOnly
            
            
            If rec.EOF = True Then
                MsgBox "SAP�д˹����ı������˲�Ʒ���벻һ�»��߸Ĺ�����HUV����"
                txtWO.Text = ""
                txtVer.Text = ""
                txtWO.SetFocus
                rec.Close
                Exit Sub
            Else
                txtVer.Text = Trim(rec.Fields(0))
                If rec.Fields(1) = "PP05" Then
                    txtVer.Enabled = True
                    chkY.Enabled = True
                    chkY2.Enabled = True
                    chkN_.Enabled = True
                    rec.Close
                    Exit Sub
                End If
                partlist = getPartListByOrder(tempWO)
                If partlist = "" Then
                    MsgBox "�Ҳ����Ĺ���" & tempWO & "��Ӧ��0302�׵���Ϣ,��ȷ��SAP�Ƿ��ͷŸĹ������Ҵ���MESϵͳ��"
                    txtWO.Text = ""
                    txtVer.Text = ""
                    txtWO.SetFocus
                    Exit Sub
                End If
                
                Dim cmd As New ADODB.Command
                cmd.ActiveConnection = conn
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "[PbHandler]"
                cmd.Parameters.Append cmd.CreateParameter("partlist", adVarChar, adParamInput, 8000, partlist)
                cmd.Parameters.Append cmd.CreateParameter("res", adVarChar, adParamOutput, 8)
                cmd.Parameters.Append cmd.CreateParameter("first", adVarChar, adParamOutput, 1)
                cmd.Execute
'                tbFirst.Text = cmd("first")    '' Cancel the 9 principle
                tbFirst.Text = "0"
                Select Case cmd("res")
                    Case "No"
                        Me.chkY2.Value = 1
                        Me.chkY2.Enabled = False
                        Me.chkY.Value = 0
                        Me.chkY.Enabled = False
                    Case "Non"
                        MsgBox "�˹�������0302�׵���δ�趨��Ǧ��Ǧ,�����MEȥ�趨!"
                        txtWO.Text = ""
                        txtVer.Text = ""
                        txtWO.SetFocus
                        Exit Sub
                    Case "Half"
                       chkY.Enabled = False
                       chkY.Value = 1
                       chkY2.Value = 0
                       chkY2.Enabled = False
                    Case "Yes"
                       chkY.Value = 1
                       chkY.Enabled = False
                       chkY2.Value = 0
                       chkY2.Enabled = False
                End Select
                
                OverridePb ''check label history
                
            End If
        End If
    Else
        txtVer.Text = ""
    End If
End Sub

Private Sub OverridePb()
    Dim labelHistory As New Label_History
    Dim sn As String
    sn = txtSN.Text
    If labelHistory.Init(sn) Then
    
        chkY.Value = 0
        chkY2.Value = 0
       
        Select Case labelHistory.Pb
        Case "Y*"
            chkY.Value = 1
        Case "Y2"
            chkY2.Value = 1

        End Select
        
    End If
End Sub
