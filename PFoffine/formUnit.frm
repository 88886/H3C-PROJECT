VERSION 5.00
Begin VB.Form formUnit 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   Picture         =   "formUnit.frx":0000
   ScaleHeight     =   7470
   ScaleWidth      =   10605
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox tbMACFrom3 
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   5760
      Width           =   2055
   End
   Begin VB.TextBox tbMACTo3 
      Height          =   375
      Left            =   4680
      TabIndex        =   18
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox tbMACFrom2 
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox tbMACTo2 
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox txtWOHID 
      Height          =   270
      Left            =   9000
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox tbTotal 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox tbMACTo 
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox tbMACFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox tbSNTo 
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox tbSNFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox tbPart 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox tbWorkOrder 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   4200
      X2              =   4560
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   4560
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line2 
      X1              =   4200
      X2              =   4560
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4560
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Label Label5 
      Caption         =   "��������"
      Height          =   255
      Left            =   615
      TabIndex        =   6
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "MAC��Χ"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "SN��Χ"
      Height          =   255
      Left            =   615
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "��������"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "���䵥��"
      Height          =   255
      Left            =   615
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "formUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim sql As String
Dim sn As String

Dim myApp As New LabelManager2.Application
Dim myDoc As New LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects

Dim ConnFTPC As New ADODB.Connection
'Dim printArray() As String
Dim workorder_qty As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Dim Ctr As Control
    For Each Ctr In Me.Controls
        If TypeOf Ctr Is TextBox Then
            Ctr.Text = ""
        End If
    Next
End Sub

Private Sub cmdPrint_Click()
    Dim startString As String
    Dim qty As Integer
    Dim result As String

    If (Me.tbPart.Text = "ASR6026PF" Or Me.tbPart.Text = "ASR6126PF" Or Me.tbPart.Text = "ASR6026PF-DC" Or Me.tbPart.Text = "ASR6126PF-DC" Or Me.tbPart.Text = "ASR7024PF-AC" Or Me.tbPart.Text = "ASR7024EPF-AC" Or Me.tbPart.Text = "ASR8048PF-AC" Or Me.tbPart.Text = "ASR8048EPF-AC" Or Me.tbPart.Text = "ASR8024PF-AC" Or Me.tbPart.Text = "ASR8024EPF-AC" Or Me.tbPart.Text = "ASR7048PF-AC" Or Me.tbPart.Text = "ASR7048EPF-AC" Or Me.tbPart.Text = "ASR6226PF-AC" Or Me.tbPart.Text = "ASR6326PF-AC") Then
    Else
        MsgBox "���ֱ��벻��,������ȷ�Ϲ����������Ƿ���ȷ"
        tbWorkOrder.SetFocus
        Exit Sub
    End If
    
'    If Right$(Trim(tbMACFrom.Text), 1) <> "0" Then
'        MsgBox "MAC ĩλ������0"
 '       tbMACFrom.SetFocus
  '      Exit Sub
  '   End If
 '    If Trim(tbMACTo.Text) <> "" And Right$(Trim(tbMACTo.Text), 1) <> "0" Then
 '       MsgBox "MAC ĩλ������0"
 '       tbMACTo.SetFocus
   '     Exit Sub
'     End If
   ' If Trim(tbMACFrom2.Text) <> "" And Right$(Trim(tbMACFrom2.Text), 1) <> "0" Then
   '     MsgBox "MAC ĩλ������0"
   '    tbMACFrom2.SetFocus
   '     Exit Sub
   '  End If
  '   If Trim(tbMACTo2.Text) <> "" And Right$(Trim(tbMACTo2.Text), 1) <> "0" Then
  '      MsgBox "MAC ĩλ������0"
  '      tbMACTo2.SetFocus
  '      Exit Sub
 '    End If
 '   If Trim(tbMACFrom3.Text) <> "" And Right$(Trim(tbMACFrom3.Text), 1) <> "0" Then
'        MsgBox "MAC ĩλ������0"
  '      tbMACFrom3.SetFocus
 '       Exit Sub
 '    End If
  '   If Trim(tbMACTo3.Text) <> "" And Right$(Trim(tbMACTo3.Text), 1) <> "0" Then
 '       MsgBox "MAC ĩλ������0"
  '      tbMACTo3.SetFocus
  '      Exit Sub
  '   End If
    
     Dim secstr As String
     secstr = Mid$(Right$(Trim(tbMACFrom.Text), 2), 1, 1)
     If secstr <> "0" And secstr <> "2" And secstr <> "4" And secstr <> "6" And secstr <> "8" And secstr <> "A" And secstr <> "C" And secstr <> "E" Then
        MsgBox "MAC �����ڶ�λ����ȷ"
        tbMACFrom.SetFocus
        Exit Sub
     End If
'     secstr = Mid$(Right$(Trim(tbMACTo.Text), 2), 1, 1)
'     If Trim(tbMACTo.Text) <> "" And secstr <> "0" And secstr <> "2" And secstr <> "4" And secstr <> "6" And secstr <> "8" And secstr <> "A" And secstr <> "C" And secstr <> "E" Then
'        MsgBox "MAC �����ڶ�λ����ȷ"
'        tbMACTo.SetFocus
'        Exit Sub
'     End If
     
     secstr = Mid$(Right$(Trim(tbMACFrom2.Text), 2), 1, 1)
     If Trim(tbMACFrom2.Text) <> "" And secstr <> "0" And secstr <> "2" And secstr <> "4" And secstr <> "6" And secstr <> "8" And secstr <> "A" And secstr <> "C" And secstr <> "E" Then
        MsgBox "MAC �����ڶ�λ����ȷ"
        tbMACFrom2.SetFocus
        Exit Sub
     End If
     secstr = Mid$(Right$(Trim(tbMACTo2.Text), 2), 1, 1)
     If Trim(tbMACTo2.Text) <> "" And secstr <> "0" And secstr <> "2" And secstr <> "4" And secstr <> "6" And secstr <> "8" And secstr <> "A" And secstr <> "C" And secstr <> "E" Then
        MsgBox "MAC �����ڶ�λ����ȷ"
        tbMACTo2.SetFocus
        Exit Sub
     End If
     
     secstr = Mid$(Right$(Trim(tbMACFrom3.Text), 2), 1, 1)
     If Trim(tbMACFrom3.Text) <> "" And secstr <> "0" And secstr <> "2" And secstr <> "4" And secstr <> "6" And secstr <> "8" And secstr <> "A" And secstr <> "C" And secstr <> "E" Then
        MsgBox "MAC �����ڶ�λ����ȷ"
        tbMACFrom3.SetFocus
        Exit Sub
     End If
     secstr = Mid$(Right$(Trim(tbMACTo3.Text), 2), 1, 1)
     If Trim(tbMACTo3.Text) <> "" And secstr <> "0" And secstr <> "2" And secstr <> "4" And secstr <> "6" And secstr <> "8" And secstr <> "A" And secstr <> "C" And secstr <> "E" Then
        MsgBox "MAC �����ڶ�λ����ȷ"
        tbMACTo3.SetFocus
        Exit Sub
     End If
     
   
      If (Me.tbPart.Text = "ASR6026PF") Then
            startString = "F60A"
      ElseIf (Me.tbPart.Text = "ASR6126PF") Then
            startString = "F61A"
     
      ElseIf (Me.tbPart.Text = "ASR7024PF-AC") Then
            startString = "F74C"
        ElseIf (Me.tbPart.Text = "ASR7024EPF-AC") Then
            startString = "F74E"
              ElseIf (Me.tbPart.Text = "ASR8048PF-AC") Then
            startString = "F88C"
              ElseIf (Me.tbPart.Text = "ASR8048EPF-AC") Then
            startString = "F88E"
              ElseIf (Me.tbPart.Text = "ASR8024PF-AC") Then
            startString = "F84C"
              ElseIf (Me.tbPart.Text = "ASR8024EPF-AC") Then
            startString = "F84E"
              ElseIf (Me.tbPart.Text = "ASR7048PF-AC") Then
            startString = "F78C"
              ElseIf (Me.tbPart.Text = "ASR7048EPF-AC") Then
            startString = "F78E"
              ElseIf (Me.tbPart.Text = "ASR6226PF-AC") Then
            startString = "F62A"
              ElseIf (Me.tbPart.Text = "ASR6326PF-AC") Then
            startString = "F63A"
      End If
    qty = CInt(Me.tbTotal.Text)
    
    If Trim(tbSNFrom.Text) = "" Then
        MsgBox "��ʼ���벻��Ϊ��"
        tbSNFrom.SetFocus
        Exit Sub
    End If
    
    Dim intsn As Integer

    intsn = CInt(Right$(Trim(tbSNFrom.Text), 3))
    
    result = saveMaxSerial(startString, qty, intsn)
    
    
    'tbWorkOrder.SetFocus
      
    Dim arr2() As String
    ReDim Preserve arr2(qty, 2) As String
    Dim MACindex As Integer
    
    MACindex = 1
    arr2(0, 0) = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 3) + Right$("000" + CStr(CInt(Right$(Me.tbSNFrom.Text, 3)) + 0), 3)
    arr2(0, 1) = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + 0 * 32)), Len(Me.tbMACFrom.Text))

    'For i = 0 To qty - 1
    For i = 1 To qty - 1
        arr2(i, 0) = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 3) + Right$("000" + CStr(CInt(Right$(Me.tbSNFrom.Text, 3)) + i), 3)
        'arr2(i, 1) = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + i * 32)), Len(Me.tbMACFrom.Text))
        
        'Add by mike 2017-9-12 for multiple MAC range
        arr2(i, 1) = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(arr2(i - 1, 1))) + 32)), Len(arr2(i - 1, 1)))
        
        If arr2(i, 1) > UCase(tbMACTo.Text) And MACindex = 1 Then
            If Trim(tbMACFrom2.Text) = "" Then
                MsgBox "MAC������̫��!"
                Exit Sub
            End If
            MACindex = 2
            arr2(i, 1) = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom2.Text)))), Len(Me.tbMACFrom2.Text))
        ElseIf arr2(i, 1) > UCase(tbMACTo2.Text) And MACindex = 2 Then
            If Trim(tbMACFrom3.Text) = "" Then
                MsgBox "MAC������̫��!"
                Exit Sub
            End If
            MACindex = 3
            arr2(i, 1) = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom3.Text)))), Len(Me.tbMACFrom3.Text))
        ElseIf MACindex = 3 Then
            If arr2(i, 1) > UCase(tbMACTo3.Text) Then
                MsgBox "MAC������̫��!"
                Exit Sub
            End If
        End If
        
    Next

    OpenLppx


    For i = 0 To qty - 1
        myVars.Item("MAC").value = arr2(i, 1)
        myVars.Item("SN").value = arr2(i, 0)
        'ASR6026PF��ASR6126PF,ASR6026PF-DC,ASR6126PF-DC
        'modified by allen.yan by alex.sha's requirement
        If Me.tbPart.Text = "ASR6026PF" Then
            myVars.Item("Model").value = "ASR6026-AC"
        ElseIf Me.tbPart.Text = "ASR6126PF" Then
            myVars.Item("Model").value = "ASR6126-AC"
        ElseIf Me.tbPart.Text = "ASR6026PF-DC" Then
            myVars.Item("Model").value = "ASR6026-DC"
        ElseIf Me.tbPart.Text = "ASR6126PF-DC" Then
            myVars.Item("Model").value = "ASR6126-DC"
            
        ElseIf Me.tbPart.Text = "ASR7024PF-AC" Then
            myVars.Item("Model").value = "ASR7024PF-AC"
        ElseIf Me.tbPart.Text = "ASR7024EPF-AC" Then
            myVars.Item("Model").value = "ASR7024EPF-AC"
        ElseIf Me.tbPart.Text = "ASR8048PF-AC" Then
            myVars.Item("Model").value = "ASR8048PF-AC"
        ElseIf Me.tbPart.Text = "ASR8048EPF-AC" Then
            myVars.Item("Model").value = "ASR8048EPF-AC"
        ElseIf Me.tbPart.Text = "ASR8024PF-AC" Then
            myVars.Item("Model").value = "ASR8024PF-AC"
        ElseIf Me.tbPart.Text = "ASR8024EPF-AC" Then
            myVars.Item("Model").value = "ASR8024EPF-AC"
        ElseIf Me.tbPart.Text = "ASR7048PF-AC" Then
            myVars.Item("Model").value = "ASR7048PF-AC"
        ElseIf Me.tbPart.Text = "ASR7048EPF-AC" Then
            myVars.Item("Model").value = "ASR7048EPF-AC"
        ElseIf Me.tbPart.Text = "ASR6226PF-AC" Then
            myVars.Item("Model").value = "ASR6226PF-AC"
        ElseIf Me.tbPart.Text = "ASR6326PF-AC" Then
            myVars.Item("Model").value = "ASR6326-AC"
                  
        End If
        
        result = savePackFrontRecords(arr2(i, 0), arr2(i, 1), Trim(Me.txtWOHID.Text))
        
        myApp.Visible = False
        myDoc.PrintLabel 1
        myDoc.FormFeed
    Next

    tbWorkOrder.SetFocus
      
    UnloadLppx
      
'    Me.tbSNTo.Text = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 3) + Right$("000" + CStr(CInt(Right$(Me.tbSNFrom.Text, 3)) + CInt(tbTotal.Text) - 1), 3)
'    Me.tbMACTo.Text = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + CInt(tbTotal) - 1)), Len(Me.tbMACFrom.Text))

End Sub

Private Sub Form_Load()
    workorder_qty = 0
    txtWOHID.Text = ""
    If ConnFTPC.State = 0 Then
      ConnFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
      ConnFTPC.Open
   End If
   If conn.State = 0 Then
      conn.ConnectionString = getConnectionstring()
      conn.Open
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ConnFTPC.State = 1 Then
        ConnFTPC.Close
    End If
    If conn.State = 1 Then
        conn.Close
    End If
    UnloadLppx
End Sub

Private Sub tbTotal_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
     If CInt(tbTotal.Text) <= 0 Or CInt(tbTotal.Text) > 999 Then
           MsgBox "��������ȷ������(1~999)!", vbInformation + vbOKOnly, "��������"
             tbTotal.SetFocus
           Exit Sub
     End If
     If CInt(tbTotal.Text) > workorder_qty Then
        MsgBox "�������ܴ���" & workorder_qty, vbInformation + vbOKOnly, "��������"
        tbTotal.SetFocus
        Exit Sub
     End If
     
     If (Me.tbSNFrom.Text = "" Or Me.tbMACFrom.Text = "") Then
        MsgBox "����дMAC����SN��ˮ�����ֵ"
        tbMACFrom.SetFocus
        Exit Sub
     End If
     If (Len(tbSNFrom.Text) = 15 Or Len(tbMACFrom.Text) = 12) Then
     Else
        MsgBox "SN Ӧ����15λ,MAC��ַ12λ"
        tbMACFrom.SetFocus
        Exit Sub
     End If
     
  '   If Right$(Trim(tbMACFrom.Text), 1) <> "0" Then
  '      MsgBox "MAC ĩλ������0"
  '      tbMACFrom.SetFocus
        Exit Sub
  '   End If
     Dim secstr As String
     secstr = Mid$(Right$(Trim(tbMACFrom.Text), 2), 1, 1)
     If secstr <> "0" And secstr <> "2" And secstr <> "4" And secstr <> "6" And secstr <> "8" And secstr <> "A" And secstr <> "C" And secstr <> "E" Then
        MsgBox "MAC �����ڶ�λ����ȷ"
        tbMACFrom.SetFocus
        Exit Sub
     End If
     
     
     If (tbTotal = "1") Then
        tbSNTo.Text = ""
        tbMACTo.Text = ""
        Exit Sub
     End If
     
     
     Me.tbSNTo.Text = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 3) + Right$("000" + CStr(CInt(Right$(Me.tbSNFrom.Text, 3)) + CInt(tbTotal.Text) - 1), 3)
     'Me.tbMACTo.Text = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + CInt(tbTotal) * 32 - 32)), Len(Me.tbMACFrom.Text))
     
End If
End Sub


Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()

   Me.MousePointer = vbHourglass
   On Error GoTo errH
   
   If Me.tbPart.Text = "ASR6026PF" Or Me.tbPart.Text = "ASR6126PF" Or Me.tbPart.Text = "ASR7024PF-AC" Or Me.tbPart.Text = "ASR7024EPF-AC" Or Me.tbPart.Text = "ASR8048PF-AC" Or Me.tbPart.Text = "ASR8048EPF-AC" Or Me.tbPart.Text = "ASR8024PF-AC" Or Me.tbPart.Text = "ASR8024EPF-AC" Or Me.tbPart.Text = "ASR7048PF-AC" Or Me.tbPart.Text = "ASR7048EPF-AC" Or Me.tbPart.Text = "ASR6226PF-AC" Or Me.tbPart.Text = "ASR6326PF-AC" Then
        myApp.Visible = False
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\BU1С��Ŀ��ǩ\PFUnitlabel-AC.Lab")
   ElseIf Me.tbPart.Text = "ASR6026PF-DC" Or Me.tbPart.Text = "ASR6126PF-DC" Then
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\BU1С��Ŀ��ǩ\PFUnitlabel-DC.Lab")
   Else
        MsgBox "��Ʒ����û�ж�Ӧ��ģ��,���ܴ�ӡ!"
        Exit Sub
        Unload Me
   End If

   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
   
   Exit Sub
errH:
    MsgBox "�������������Ժ�����" + Err.Number, vbInformation + vbOKOnly, "����"

End Sub
    

Private Sub tbWorkOrder_KeyPress(KeyAscii As Integer)
    Dim tempWO As String '��ɨ�蹤���Ÿ�Ϊɨ���䵥��
    Dim detail As String
    tempWO = Me.tbWorkOrder.Text
    If (KeyAscii = 13) Then
        If Len(Trim(tbWorkOrder.Text)) <> 14 Then
            MsgBox "���볤�ȱ���Ϊ14λ!"
            tbWorkOrder.SetFocus
            Exit Sub
        Else
            sql = "SELECT  ID,ORDER_NUMBER,WO_QTY,MAC_START,MAC_END FROM [afg_active_90].[dbo].[MAC_RECORD] where ID='" & tempWO & "'"
            rec.Open sql, ConnFTPC, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                MsgBox "�˷��䵥��Ϣ������!"
                txtWO.Text = ""
                txtWO.SetFocus
                rec.Close
                Exit Sub
            Else
                Dim str_wo As String
                str_wo = Trim(rec.Fields(1))
                txtWOHID.Text = str_wo
                tbMACFrom.Text = Trim(rec.Fields(3))
                tbMACTo.Text = Trim(rec.Fields(4))
                workorder_qty = Trim(rec.Fields(2))
                
                'Add by mike 2017-9-12 for multiple MAC range
                rec.MoveNext
                If rec.EOF = True Then
                    
                Else
                    tbMACFrom2.Text = Trim(rec.Fields(3))
                    tbMACTo2.Text = Trim(rec.Fields(4))
                    rec.MoveNext
                    If rec.EOF = True Then
                        
                    Else
                        tbMACFrom3.Text = Trim(rec.Fields(3))
                        tbMACTo3.Text = Trim(rec.Fields(4))
                    End If
                End If
                
                rec.Close
                
                sql = "select part_revision,part_number,isnull(c.customer_po_number_S,'') from dbo.WORK_ORDER A,dbo.WORK_ORDER_ITEMS B, dbo.UDA_Order c WHERE A.order_key = B.order_key and b.order_key = c.object_key AND A.order_number ='" & str_wo & "'"
                rec.Open sql, ConnFTPC, adOpenKeyset, adLockReadOnly
                If rec.EOF = True Then
                    MsgBox "�˹�����Ϣ������!"
                    txtWO.Text = ""
                    txtWOHID.Text = ""
                    txtWO.SetFocus
                    rec.Close
                    Exit Sub
                Else
                    Me.tbPart.Text = Trim(rec.Fields(1))
                    Dim p_version As String
                    p_version = Trim(rec.Fields(0))
                    
                    rec.Close
                               If (Me.tbPart.Text = "ASR6026PF") Then
                        Me.tbSNFrom.Text = getNewStartSN("F60A", p_version)
                    ElseIf (Me.tbPart.Text = "ASR6126PF") Then
                         Me.tbSNFrom.Text = getNewStartSN("F61A", p_version)
                    ElseIf Me.tbPart.Text = "ASR6026PF-DC" Then
                        Me.tbSNFrom.Text = getNewStartSN("F60D", p_version)
                    ElseIf Me.tbPart.Text = "ASR6126PF-DC" Then
                        Me.tbSNFrom.Text = getNewStartSN("F61D", p_version)
                    ElseIf (Me.tbPart.Text = "ASR7024PF-AC") Then
                        Me.tbSNFrom.Text = getNewStartSN("F74CA", p_version)
                    ElseIf (Me.tbPart.Text = "ASR6326PF-AC") Then
                        Me.tbSNFrom.Text = getNewStartSN_6("F63A", p_version)
                '    ElseIf (Me.tbPart.Text = "GP1126PF-DC") Then
                '         Me.tbSNFrom.Text = getNewStartSN("F61D", p_version)
                 '   ElseIf (Me.tbPart.Text = "GP1126PF-DC") Then
                 '        Me.tbSNFrom.Text = getNewStartSN("F61D", p_version)
                 '   ElseIf (Me.tbPart.Text = "GP1126PF-DC") Then
                 '        Me.tbSNFrom.Text = getNewStartSN("F61D", p_version)
                 '   ElseIf (Me.tbPart.Text = "GP1126PF-DC") Then
                 '        Me.tbSNFrom.Text = getNewStartSN("F61D", p_version)
                  '  ElseIf (Me.tbPart.Text = "GP1126PF-DC") Then
                  '       Me.tbSNFrom.Text = getNewStartSN("F61D", p_version)
                  '  ElseIf (Me.tbPart.Text = "GP1126PF-DC") Then
                  '       Me.tbSNFrom.Text = getNewStartSN("F61D", p_version)
                  '  ElseIf (Me.tbPart.Text = "GP1126PF-DC") Then
                  '       Me.tbSNFrom.Text = getNewStartSN("F61D", p_version)
                  '  ElseIf (Me.tbPart.Text = "GP1126PF-DC") Then
                  '       Me.tbSNFrom.Text = getNewStartSN("F61D", p_version)
                   End If
               End If
           End If
        End If
    End If
End Sub

