VERSION 5.00
Begin VB.Form frmMAC 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10605
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox tbMACTo3 
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox tbMACFrom3 
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox tbMACTo2 
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox tbMACFrom2 
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox txtWOHID 
      Height          =   270
      Left            =   6480
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.TextBox tbWorkOrder 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox tbPart 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox tbSNFrom 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox tbSNTo 
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox tbMACFrom 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox tbMACTo 
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox tbTotal 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   4980
      TabIndex        =   1
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   4680
      X2              =   5040
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   5040
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label1 
      Caption         =   "���䵥��"
      Height          =   255
      Left            =   1095
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "��������"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "SN��Χ"
      Height          =   255
      Left            =   1095
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   5040
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Label Label4 
      Caption         =   "MAC��Χ"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   4680
      X2              =   5040
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label5 
      Caption         =   "��������"
      Height          =   255
      Left            =   1095
      TabIndex        =   10
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1830
      Left            =   0
      Picture         =   "frmMAC.frx":0000
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmMAC"
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

Option Explicit

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
    
    Dim tempWO As String '��ɨ�蹤���Ÿ�Ϊɨ���䵥��
    tempWO = Me.tbWorkOrder.Text

    If (Me.tbPart.Text = "GP-MPC480-MB" Or Me.tbPart.Text = "MPC480-MB4818") Then
    Else
        MsgBox "���ֱ��벻��,������ȷ�Ϲ����������Ƿ���ȷ"
        tbWorkOrder.SetFocus
        Exit Sub
    End If
    
    If Right$(Trim(tbMACFrom.Text), 1) <> "0" Then
        MsgBox "MAC ĩλ������0"
        tbMACFrom.SetFocus
        Exit Sub
     End If
     If Trim(tbMACTo.Text) <> "" And Right$(Trim(tbMACTo.Text), 1) <> "0" Then
        MsgBox "MAC ĩλ������0"
        tbMACTo.SetFocus
        Exit Sub
     End If
     
    If Trim(tbMACFrom2.Text) <> "" And Right$(Trim(tbMACFrom2.Text), 1) <> "0" Then
        MsgBox "MAC ĩλ������0"
        tbMACFrom2.SetFocus
        Exit Sub
     End If
     If Trim(tbMACTo2.Text) <> "" And Right$(Trim(tbMACTo2.Text), 1) <> "0" Then
        MsgBox "MAC ĩλ������0"
        tbMACTo2.SetFocus
        Exit Sub
     End If
     
    If Trim(tbMACFrom3.Text) <> "" And Right$(Trim(tbMACFrom3.Text), 1) <> "0" Then
        MsgBox "MAC ĩλ������0"
        tbMACFrom3.SetFocus
        Exit Sub
     End If
     If Trim(tbMACTo3.Text) <> "" And Right$(Trim(tbMACTo3.Text), 1) <> "0" Then
        MsgBox "MAC ĩλ������0"
        tbMACTo3.SetFocus
        Exit Sub
     End If
     
     Dim secstr As String
     secstr = Mid$(Right$(Trim(tbMACFrom.Text), 2), 1, 1)
     If secstr <> "0" And secstr <> "2" And secstr <> "4" And secstr <> "6" And secstr <> "8" And secstr <> "A" And secstr <> "C" And secstr <> "E" Then
        MsgBox "MAC �����ڶ�λ����ȷ"
        tbMACFrom.SetFocus
        Exit Sub
     End If
     secstr = Mid$(Right$(Trim(tbMACTo.Text), 2), 1, 1)
     If Trim(tbMACTo.Text) <> "" And secstr <> "0" And secstr <> "2" And secstr <> "4" And secstr <> "6" And secstr <> "8" And secstr <> "A" And secstr <> "C" And secstr <> "E" Then
        MsgBox "MAC �����ڶ�λ����ȷ"
        tbMACTo.SetFocus
        Exit Sub
     End If
     
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
     
    If (Me.tbPart.Text = "GP-MPC480-MB" Or Me.tbPart.Text = "MPC480-MB4818") Then
          startString = "F846"
    End If
      
    qty = CInt(Me.tbTotal.Text)
    
    If Trim(tbSNFrom.Text) = "" Then
        MsgBox "��ʼ���벻��Ϊ��"
        tbSNFrom.SetFocus
        Exit Sub
    End If
    
    Dim intsn As Integer

    intsn = CInt(HEXTODEC(Right$(Trim(tbSNFrom.Text), 2)))
    
    result = saveMaxSerialF846(startString, qty, intsn)
    
    
    'tbWorkOrder.SetFocus
    Dim i As Integer
    Dim arr2() As String
    ReDim Preserve arr2(qty, 2) As String
    Dim MACindex As Integer
    
    MACindex = 1
    arr2(0, 0) = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 2) + Right$("00" + dectohex(CInt(HEXTODEC(Right$(Me.tbSNFrom.Text, 2))) + 0), 2)
    arr2(0, 1) = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + 0 * 128)), Len(Me.tbMACFrom.Text))
    
    'For i = 0 To qty - 1
    For i = 1 To qty - 1
        arr2(i, 0) = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 2) + Right$("00" + dectohex(CInt(HEXTODEC(Right$(Me.tbSNFrom.Text, 2))) + i), 2)
        'arr2(i, 1) = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + i * 128)), Len(Me.tbMACFrom.Text))
    
        'Add by mike 2017-9-12 for multiple MAC range
        arr2(i, 1) = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(arr2(i - 1, 1))) + 128)), Len(arr2(i - 1, 1)))
        
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
        'myVars.Item("SN").value = arr2(i, 0)
        
        result = savePackFrontRecords(arr2(i, 0), arr2(i, 1), Trim(Me.txtWOHID.Text))
        
        myApp.Visible = False
        myDoc.PrintLabel 1
        myDoc.FormFeed
        
        
    Next
    

    
    tbWorkOrder.SetFocus
    
    Dim ssn As String
    ssn = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 2) + Right$("00" + dectohex(CInt(HEXTODEC(Right$(Me.tbSNFrom.Text, 2))) + 0), 2)
    Dim smac As String
    smac = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + 0 * 128)), Len(Me.tbMACFrom.Text))
    result = savePackFrontRecordsF846(ssn, smac, Trim(Me.txtWOHID.Text), tempWO, qty)
      
    UnloadLppx
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
     If (Len(tbSNFrom.Text) = 11 And Len(tbMACFrom.Text) = 12) Then
     Else
        MsgBox "SN Ӧ����11λ,MAC��ַ12λ"
        tbMACFrom.SetFocus
        Exit Sub
     End If
     
     If Right$(Trim(tbMACFrom.Text), 1) <> "0" Then
        MsgBox "MAC ĩλ������0"
        tbMACFrom.SetFocus
        Exit Sub
     End If
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
     
     
     'Me.tbSNTo.Text = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 3) + Right$("000" + CStr(CInt(Right$(Me.tbSNFrom.Text, 3)) + CInt(tbTotal.Text) - 1), 3)
     'Me.tbMACTo.Text = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + CInt(tbTotal) * 32 - 32)), Len(Me.tbMACFrom.Text))
     
     Me.tbSNTo.Text = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 2) + Right$("00" + dectohex(CInt(HEXTODEC(Right$(Me.tbSNFrom.Text, 2))) + CInt(tbTotal.Text) - 1), 2)
     'Me.tbMACTo.Text = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + CInt(tbTotal) * 128 - 128)), Len(Me.tbMACFrom.Text))
     
     
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
   
   If Me.tbPart.Text = "GP-MPC480-MB" Or Me.tbPart.Text = "MPC480-MB4818" Then
        myApp.Visible = False
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\BU1С��Ŀ��ǩ\GP-MPC480-MB-MAC.lab")
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
                tbWorkOrder.SetFocus
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
                    txtWOHID.Text = ""
                    tbWorkOrder.SetFocus
                    rec.Close
                    Exit Sub
                Else
                    Me.tbPart.Text = Trim(rec.Fields(1))
                    rec.Close
                    '=================================
                    Dim myF As Boolean
                    sql = "select top 1 [fenpeihao],[wo],[sn],[mac],[qty],[lastprinttime] from [Print].[dbo].[PacketFrontRecordF846] where fenpeihao='" & tempWO & "' and wo='" & str_wo & "' order by lastprinttime desc"
                    Dim beginsn As String
                    Dim qty As Integer
                    myF = getPrintedhistoryData(sql, beginsn, qty)

                    If myF = False Then

                        sql = "SELECT Material_part_rev FROM [afg_active_90].[dbo].[WORK_ORDER_BOM_LOCATIONS] where order_number='" & str_wo & "' and ISNULL(sub_order_number,'')<>'' "
                        
                        'sql = "select part_revision from dbo.WORK_ORDER A,dbo.WORK_ORDER_ITEMS B, dbo.UDA_Order c WHERE A.order_key = B.order_key and b.order_key = c.object_key AND A.order_number ='" & str_wo & "'"
                        
                        rec.Open sql, ConnFTPC, adOpenKeyset, adLockReadOnly
                        If rec.EOF = True Then
                            MsgBox "1�׹����汾������!"
                            txtWOHID.Text = ""
                            tbWorkOrder.SetFocus
                            rec.Close
                            Exit Sub
                        Else
                            Dim p_version As String
                            p_version = Trim(rec.Fields(0))
                            rec.Close

                            If (Me.tbPart.Text = "GP-MPC480-MB" Or Me.tbPart.Text = "MPC480-MB4818") Then
                                Me.tbSNFrom.Text = getStartSNF846("F846", p_version)
                            End If
                        End If
                    Else
                        Me.tbSNFrom.Text = beginsn
                        Me.tbTotal.Text = qty
                        tbTotal_KeyPress 13
                    End If
                    '=================================
                    
                End If
                
            End If
            
        End If
    End If
End Sub




