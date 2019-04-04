VERSION 5.00
Begin VB.Form formUnit_128 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   Picture         =   "formUnit_128.frx":0000
   ScaleHeight     =   7470
   ScaleWidth      =   10605
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox tbVer 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
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
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox tbMACTo 
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox tbMACFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox tbSNTo 
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox tbSNFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox tbPart 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   3960
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
      Caption         =   "打印"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "机种版本"
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   4200
      X2              =   4560
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4560
      Y1              =   4740
      Y2              =   4740
   End
   Begin VB.Label Label5 
      Caption         =   "工单数量"
      Height          =   255
      Left            =   615
      TabIndex        =   6
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "MAC范围"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "SN范围"
      Height          =   255
      Left            =   615
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "机种名称"
      Height          =   255
      Left            =   615
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "工单号"
      Height          =   255
      Left            =   615
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "formUnit_128"
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
Dim ConnFTPC As New ADODB.Connection
'Dim printArray() As String

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
    
    Dim dictionary As New dictionary
    Dim File As String
    File = "\\sz-fs01\Labels\Packet Front\ASSY SN label\0846PFNP-R BB SN LABEL.btw"

    If (Me.tbPart.Text = "0846PFNP-R") Then
    Else
        MsgBox "机种编码不对,请重新确认工单号输入是否正确"
        tbWorkOrder.SetFocus
        Exit Sub
    End If
   
    If (Me.tbPart.Text = "0846PFNP-R") Then
          startString = "F846"
    End If
    qty = CInt(Me.tbTotal.Text)
    
    If Trim(tbSNFrom.Text) = "" Then
        MsgBox "起始条码不可为空"
        tbSNFrom.SetFocus
        Exit Sub
    End If
    
    Dim intsn As Integer
    intsn = CInt(Right$(Trim(tbSNFrom.Text), 3))
    
    result = saveMaxSerial(startString, qty, intsn)
    
    Dim arr2() As String
    ReDim Preserve arr2(qty, 2) As String
    For i = 0 To qty - 1
        arr2(i, 0) = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 2) + Right$("00" + dectohex(CInt(HEXTODEC(Right$(Me.tbSNFrom.Text, 2))) + i), 2)
        arr2(i, 1) = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + i * 128)), Len(Me.tbMACFrom.Text))
    Next
    
    For i = 0 To qty - 1

        Call dictionary.PutValue("MAC", arr2(i, 1))
        Call dictionary.PutValue("SN", arr2(i, 0))
        
        Call Printer.PrintLabel(File, dictionary)
        
        result = savePackFrontRecords(arr2(i, 0), arr2(i, 1), Trim(Me.tbWorkOrder.Text))
        
    Next
    
    
    MsgBox "打印完成！"
    
    tbWorkOrder.Text = ""
    tbPart.Text = ""
    tbVer.Text = ""
    tbSNFrom.Text = ""
    tbSNTo.Text = ""
    tbMACFrom.Text = ""
    tbMACTo.Text = ""
    tbTotal.Text = ""
    

End Sub

Private Sub Command1_Click()
Dim a As New dictionary

Call a.PutValue("key1", "value1")
Call a.PutValue("key2", "value2")
Call a.PutValue("key1", "value11")
Call a.PutValue("MAC", "value11")
Debug.Print ("key1 = " & a.GetValue("key1"))
Debug.Print ("key2 = " & a.GetValue("key2"))


Dim File As String

File = "D:\WORK\template\1.btw"
Call Printer.PrintLabel(File, a)


End Sub

Private Sub Form_Load()
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
           MsgBox "请输入正确的数量(1~999)!", vbInformation + vbOKOnly, "数量不对"
             tbTotal.SetFocus
           Exit Sub
     End If
     If (Me.tbSNFrom.Text = "" Or Me.tbMACFrom.Text = "") Then
        MsgBox "请填写MAC或者SN流水号相关值"
        tbMACFrom.SetFocus
        Exit Sub
     End If
     If (Len(tbSNFrom.Text) = 11 And Len(tbMACFrom.Text) = 12) Then
     Else
        MsgBox "SN 应该是11位,MAC地址12位"
        tbMACFrom.SetFocus
        Exit Sub
     End If
     
     
     If (tbTotal = "1") Then
        tbSNTo.Text = ""
        tbMACTo.Text = ""
        Exit Sub
     End If
     
     
     Me.tbSNTo.Text = Mid$(Me.tbSNFrom.Text, 1, Len(Me.tbSNFrom.Text) - 2) + Right$("00" + dectohex(CInt(HEXTODEC(Right$(Me.tbSNFrom.Text, 2))) + CInt(tbTotal.Text) - 1), 2)
     Me.tbMACTo.Text = Right$("000" + dectohex(CStr(CDbl(HEXTODEC(tbMACFrom.Text)) + CInt(tbTotal) * 128 - 128)), Len(Me.tbMACFrom.Text))
     
End If
End Sub


Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   If Me.tbPart.Text = "ASR6026PF" Or Me.tbPart.Text = "ASR6126PF" Then
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\BU1小项目标签\" & "PFUnitlabel-AC.Lab")
   ElseIf Me.tbPart.Text = "ASR6026PF-DC" Or Me.tbPart.Text = "ASR6126PF-DC" Then
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\BU1小项目标签\" & "PFUnitlabel-DC.Lab")
   Else
        MsgBox "产品编码没有对应的模板,不能打印!"
        Exit Sub
        Unload Me
   End If

   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub
    

Private Sub tbWorkOrder_KeyPress(KeyAscii As Integer)
    Dim tempWO As String
    Dim detail As String
    tempWO = Me.tbWorkOrder.Text
    If (KeyAscii = 13) Then
        If Len(Trim(tbWorkOrder.Text)) <> 6 And Len(Trim(tbWorkOrder.Text)) <> 8 Then
            MsgBox "工单号长度必须为6位!"
            tbWorkOrder.SetFocus
            Exit Sub
        Else
            sql = "select part_revision,part_number,isnull(c.customer_po_number_S,'') from dbo.WORK_ORDER A,dbo.WORK_ORDER_ITEMS B, dbo.UDA_Order c WHERE A.order_key = B.order_key and b.order_key = c.object_key AND A.order_number ='" & tempWO & "'"
            rec.Open sql, ConnFTPC, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                MsgBox "此工单信息不存在!"
                txtWO.Text = ""
                txtWO.SetFocus
                rec.Close
                Exit Sub
            Else
                Me.tbPart.Text = Trim(rec.Fields(1))
                Me.tbVer.Text = Trim(rec.Fields(0))
                rec.Close
                If (Me.tbPart.Text = "0846PFNP-R") Then
                    Me.tbSNFrom.Text = getStartSN128("F846", Me.tbVer.Text)
                End If
            End If
        End If
    End If
End Sub

