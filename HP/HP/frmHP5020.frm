VERSION 5.00
Begin VB.Form frmHP5020 
   Caption         =   "HP5020"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   8430
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chkN 
      Caption         =   "N*"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   24
      Top             =   5640
      Width           =   615
   End
   Begin VB.CheckBox chkN4 
      Caption         =   "N4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   23
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtModel 
      Height          =   285
      Left            =   7320
      TabIndex        =   22
      Top             =   7440
      Width           =   495
   End
   Begin VB.CheckBox chkY2 
      Caption         =   "Y2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   5640
      Width           =   615
   End
   Begin VB.CheckBox chkY 
      Caption         =   "Y*"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtRevision 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtPart 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txtWorkOrder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtSN 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox txtProduct 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox txtDesc1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox txtDesc2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CommandButton cmdPrint_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "�� ӡ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "ȡ ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdReturn_HPSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdMPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "������ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2100
      Left            =   1680
      Picture         =   "frmHP5020.frx":0000
      ScaleHeight     =   2040
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000004&
      Caption         =   "��������:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "�汾:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   5040
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8400
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ���к�:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ���:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ����1:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ����2:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
End
Attribute VB_Name = "frmHP5020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New Recordset
Dim bom_code As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim rs As New Recordset
Dim newLableFlag As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Sub cmdCancel_HPSN_Click()
    Me.txtSN.Text = ""
    Me.txtProduct.Text = ""
    Me.txtDesc1.Text = ""
    Me.txtDesc2.Text = ""
    Me.txtWorkOrder.Text = ""
    Me.txtPart.Text = ""
    Me.txtRevision.Text = ""
    Me.chkY.Value = 0
    Me.chkY2.Value = 0
    Me.chkN.Value = 0
    Me.chkN4.Value = 0
End Sub

Private Sub cmdMPrint_Click()
    Dim model As String
    If Me.txtPart.Text = "" And txtModel.Text = "" Then
        MsgBox "�û�����Ϣ���ܴ�ӡ!"
        Exit Sub
    ElseIf Me.txtPart.Text <> "" Then
        model = Mid(txtPart.Text, 4, 8)
    ElseIf txtModel.Text <> "" Then
        model = Trim(txtModel.Text)
    End If
    

 If Connect.checkPrintPreCondition(model, 3) = False Then
    MsgBox "�û���û����HP���к�����ά��Ϊ[50*20]��ӡѡ��!"
    Exit Sub
End If
 cmdReturn_HPSN.Enabled = False
'cmdPrint_HPSN.Enabled = False
cmdCancel_HPSN.Enabled = False
sql = "select ITEM_CODE,BARCODE from tblHP_Print where isnull(BARCODE,'')<>'' and isnull(ITEM_CODE,'')<>'' order by BARCODE"
If conn1.State = 0 Then
    conn1.Open
End If
rs.Open sql, conn1, adOpenStatic, adLockReadOnly
If rs.EOF = True Then
    MsgBox ("���к�δ���룡")
    rs.Close
    cmdReturn_HPSN.Enabled = True
    'cmdPrint_HPSN.Enabled = True
    cmdCancel_HPSN.Enabled = True
    Exit Sub
Else
    For i = 1 To rs.RecordCount
    
        txtSN.Text = rs("BARCODE")
        txtModel.Text = rs("ITEM_CODE")
        'begin
        If Len(txtSN.Text) < 10 Then
            MsgBox "��Ʒ��ų��Ȳ���С��10!"
            txtSN.SetFocus
            Exit Sub
        End If
        If InStr(1, txtPart.Text, txtModel.Text) <= 0 Then
            MsgBox ("�ù����Ϻź������Ӧ���ϺŲ�һ�£���ȷ�����빤���Ƿ���ȷ!")
            rs.Close
            Exit Sub
        End If
        updateHPInformation
        cmdPrint_HPSN_Click
        rs.MoveNext
        If i Mod 100 = 0 Then
            Sleep (1000 * 10)
       End If
    Next
    UnloadLppx
    cmdCancel_HPSN_Click
    rs.Close
End If
'del_excel
del_sql
cmdReturn_HPSN.Enabled = True
'cmdPrint_HPSN.Enabled = True
cmdCancel_HPSN.Enabled = True
'MsgBox ("������ӡ�ɹ���")
End Sub

Private Sub cmdPrint_HPSN_Click()
    Dim Pb As String
    
    If txtSN.Text = "" Then
        MsgBox ("���к�δ���룬���ܴ�ӡ��")
        txtSN.SetFocus
        Exit Sub
    End If
    If txtProduct.Text = "" Then
        MsgBox ("��Ʒ����δ���������ܴ�ӡ��")
        Exit Sub
    End If
    If txtDesc1.Text = "" Then
        MsgBox ("��Ʒ����1δ���������ܴ�ӡ��")
        Exit Sub
    End If
     If txtModel.Text = "" Then
        MsgBox ("����������ITEM_CODE������Ϊ�գ�")
        Exit Sub
    End If
    OpenLppx
    myVars.Item("SN").Value = UCase(txtSN.Text)
    myVars.Item("PN").Value = UCase(txtProduct.Text)
    myVars.Item("Model").Value = UCase(txtModel.Text)
    myVars.Item("Rev").Value = UCase(txtRevision.Text)
    
    If (Me.chkY.Value = 1) Then
        Pb = "Y*"
        OverridePb Pb
        myVars.Item("Rohs").Value = Pb
    ElseIf (Me.chkY2.Value = 1) Then
        Pb = "Y2"
        OverridePb Pb
        myVars.Item("Rohs").Value = Pb
    ElseIf (Me.chkN.Value = 1) Then
        Pb = "N*"
        OverridePb Pb
        myVars.Item("Rohs").Value = Pb
    ElseIf (Me.chkN4.Value = 1) Then
        Pb = "N4"
        OverridePb Pb
        myVars.Item("Rohs").Value = Pb
    Else
        MsgBox "��������δѡ�񣬲��ܴ�ӡ!"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    
    If UploadH3CInfo(Pb, Trim(UCase(txtSN.Text)), Trim(UCase(txtRevision.Text)), "NA", "N/A", "CHINA", "frmHP5020") = False Then
        MsgBox "���ϱ���ʧ�ܲ��ܴ�ӡ!", vbInformation + vbOKOnly, "���ϱ���ʧ��"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
    If UploadH3C_PB(Pb, Trim(UCase(txtSN.Text)), Trim(UCase(txtRevision.Text)), "NA", "N/A", "CHINA", "frmHP5020") = False Then
        MsgBox "PB���ϱ���ʧ�ܲ��ܴ�ӡ!", vbInformation + vbOKOnly, "���ϱ���ʧ��"
        txtSN.SetFocus
        UnloadLppx
        Exit Sub
    End If
    '======Add by mike 2015.3.24 for data upload to FTPC============
    
    If txtDesc2.Text <> "" Then
        myVars.Item("ID-1").Value = txtDesc1.Text
        myVars.Item("ID-2").Value = txtDesc2.Text
    Else
        myVars.Item("ID-1").Value = txtDesc1.Text
        myVars.Item("ID-2").Value = ""
    End If
    'OpenLppx
    myDoc.PrintLabel 1
    myDoc.FormFeed
End Sub

Private Sub OverridePb(ByRef Pb As String)
    Dim labelHistory As New Label_History
    Dim sn As String
    sn = txtSN.Text
    If labelHistory.Init(sn) Then
        Pb = labelHistory.Pb
    End If
End Sub

Private Sub cmdReturn_HPSN_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    If conn1.State = 0 Then
      conn1.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
      conn1.Open
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
   
    If conn1.State = 1 Then
      conn1.Close
      Set conn1 = Nothing
   End If
   
   If connFTPC.State = 1 Then
        connFTPC.Close
        Set connFTPC = Nothing
   End If
End Sub

Private Sub updateHPInformation()

      sql = "select * from hp where hp_sn_iii=substring('" & Trim(txtSN.Text) & "',5,3) and h3c_bom_code = '" + txtModel.Text + "'"
      If rec.State = 1 Then
        rec.Close
      End If
      
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      
      If rec.EOF = True Then
          MsgBox "�����к�δά����Ϣ!"
          txtSN.Text = ""
          txtSN.SetFocus
          rec.Close
          Exit Sub
      Else
          If IsNull(rec.Fields("hpsnproduct")) Then
              MsgBox ("�����к�δά����Ʒ����!")
              rec.Close
              Exit Sub
          Else
              txtProduct = rec.Fields("hpsnproduct")
          End If
'          hpsnproduct
    
          If IsNull(rec.Fields("hp_desc1")) Then
              MsgBox ("�����к�δά��������Ϣ!")
              rec.Close
              Exit Sub
          Else
              txtDesc1 = rec.Fields("hp_desc1")
          End If
    
          If Not IsNull(rec.Fields("hp_desc2")) Then
              txtDesc2 = rec.Fields("hp_desc2")
          End If
          
          If IsNull(rec.Fields("new_label")) Or Trim(rec.Fields("new_label")) = "" Then
              newLableFlag = False
              MsgBox ("�˻������к�û��ά��new_labelѡ�����ϵME!")
              rec.Close
              Exit Sub
          Else
              newLableFlag = True
          End If
      End If
    
End Sub
'Private Sub txtSN_KeyPress(KeyAscii As Integer)
'
'End Sub



'Private Sub txtSN_Change()
'
'End Sub

Private Sub txtWorkOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPart.Text = ""
        Me.txtRevision.Text = ""
        Me.chkY.Value = 0
        Me.chkY2.Value = 0
        Me.chkN.Value = 0
        Me.chkN4.Value = 0
        If Trim(Me.txtWorkOrder.Text) <> "" Then
            If rec.State = 1 Then
                rec.Close
            End If
           
            If connFTPC.State = 0 Then
                connFTPC.Open
            End If
            Dim tempWO As String
         
            tempWO = Trim(Me.txtWorkOrder.Text)
            sql = "select b.part_number,b.part_revision from WORK_ORDER a,WORK_ORDER_ITEMS b where a.order_key = b.order_key and a.order_number = '" & tempWO & "'"
            rec.Open sql, connFTPC, adOpenKeyset, adLockReadOnly
            If rec.EOF = True Then
                MsgBox ("�ù��������ڣ���ȷ�����빤���Ƿ���ȷ!")
                rec.Close
                Exit Sub
            Else
'
                Me.txtPart.Text = rec.Fields("part_number")
                Me.txtRevision.Text = rec.Fields("part_revision")
                
                
                If Connect.getPartList(Trim(Me.txtWorkOrder.Text)) <> "" Then
                    lPB = Connect.GetPBState(Connect.getPartList(Trim(Me.txtWorkOrder.Text)))
                    If (lPB = "NPb") Then
                        Me.chkY2.Value = 1
                        Me.chkY.Value = 0
                        Me.chkN.Value = 0
                        Me.chkN4.Value = 0
                    ElseIf (lPB = "N*") Then
                        Me.chkY2.Value = 0
                        Me.chkY.Value = 0
                        Me.chkN.Value = 1
                        Me.chkN4.Value = 0
                    ElseIf (lPB = "N4") Then
                        Me.chkY2.Value = 0
                        Me.chkY.Value = 0
                        Me.chkN.Value = 0
                        Me.chkN4.Value = 1
                    Else
                        Me.chkY.Value = 0
                        Me.chkY2.Value = 0
                        Me.chkN.Value = 0
                        Me.chkN4.Value = 0
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub
Private Sub OpenLppx()
    
        If c = True Then
            Me.MousePointer = vbHourglass
            Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\" & "HPE ��SN��ǩ5020.lab")
        Else
            Me.MousePointer = vbHourglass
            
            Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\HP�����ǩ����\" & "HP SN��ǩ5020.lab")
        End If
    
'    Me.MousePointer = vbHourglass
'    Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\HP�����ǩ����\" & "HP SN��ǩ5020.lab")
'    If txtDesc2.Text = "" Then
'        Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HP���кű�ǩС��30λ����.lab")
'    Else
'        Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HP���кű�ǩ����30λ����.lab")
'    End If
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub

Sub del_sql()
    Dim delsql As String
    delsql = "delete from tblHP_Print"
    conn1.Execute delsql
End Sub

