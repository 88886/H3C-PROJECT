VERSION 5.00
Begin VB.Form FormHPFahuoSNX 
   BackColor       =   &H80000005&
   Caption         =   "HP������ǩ(��)"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10140
   ScaleWidth      =   7215
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtImei1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      TabIndex        =   20
      Top             =   7800
      Width           =   2895
   End
   Begin VB.TextBox txtImei2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      TabIndex        =   19
      Top             =   8640
      Width           =   2895
   End
   Begin VB.TextBox txtMac 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      TabIndex        =   17
      Top             =   7080
      Width           =   2895
   End
   Begin VB.ComboBox txtSZ 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "FormHPFahuoSNX.frx":0000
      Left            =   3000
      List            =   "FormHPFahuoSNX.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtModel 
      Height          =   270
      Left            =   6720
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   180
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
      TabIndex        =   14
      Top             =   9480
      Width           =   1215
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
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   9480
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
      TabIndex        =   12
      Top             =   9480
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
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox txtDesc 
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
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtUPC 
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
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5640
      Width           =   2895
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
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox txtPN 
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
      Height          =   525
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox txtSN 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      TabIndex        =   0
      Top             =   3240
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   3015
      Left            =   0
      Picture         =   "FormHPFahuoSNX.frx":0014
      ScaleHeight     =   2955
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000004&
      Caption         =   "IMEI1��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   22
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "IMEI2��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "MAC��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      Caption         =   "��ƷUPC��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ��ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ���֣�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ���кţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
End
Attribute VB_Name = "FormHPFahuoSNX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects


Private Sub cmdMPrint_Click()
cmdReturn_HPSN.Enabled = False
'cmdPrint_HPSN.Enabled = False
cmdCancel_HPSN.Enabled = False
'sql = "select sn from hp_print where isnull(sn,'')<>'' order by sn"
sql = "select ITEM_CODE,BARCODE from tblHP_Print where isnull(BARCODE,'')<>'' and isnull(ITEM_CODE,'')<>'' order by BARCODE"
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
        txtSN = rs("BARCODE")
        txtModel = rs("ITEM_CODE")
    'begin
            If Len(txtSN.Text) < 10 Then
            MsgBox "��Ʒ��ų��Ȳ���С��10!"
            txtSN.SetFocus
            Exit Sub
        End If
sql = "select * from hp where hp_sn_iii=substring('" & Trim(txtSN.Text) & "',5,3)  and h3c_bom_code='" & Trim(txtModel.Text) & "'"
        'MsgBox (sql)
        rec.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rec.EOF = True Then
            MsgBox "�����к�δά����Ϣ!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
        Else
           If rec.Fields("new_label") = "N" Then
                 MsgBox "�˱����ھ�ģ���ӡ!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
            End If
        
        
           If IsNull(rec.Fields("new_label")) Then
                 MsgBox "�˱����ھ�ģ���ӡ!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
            End If
        
        
            If IsNull(rec.Fields("hp_pn")) Then
               ' MsgBox ("�����к�δά������!")
              '  rec.Close
               ' Exit Sub
               txtPN = ""
            Else
                txtPN = rec.Fields("hp_pn")
            End If
             
            If IsNull(rec.Fields("hp_gtin_number")) Then
                'MsgBox ("�����к�δά��UPC!")
                'rec.Close
                'Exit Sub
                txtUPC = ""
            Else
                txtUPC = rec.Fields("hp_gtin_number")
            End If
        
            If IsNull(rec.Fields("hp_product")) Then
                'MsgBox ("�����к�δά����Ʒ����!")
                'rec.Close
                'Exit Sub
            Else
                txtProduct = rec.Fields("hp_product")
            End If
      
            If IsNull(rec.Fields("hp_desc1")) Then
                MsgBox ("�����к�δά��������Ϣ!")
                rec.Close
                Exit Sub
            Else
                txtDesc = rec.Fields("hp_desc1")
            End If
            
            If Not IsNull(rec.Fields("hp_desc2")) Then
                txtDesc = txtDesc & " " & rec.Fields("hp_desc2")
            End If
            
          
            
            
            
            rec.Close
            cmdPrint_HPSN_Click
        End If
    'end
    rs.MoveNext
    Next
    rs.Close
End If
del_excel
del_sql
cmdReturn_HPSN.Enabled = True
'cmdPrint_HPSN.Enabled = True
cmdCancel_HPSN.Enabled = True
End Sub
Sub del_sql()
    Dim delsql As String
    delsql = "delete from tblHP_Print"
    conn1.Execute delsql
End Sub
Sub del_excel()
          Dim tempxlApp     As New Excel.Application
          Dim tempxlWorkbook     As New Excel.Workbook
          Dim tempxlSheet     As New Excel.Worksheet
          Set tempxlWorkbook = tempxlApp.Workbooks.Open(App.Path & "\import.xls")
          'tempxlApp.DisplayAlerts = False
          Set tempxlSheet = tempxlWorkbook.Worksheets("Sheet1")
          tempxlSheet.Select
          tempxlSheet.Cells.Select
          Selection.Delete Shift:=xlUp
          'Cells(1, 1) = "SN"
          Cells(1, 1) = "ITEM_CODE"
          Cells(1, 2) = "BARCODE"
          Set tempxlSheet = Nothing
          Set tempxlWorkbook = Nothing
          tempxlApp.Quit
          Set tempxlApp = Nothing

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
End Sub
Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub
Private Sub OpenLppx()
    Me.MousePointer = vbHourglass
    ''Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HUAWEI-����.lab")
    ''If isPN = "Y" Then
    'Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\�����ǩģ��\" & "HP���򷢻���ǩ.lab")
    ''Else
    ''Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\�����ǩģ��\" & "HP���򷢻���ǩ_NO_PN.lab")
    ''End If
    'Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\�����ǩģ��\" & "HP������ǩNEW.lab")
    'Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\�����ǩģ��\" & "��HPE������ǩNEW.lab")
    Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\�����ǩģ��\" & "��HPE������ǩNEW1.lab")
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub

Private Sub cmdCancel_HPSN_Click()
txtSN.Text = ""
txtProduct.Text = ""
txtDesc.Text = ""
txtUPC.Text = ""
txtPN.Text = ""
txtSN.SetFocus

End Sub

Private Sub cmdPrint_HPSN_Click()

    If txtSN.Text = "" Then
        MsgBox ("���к�δ���룬���ܴ�ӡ��")
        txtSN.SetFocus
        Exit Sub
    End If
    'If txtPN.Text = "" Then
     '   MsgBox ("����δ���������ܴ�ӡ��")
      '  txtSN.SetFocus
       ' Exit Sub
    'End If
    'If txtProduct.Text = "" Then
     '   MsgBox ("��Ʒ����δ���������ܴ�ӡ��")
      '  Exit Sub
    'End If
    If txtDesc.Text = "" Then
        MsgBox ("��Ʒ����δ���������ܴ�ӡ��")
        Exit Sub
    End If
    'If txtUPC.Text = "" Then
    '    MsgBox ("��ƷUPCδ���������ܴ�ӡ��")
    '    Exit Sub
    'End If

    If txtModel.Text = "" Then
        MsgBox ("����������ITEM_CODE������Ϊ�գ�")
        Exit Sub
    End If
    
    OpenLppx
    
    If txtSZ.Text <> "SZ" Then
         myObjs("SZ").Top = 10000
    End If
    
    myVars.Item("ID").Value = txtDesc.Text
    'myVars.Item("SN1").Value = UCase(txtSN.Text)
    'myVars.Item("SN2").Value = "S" & UCase(txtSN.Text)
    myVars.Item("SN2").Value = UCase(txtSN.Text)
    myVars.Item("UPC").Value = Left(txtUPC.Text, 11)
    'myVars.Item("UPC").Value = "88278128481"
    If txtPN.Text = "" Then
    'myVars.Item("PN1").Value = ""
    'myObjs("tpn1").Top = 10000
    myObjs("bcPN").Top = 100000
    Else
    'myVars.Item("PN1").Value = UCase(txtPN.Text)
    'myVars.Item("PN2").Value = "P" & UCase(txtPN.Text)
    myVars.Item("PN2").Value = UCase(txtPN.Text)
    End If
    'myVars.Item("Product1").Value = UCase(txtProduct.Text)
    'myVars.Item("Product2").Value = "1P" & UCase(txtProduct.Text)
    myVars.Item("Product2").Value = UCase(txtProduct.Text)
    
    If Trim(txtMac.Text) <> "" Then
        myVars.Item("MAC1").Value = UCase(Trim(txtMac.Text))
    Else
        myObjs("MAC(1)").Top = 10000
        myObjs("MAC").Top = 10000
        myObjs("MAC1").Top = 10000
    End If

    If Trim(txtImei1.Text) <> "" Then
        myVars.Item("IMEI1").Value = UCase(Trim(txtImei1.Text))
    Else
        myObjs("IMEI1").Top = 10000
        myObjs("Barcode6").Top = 10000
        myObjs("IMEI2").Top = 10000
        myObjs("Barcode7").Top = 10000
    End If

    If Trim(txtImei2.Text) <> "" Then
        myVars.Item("IMEI2").Value = UCase(Trim(txtImei2.Text))
    Else
        myObjs("IMEI2").Top = 10000
        myObjs("Barcode7").Top = 10000
    End If
    
    If Trim(txtPN.Text) = "" Then
        myObjs("Text1(16)").Top = 10000
        myObjs("bcPN").Top = 10000
    Else
        myVars.Item("PN2").Value = UCase(txtPN.Text)
    End If
    
    If Trim(txtUPC.Text) = "" Then
        myObjs("Text1(21)").Top = 10000
        myObjs("Barcode26(6)").Top = 10000
    Else
        myVars.Item("UPC").Value = Left(txtUPC.Text, 11)
    End If
    
    myDoc.PrintLabel 1
    myDoc.FormFeed
    UnloadLppx
    cmdCancel_HPSN_Click
    
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    MsgBox "��ʹ��������ӡ"
    Exit Sub
    
    If (KeyAscii = 13) Then
        If Len(txtSN.Text) < 10 Then
            MsgBox "��Ʒ��ų��Ȳ���С��10!"
            txtSN.SetFocus
            Exit Sub
        End If
sql = "select * from hp where hp_sn_iii=substring('" & Trim(txtSN.Text) & "',5,3)"
        'MsgBox (sql)
        
        rec.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rec.EOF = True Then
            MsgBox "�����к�δά����Ϣ!"
            txtSN.Text = ""
            txtSN.SetFocus
            rec.Close
            Exit Sub
        Else
        
            If IsNull(rec.Fields("hp_pn")) Then
                'MsgBox ("�����к�δά������!")
                'rec.Close
                'Exit Sub
            Else
                txtPN = rec.Fields("hp_pn")
            End If
             
            If IsNull(rec.Fields("hp_gtin_number")) Then
                'MsgBox ("�����к�δά��UPC!")
                'rec.Close
                'Exit Sub
                txtUPC = ""
            Else
                txtUPC = rec.Fields("hp_gtin_number")
            End If
        
            If IsNull(rec.Fields("hp_product")) Then
                'MsgBox ("�����к�δά����Ʒ����!")
                'rec.Close
                'Exit Sub
            Else
                txtProduct = rec.Fields("hp_product")
            End If
      
            If IsNull(rec.Fields("hp_desc1")) Then
                MsgBox ("�����к�δά��������Ϣ!")
                rec.Close
                Exit Sub
            Else
                txtDesc = rec.Fields("hp_desc1")
            End If
            
            If Not IsNull(rec.Fields("hp_desc2")) Then
                txtDesc = txtDesc & " " & rec.Fields("hp_desc2")
            End If
        
        End If
        rec.Close
       '============add by carson start for TR5=============
'        Dim conSZ As ADODB.Connection
'        Dim rsSZ As ADODB.Recordset
'        Set conSZ = New ADODB.Connection
'        Set rsSZ = New ADODB.Recordset
'        conSZ.ConnectionString = "Provider=SQLOLEDB;User ID=sa;PWD=Flash123;Initial Catalog=afg_active_90;Data Source=10.11.1.130"
'        conSZ.ConnectionTimeout = 50
'        conSZ.Open
''        Dim stringSQL As String
'        Set rsSZ.ActiveConnection = conSZ
'        rsSZ.CursorType = adOpenDynamic
'
'        stringSQL = " select TOP 1 'SZ' from C_NoTR5_Part where EFFE_FLAG='1' AND  Part_Number ='" & txtPN.Text & "'  "
'
'        rsSZ.Open stringSQL
'        If rsSZ.EOF = True Then
'            txtSZ.Text = ""
'        Else
'            txtSZ.Text = rsSZ.Fields(0)
'        End If
'        rsSZ.Close
      '============add by carson end  =============
    End If
End Sub

Private Sub cmdReturn_HPSN_Click()
Unload Me
End Sub

