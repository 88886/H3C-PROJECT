VERSION 5.00
Begin VB.Form frmHP14677NX 
   Caption         =   "����-��HP��Ʒ--HP SN��ǩ14.6*7.7"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   8025
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chkN 
      Caption         =   "N*"
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
      Left            =   4320
      TabIndex        =   23
      Top             =   5640
      Width           =   615
   End
   Begin VB.CheckBox chkBackup 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "����ǩ"
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
      Left            =   3000
      TabIndex        =   22
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CheckBox chkMain 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "����ǩ"
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
      Left            =   1080
      TabIndex        =   21
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtModel 
      Height          =   285
      Left            =   7320
      TabIndex        =   20
      Top             =   7440
      Width           =   495
   End
   Begin VB.CheckBox chkY2 
      Caption         =   "Y2"
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
      Left            =   3480
      TabIndex        =   19
      Top             =   5640
      Width           =   615
   End
   Begin VB.CheckBox chkY 
      Caption         =   "Y*"
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
      Left            =   2640
      TabIndex        =   18
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtRevision 
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
      Left            =   5040
      TabIndex        =   15
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtPart 
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
      Width           =   2055
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
      Left            =   1080
      Picture         =   "frmHP14677NX.frx":0000
      ScaleHeight     =   2040
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   0
      Width           =   6015
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
      Left            =   1080
      TabIndex        =   17
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
      Left            =   4320
      TabIndex        =   16
      Top             =   5040
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
      Left            =   1080
      TabIndex        =   14
      Top             =   5040
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8760
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
Attribute VB_Name = "frmHP14677NX"
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
Dim myApp2 As New LabelManager2.Application
Dim myDoc2 As LabelManager2.Document
Dim myVars2 As LabelManager2.Variables
Dim myObjs2 As LabelManager2.DocObjects
Dim rs As New Recordset



Private Sub cmdCancel_HPSN_Click()
    Me.txtSN.Text = ""
    Me.txtProduct.Text = ""
    Me.txtDesc1.Text = ""
    Me.txtDesc2.Text = ""
    Me.txtPart.Text = ""
    Me.txtRevision.Text = ""
    Me.chkY.Value = 0
    Me.chkY2.Value = 0
End Sub

Private Sub cmdMPrint_Click()
 cmdReturn_HPSN.Enabled = False
'cmdPrint_HPSN.Enabled = False
cmdCancel_HPSN.Enabled = False
sql = "select ITEM_CODE,BARCODE from tblHP_Print where isnull(BARCODE,'')<>'' and isnull(ITEM_CODE,'')<>'' order by BARCODE"
If conn1.State <> 1 Then
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
    Dim printSet(0 To 1000, 0 To 1000) As String
    Dim total As Integer
    total = rs.RecordCount
    
    For i = 0 To total - 1 Step 1
        printSet(i, 0) = rs("BARCODE")
        printSet(i, 1) = rs("ITEM_CODE")
        rs.MoveNext
    Next
    rs.Close
    
    txtModel.Text = printSet(0, 1)
    
    If HPPrintDouble(txtModel.Text) = 0 Then
        MsgBox "�û���û����HP���к�����ά��Ϊ[14.6*7.7]��ӡѡ��!"
'        Me.txtSN.SetFocus
        Exit Sub
    ElseIf HPPrintDouble(txtModel.Text) = 2 Then
        Me.chkMain.Value = 1
        Me.chkBackup.Value = 1
    ElseIf HPPrintDouble(txtModel.Text) = 1 Then
        Me.chkMain.Value = 1
        Me.chkBackup.Value = 0
    End If

    'begin print
    OpenLppx
    If chkBackup.Value = 1 Then
        OpenLppx2
    End If

    For i = 0 To total - 1 Step 3
        txtSN.Text = printSet(i, 0)
        txtModel.Text = Trim(printSet(i, 1))
        
        
        If InStr(1, txtPart.Text, txtModel.Text) <= 0 Then
            MsgBox ("�ù����Ϻź������Ӧ���ϺŲ�һ�£���ȷ�����빤���Ƿ���ȷ!")
            cmdCancel_HPSN_Click
            Exit Sub
        End If
        updateHPInformation
        
        myVars.Item("SN1").Value = printSet(i, 0)
        If printSet(i + 1, 0) = "" Then
            myVars.Item("SN2").Value = "XXXXXXXXXX"
        Else
            myVars.Item("SN2").Value = printSet(i + 1, 0)
        End If
        
        If printSet(i + 2, 0) = "" Then
            myVars.Item("SN3").Value = "XXXXXXXXXX"
        Else
            myVars.Item("SN3").Value = printSet(i + 2, 0)
        End If
        
'        myVars.Item("SN3").Value = printSet(i + 2, 0)
        
        'myVars.Item("Item").Value = "03" & UCase(Left(txtSN.Text, 6))
        If Me.txtRevision.Text = "" Or Me.txtRevision.Text = "/" Then
            'myObjs("Sver").Top = 5
            myVars.Item("Rev").Value = "N/A"
        ElseIf txtRevision.Text = "00" Then
            myVars.Item("Rev").Value = ""
        Else
            'myObjs("Sver").Top = 5
            myVars.Item("Rev").Value = UCase(Me.txtRevision.Text)
        End If
        myVars.Item("Model").Value = Trim(Me.txtModel.Text)
    
        If chkY.Value = 1 Then
            myVars.Item("Rohs").Value = "Y*"
        ElseIf chkY2.Value = 1 Then
            myVars.Item("Rohs").Value = "Y2"
        ElseIf chkN.Value = 1 Then
            myVars.Item("Rohs").Value = "N*"
        End If
        
        If Trim(Me.txtProduct.Text) <> "" Then
            myVars.Item("PN").Value = Trim(Me.txtProduct.Text)
        Else
            myVars.Item("PN").Value = ""
        End If
        
        
        If Me.txtDesc1.Text <> "" Then
            myVars.Item("desc1").Value = Trim(Me.txtDesc1.Text)
        End If
        If Me.txtDesc2.Text <> "" Then
            myVars.Item("desc2").Value = Trim(Me.txtDesc2.Text)
        Else
            myVars.Item("desc2").Value = ""
        End If
    
        'myApp.Visible = True
        myDoc.PrintLabel 1
        myDoc.FormFeed
        
        If chkBackup.Value = 1 Then
        
            myVars2.Item("SN1").Value = printSet(i, 0)
            If printSet(i + 1, 0) = "" Then
                myVars2.Item("SN2").Value = "XXXXXXXXXX"
            Else
                myVars2.Item("SN2").Value = printSet(i + 1, 0)
            End If
            
            If printSet(i + 2, 0) = "" Then
                myVars2.Item("SN3").Value = "XXXXXXXXXX"
            Else
                myVars2.Item("SN3").Value = printSet(i + 2, 0)
            End If
            
    '        myVars.Item("SN3").Value = printSet(i + 2, 0)
            
            'myVars.Item("Item").Value = "03" & UCase(Left(txtSN.Text, 6))
            If Me.txtRevision.Text = "" Or Me.txtRevision.Text = "/" Then
                'myObjs("Sver").Top = 5
                myVars2.Item("Rev").Value = "N/A"
            ElseIf txtRevision.Text = "00" Then
                myVars2.Item("Rev").Value = ""
            Else
                'myObjs("Sver").Top = 5
                myVars2.Item("Rev").Value = UCase(Me.txtRevision.Text)
            End If
            myVars2.Item("Model").Value = Trim(Me.txtModel.Text)

            
            If Trim(Me.txtProduct.Text) <> "" Then
                myVars2.Item("PN").Value = Trim(Me.txtProduct.Text)
            Else
                myVars2.Item("PN").Value = ""
            End If
        
            If chkY.Value = 1 Then
                myVars2.Item("Rohs").Value = "Y*"
            ElseIf Me.chkN.Value = 1 Then
                myVars2.Item("Rohs").Value = "N*"
            ElseIf Me.chkY2.Value = 1 Then
                myVars2.Item("Rohs").Value = "Y2"
            End If
           
        
            'myApp.Visible = True
            myDoc2.PrintLabel 1
            myDoc2.FormFeed
            
        End If
   Next
   UnloadLppx
   UnloadLppx2
   cmdCancel_HPSN_Click
   If rs.State = 1 Then
        rs.Close
   End If
   
End If
del_sql
cmdReturn_HPSN.Enabled = True
cmdCancel_HPSN.Enabled = True
'MsgBox ("������ӡ�ɹ���")
End Sub

Private Sub cmdPrint_HPSN_Click()
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
    If (Me.chkY2.Value = 1) Then
        myVars.Item("Rohs").Value = "Y2"
    ElseIf chkY.Value = 1 Then
        myVars.Item("Rohs").Value = "Y*"
    ElseIf chkN.Value = 1 Then
        myVars.Item("Rohs").Value = "N*"
    End If
    
        
    
    
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
          updateHPDesc
      End If
    
End Sub

Private Sub updateHPDesc()
      sql = "select part_number,hp_desc,hp_desc2 from ean_hpDesc_maintain where part_number = '" + txtModel.Text + "'"
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
    
          If IsNull(rec.Fields("hp_desc")) Then
              MsgBox ("�����к�δά��������Ϣ!")
              rec.Close
              Exit Sub
          Else
              txtDesc1.Text = rec.Fields("hp_desc")
          End If
    
          If Not IsNull(rec.Fields("hp_desc2")) Then
              txtDesc2.Text = rec.Fields("hp_desc2")
          End If
          rec.Close
      End If
      
    
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub
Private Sub OpenLppx()
    Me.MousePointer = vbHourglass
    Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\�����ǩģ��\HP�����ǩ����\" & "HP SN��ǩ14677-1.lab")
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub

Private Sub OpenLppx2()
    Me.MousePointer = vbHourglass
    Set myDoc2 = myApp2.Documents.Open("\\sz-fs01\Public\Manufacture\�����ǩģ��\HP�����ǩ����\" & "HP SN��ǩ14677-2.lab")
    Me.MousePointer = vbDefault
    Set myVars2 = myDoc2.Variables
    Set myObjs2 = myDoc2.DocObjects
End Sub

Private Sub UnloadLppx2()
    myApp2.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp2.Quit
    Set myApp2 = Nothing
End Sub

Sub del_sql()
    Dim delsql As String
    delsql = "delete from tblHP_Print"
    conn1.Execute delsql
End Sub
