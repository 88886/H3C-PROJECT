VERSION 5.00
Begin VB.Form frmDaHuashenchan 
   Caption         =   "������"
   ClientHeight    =   10935
   ClientLeft      =   3765
   ClientTop       =   225
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   12011.81
   ScaleMode       =   0  'User
   ScaleWidth      =   13050
   Begin VB.TextBox txtEN 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   37
      Top             =   7680
      Width           =   3255
   End
   Begin VB.CheckBox chkN4 
      Caption         =   "N4"
      Height          =   495
      Left            =   9840
      TabIndex        =   35
      Top             =   5160
      Width           =   615
   End
   Begin VB.CheckBox chkN 
      Caption         =   "N*"
      Height          =   495
      Left            =   9120
      TabIndex        =   34
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtVer 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9120
      TabIndex        =   33
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtSZ 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   450
      Left            =   9120
      TabIndex        =   31
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox txtCustomerCode 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   9960
      Width           =   4575
   End
   Begin VB.TextBox txtDCSign 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7800
      Width           =   4575
   End
   Begin VB.TextBox txtWeight 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   8520
      Width           =   4575
   End
   Begin VB.TextBox txtSize 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   9240
      Width           =   4575
   End
   Begin VB.PictureBox Picture15 
      Height          =   495
      Left            =   7320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   5640
      Width           =   615
   End
   Begin VB.CheckBox chkNonCCC 
      BackColor       =   &H0000C000&
      Caption         =   "��"
      Enabled         =   0   'False
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
      Left            =   10320
      TabIndex        =   20
      Top             =   5760
      Width           =   735
   End
   Begin VB.CheckBox chkCCC 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtACSign 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7080
      Width           =   4575
   End
   Begin VB.TextBox txtRev 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6720
      TabIndex        =   16
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton CommandCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton CommandPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   14
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkY 
      Caption         =   "Y*"
      Height          =   495
      Left            =   9840
      TabIndex        =   12
      Top             =   4800
      Width           =   615
   End
   Begin VB.CheckBox chkY2 
      Caption         =   "Y2"
      Height          =   495
      Left            =   9120
      TabIndex        =   11
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtSN 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2400
      TabIndex        =   10
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txtPart 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox txtModel 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5640
      Width           =   4575
   End
   Begin VB.TextBox txtExeStandard 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   6360
      Width           =   4575
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   4920
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   12615
      Begin VB.Image Image3 
         Height          =   5385
         Left            =   0
         Picture         =   "frmDaHuashenchan.frx":0000
         Top             =   -360
         Width           =   13365
      End
   End
   Begin VB.Label Label14 
      Caption         =   "Ӣ������:"
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
      Left            =   7200
      TabIndex        =   36
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SZ:"
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
      Left            =   7680
      TabIndex        =   32
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000004&
      Caption         =   "�ͻ����룺"
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
      Left            =   600
      TabIndex        =   29
      Top             =   10080
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000004&
      Caption         =   "������Դ��"
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
      Left            =   600
      TabIndex        =   28
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000004&
      Caption         =   "ֱ����Դ��"
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
      Left            =   600
      TabIndex        =   27
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000004&
      Caption         =   "������"
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
      Left            =   600
      TabIndex        =   26
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000004&
      Caption         =   "�汾:"
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
      Left            =   5640
      TabIndex        =   22
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "ִ�б�׼��"
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
      Left            =   600
      TabIndex        =   18
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "�������ԣ�"
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
      Left            =   7320
      TabIndex        =   13
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9000
      Y1              =   4481.772
      Y2              =   4481.772
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "SN��"
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
      Left            =   720
      TabIndex        =   9
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "���֣�"
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
      Left            =   600
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ�ͺţ�"
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
      Left            =   600
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      Caption         =   "����ߴ磺"
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
      Left            =   600
      TabIndex        =   5
      Top             =   9360
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "��Ʒ���"
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
      Left            =   600
      TabIndex        =   4
      Top             =   5040
      Width           =   1575
   End
End
Attribute VB_Name = "frmDaHuashenchan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim rs As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim mydoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim hpsn As String
'Dim myApp2 As New LabelManager2.Application
'Dim myDoc2 As LabelManager2.Document
'Dim myVars2 As LabelManager2.Variables
'Dim myObjs2 As LabelManager2.DocObjects
Dim strDir As String
Dim rec As New ADODB.Recordset
Dim res2 As New ADODB.Recordset
Dim rec13 As New ADODB.Recordset
Dim newLableFlag As Boolean

Private Sub chkY_Click()
    If chkY.Value = 1 Then
        chkY2.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub chkY2_Click()
    If chkY2.Value = 1 Then
        chkY.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub chkN_Click()
    If chkN.Value = 1 Then
        chkY2.Value = 0
        chkY.Value = 0
        chkN4.Value = 0
    End If
End Sub

Private Sub chkN4_Click()
    If chkN4.Value = 1 Then
        chkY.Value = 0
        chkN.Value = 0
        chkY2.Value = 0
    End If
End Sub

Private Sub CommandCancel_Click()
    cmdCancel_HPSN_Click
End Sub

Private Sub CommandPrint_Click()
   Dim Pb As String
    
    If verifyPB() = False Then
        Exit Sub
    End If
    
    If (chkY2.Value = 1) Then
        Pb = CommonValue.PB_Y2
    ElseIf (chkY.Value = 1) Then
        Pb = CommonValue.PB_Y
    ElseIf (chkN.Value = 1) Then
        Pb = CommonValue.PB_N
    ElseIf (chkN4.Value = 1) Then
        Pb = CommonValue.PB_N4
    End If


    
      If UploadH3CInfo2(False, Trim(Me.txtSN.Text), Trim(Me.txtVer.Text), "NA", "", "CHINA", golUSERNAME, Pb) = False Then
                MsgBox "���ϱ���ʧ�ܲ��ܴ�ӡ!", vbInformation + vbOKOnly, "���ϱ���ʧ��"
                xtSN.SetFocus
                Exit Sub
        End If
    
        cmdPrint_HPSN_Click
        
        Sleep 2000

        Call Connect.addPrintedLabel(Me.txtSN.Text, Me.Name)

        cmdCancel_HPSN_Click
End Sub

Private Sub Form_Load()
    Me.Show
    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    
   txtSN.SetFocus
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If conn.State = 1 Then
      conn.Close
      Set conn = Nothing
   End If
End Sub
Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '���ĵ�����ʹ��CloseAll�������ر������ĵ�
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
    Me.MousePointer = vbHourglass
    myApp.EnableEvents = True
    

    Set mydoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\H3C������ǩNEW\����H3CNEW\��H3C��ַģ��\" & "��-��������.lab")
    
    Me.MousePointer = vbDefault
    Set myVars = mydoc.Variables
    Set myObjs = mydoc.DocObjects
End Sub


Private Sub cmdPrint_HPSN_Click()

    If txtSN.Text = "" Then
        MsgBox ("SN���к�δ���룬���ܴ�ӡ��")
        txtSN.SetFocus
        Exit Sub
    End If
    
    If txtPart.Text = "" Then
        MsgBox ("����δ���������ܴ�ӡ��")
        Exit Sub
    End If
    
    If txtRev.Text = "" Then
        MsgBox ("�汾δ���������ܴ�ӡ��")
        Exit Sub
    End If
    
    If txtDesc.Text = "" Then
        MsgBox ("��Ʒ���δ���������ܴ�ӡ��")
        Exit Sub
    End If
   
    If txtModel.Text = "" Then
        MsgBox ("��Ʒ�ͺ�δ���������ܴ�ӡ��")
        Exit Sub
    End If
    
'    If Trim(txtExeStandard.Text) = "" Then
'         MsgBox ("ִ�б�׼δ���������ܴ�ӡ��")
'         Exit Sub
'    End If
    
   If Trim(txtWeight.Text) = "" Then
        MsgBox ("����δ���������ܴ�ӡ��")
        Exit Sub
   End If
       
   If Trim(txtSize.Text) = "" Then
        MsgBox ("����ߴ�δ���������ܴ�ӡ��")
        Exit Sub
   End If
   
   If Trim(txtCustomerCode.Text) = "" Then
        MsgBox ("�ͻ�����δ���������ܴ�ӡ��")
        Exit Sub
   End If
   
   If chkCCC.Value = 0 And chkNonCCC.Value = 0 Then
'        MsgBox ("CCCδ���������ܴ�ӡ��")
'        Exit Sub
   Else
        If chkCCC.Value = 1 Then
            CCC = "1"
        Else
            CCC = "0"
        End If
   End If
   
   If chkY2.Value = 0 And chkY.Value = 0 And chkN.Value = 0 And chkN4.Value = 0 Then
        MsgBox ("SN���кŵ�Pbδ���������ܴ�ӡ��")
        Exit Sub
   Else
        If chkY2.Value = 1 Then
            Pb = "Y2"
        ElseIf chkY.Value = 1 Then
            Pb = "Y*"
        ElseIf chkN.Value = 1 Then
            Pb = "N*"
        ElseIf chkN4.Value = 1 Then
            Pb = "N4"
        End If
   End If

    OpenLppx

    myVars.Item("Product Name1").Value = UCase(Trim(txtDesc.Text))
    myVars.Item("Product ID").Value = UCase(Trim(txtModel.Text))
    myVars.Item("Product Name1").Value = UCase(Trim(txtDesc.Text))
    myVars.Item("Size").Value = Trim(txtSize.Text)
    myVars.Item("Weight").Value = Trim(txtWeight.Text)
    If Trim(txtExeStandard.Text) <> "" Then
        myVars.Item("MS").Value = UCase(Trim(txtExeStandard.Text))
    Else
        myObjs("MS").Top = 10000
        myObjs("MS Title").Top = 10000
        myObjs("Text1(7) Copy(3)").Top = 10000
    End If
    
    If Trim(txtACSign.Text) <> "" Then
        myVars.Item("Jiaoliu").Value = Trim(txtACSign.Text)
    Else
        myObjs("Jiaoliu").Top = 10000
        myObjs("Image4").Top = 10000
    End If
    
    If Trim(txtDCSign.Text) <> "" Then
        myVars.Item("Zhiliu").Value = Trim(txtDCSign.Text)
    Else
        myObjs("Zhiliu").Top = 10000
        myObjs("Image3").Top = 10000
    End If
    
    If Trim(txtACSign.Text) = "" And Trim(txtDCSign.Text) = "" Then
        myObjs("Text1(9)").Top = 10000
    End If
    
    If txtSZ.Text <> "SZ" Then
        myObjs("SZ").Top = 10000
    End If
    
    myVars.Item("P/N").Value = UCase(Trim(txtCustomerCode.Text))
    myVars.Item("SN").Value = UCase(Trim(txtSN.Text)) + " " + UCase(Trim(txtRev.Text)) + " " + Pb
    
    
    If Trim(txtEN.Text) <> "" And Trim(txtEN.Text) <> "/" Then
       myVars.Item("Product Name2").Value = Trim(txtEN.Text)
    Else
        myObjs("Product Name2").Top = 10000
    End If
    
    
    
'    If CCC = 1 Then
'        myObjs("addr").Top = 10000
'    Else
'        myObjs("addr3c").Top = 10000
'    End If

    mydoc.PrintLabel 1
    mydoc.FormFeed
    UnloadLppx
                        
End Sub


Private Sub cmdCancel_HPSN_Click()
txtSN.Text = ""
txtPart.Text = ""
txtRev.Text = ""
txtDesc.Text = ""
txtModel.Text = ""
txtExeStandard.Text = ""
txtACSign.Text = ""
txtDCSign.Text = ""
txtWeight.Text = ""
txtSize.Text = ""
txtCustomerCode.Text = ""
txtEN.Text = ""
txtSN.SetFocus

chkY.Value = 0
chkY2.Value = 0
chkN.Value = 0
chkN4.Value = 0
chkCCC.Value = 0
chkNonCCC.Value = 0
'chkY.Enabled = True
'chkY2.Enabled = True

End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
    
        Dim strModel As String
        Dim strVer As String
        Dim strIII As String
        Dim str2 As String
        Dim strPartNumber As String, Status As String, uploadPowerCode As Boolean
        
        txtPart.Text = ""
        txtRev.Text = ""
        txtDesc.Text = ""
        txtModel.Text = ""
        txtExeStandard.Text = ""
        txtACSign.Text = ""
        txtDCSign.Text = ""
        txtWeight.Text = ""
        txtSize.Text = ""
        txtCustomerCode.Text = ""
        txtSN.SetFocus
        
        chkY.Value = 0
        chkY2.Value = 0
        chkN.Value = 0
        chkN4.Value = 0
        chkCCC.Value = 0
        chkNonCCC.Value = 0
        
        Me.txtSN.Text = Trim(Me.txtSN.Text)
        
        '==================
        Dim con As ADODB.Connection
        Dim rs3 As ADODB.Recordset
        Dim rs4 As ADODB.Recordset
        

        Set con = New ADODB.Connection

        Set rs4 = New ADODB.Recordset
        
        con.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
        con.ConnectionTimeout = 50
        con.Open
        Dim str As String

                
        Set rs4.ActiveConnection = con
        rs4.CursorType = adOpenForwardOnly
        

            Dim con13 As ADODB.Connection
            Dim rs13 As ADODB.Recordset
            Dim com As ADODB.Command
 
            Set con13 = New ADODB.Connection
            Set rs13 = New ADODB.Recordset
            strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
            'con13.ConnectionTimeout = 50
            con13.Open ConnectionString:=strConn
            Set com = New ADODB.Command
            com.ActiveConnection = con13
'            str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtSN.Text) & "'"
'             str = " select top 1 t.part_number,t.part_revision,t.creation_time,t.order_number from (" & _
'            "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
'            "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
'            "where b.original_sn_S = '" & Trim(Me.txtSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
'            Debug.Print str
'            com.CommandText = str
'            rs13.Open Source:=com
'            If rs13.EOF = True Then
''                MsgBox "û�ж�Ӧ���ְ汾��Ϣ", vbOKOnly + vbExclamation, "����"
''                cmdCancel_HPSN_Click
''                rs13.Close
''                Exit Sub
'
'            Else
                'txtPart.Text = rs13.Fields(0)
              '  txtRev.Text = rs13.Fields(1)
              txtPart.Text = "HWF" & Mid(Trim(Me.txtSN.Text), 3, 8)
              
'             sql = "select Rev from tblDaHuaNew where Part_Number='" & Mid(txtPart.Text, 4, 8) & "'"
'              If rec.State = 1 Then
'               rec.Close
'              End If
'
'             rec.Open sql, conn, adOpenKeyset, adLockReadOnly
'
'              If rec.EOF = True Then
'                  MsgBox "�����к�δά����Ϣ!"
'                  cmdCancel_HPSN_Click
'                  rec.Close
'                  Exit Sub
'              Else
'                 txtRev.Text = rec.Fields(0)
'              End If
              
      
      
      
                
                If getDaHuaInformation = False Then
                    If con13.State = 1 Then
                        con13.Close
                    End If
                    
                    If con.State = 1 Then
                        con.Close
                    End If
                    Exit Sub
                End If
                
                
       '============add by carson start for TR5=============
        Dim conSZ As ADODB.Connection
        Dim rsSZ As ADODB.Recordset
        Set conSZ = New ADODB.Connection
        Set rsSZ = New ADODB.Recordset
        conSZ.ConnectionString = "Provider=SQLOLEDB;User ID=sa;PWD=Flash123;Initial Catalog=afg_active_90;Data Source=10.11.1.130"
        conSZ.ConnectionTimeout = 50
        conSZ.Open
'        Dim stringSQL As String
        Set rsSZ.ActiveConnection = conSZ
        rsSZ.CursorType = adOpenDynamic

        stringSQL = " select TOP 1 'SZ' from C_NoTR5_Part where EFFE_FLAG='1' AND  Part_Number ='" & Mid(txtSN.Text, 3, 8) & "'  "

        rsSZ.Open stringSQL
        If rsSZ.EOF = True Then
            txtSZ.Text = ""
        Else
            txtSZ.Text = rsSZ.Fields(0)
        End If
        rsSZ.Close
      '============add by carson end  =============
      
    'add by carson 20160810 for Roy required  software
    '''''''''''''''''''''''''''''''''''''''''''''''
        Dim conSQL01_1 As ADODB.Connection
        Dim rsSQL01_1 As ADODB.Recordset
        Dim comSQL01_1 As ADODB.Command
        Set conSQL01_1 = New ADODB.Connection
        Set rsSQL01_1 = New ADODB.Recordset
        strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.31;Initial Catalog=dataT; User ID=sa; PWD=Itadmin1"
        
        lRemark = "N/A"
        conSQL01_1.Open ConnectionString:=strConn
        Set comSQL01_1 = New ADODB.Command
        comSQL01_1.ActiveConnection = conSQL01_1
        
        str = "select top 1 remark from [test_equ_ATE] " & _
        " where barcode = '" & Trim(Me.txtSN.Text) & "' and pass = N'ͨ��' and remark<>'' order by id DESC "
        comSQL01_1.CommandText = str
        rsSQL01_1.Open Source:=comSQL01_1
        If rsSQL01_1.EOF = False Then
            lRemark = UCase(Trim(rsSQL01_1.Fields("remark")))
        End If
        rsSQL01_1.Close
        txtVer.Text = lRemark
    '''''''''''''''''''''''''''''''''''''''''''''''

    '===============add by ben 2012-02-05 start===============
'                    If Connect.reprint = False Then
'    '                If reprint = False Then
'                        If Connect.isPrintedLabel(Me.txtSN.Text, Me.Name) Then
'                            MsgBox ("�����к��Ѵ�ӡ��")
'                            cmdCancel_HPSN_Click
'                            If con.State = 1 Then
'                                con.Close
'                            End If
'
'                            If con13.State = 1 Then
'                                con13.Close
'                            End If
'                            Exit Sub
'                        End If
'                    End If
'                    rec.Close
    '======================================================================
    
'                    Dim lh As New Label_History, Pb As String
'                    Dim sn As String
'                    sn = txtSN.Text
'                    If (lh.Init(sn)) Then
'                        If lh.Pb = "Y*" Then
'                            chkY.Value = 1
'                            chkY2.Value = 0
'                            chkN.Value = 0
'                            chkN4.Value = 0
'                            chkY.Enabled = False
'                            chkY2.Enabled = False
'                            chkN.Enabled = False
'                            chkN4.Enabled = False
'                            Pb = CommonValue.PB_Y
'                        ElseIf lh.Pb = "Y2" Then
'                            chkY.Value = 0
'                            chkY2.Value = 1
'                            chkN.Value = 0
'                            chkN4.Value = 0
'                            chkY.Enabled = False
'                            chkY2.Enabled = False
'                            chkN.Enabled = False
'                            chkN4.Enabled = False
'                            Pb = CommonValue.PB_Y2
'                        ElseIf lh.Pb = "N*" Then
'                            chkY.Value = 0
'                            chkY2.Value = 0
'                            chkN.Value = 1
'                            chkN4.Value = 0
'                            chkY.Enabled = False
'                            chkY2.Enabled = False
'                            chkN.Enabled = False
'                            chkN4.Enabled = False
'                            Pb = CommonValue.PB_N
'                        ElseIf lh.Pb = "N4" Then
'                            chkY.Value = 0
'                            chkY2.Value = 0
'                            chkN.Value = 0
'                            chkN4.Value = 1
'                            chkY.Enabled = False
'                            chkY2.Enabled = False
'                            chkN.Enabled = False
'                            chkN4.Enabled = False
'                            Pb = CommonValue.PB_N4
'                        End If
'                    Else
'                        chkY.Enabled = True
'                        chkY2.Enabled = True
'                        chkN.Enabled = True
'                        chkN4.Enabled = True
'                        TextModel.Text = strModel
'                        TextType.Text = str2
'                        TextVersion.Text = strVer
'                        TextPowerCode.Text = CStr(uploadPowerCode)
'                        TextStatus.Text = Status
'                        MsgBox "��ȷ�ϻ�������ֵ"
'                        Exit Sub
'                    End If
''

                 
'                     If UploadDaHuaShip_Rec(Trim(UCase(txtSN.Text)), Trim(UCase(txtPart.Text)), Trim(UCase(txtRev.Text)), Trim(UCase(txtHKPart.Text)), Trim(UCase(txtHKSN.Text)), Trim(UCase(txtDesc.Text)), Trim(UCase(txtType.Text)), Trim(UCase(txtMaterial.Text)), Trim(UCase(txtOrder.Text)), IIf(chkY2.Value, "1", "0"), IIf(chkCCC.Value, "1", "0"), golUSERNAME) = False Then
'                        MsgBox "�����ϱ���ʧ�ܲ��ܴ�ӡ!", vbInformation + vbOKOnly, "���ϱ���ʧ��"
'                        '        txtSN.SetFocus
'                        UnloadLppx
'                        Exit Sub
'                    End If
    '===============add by ben 2012-02-05 end=================
          End If
                 
End Sub

Private Sub OpenLppx2(model As String)
    Me.MousePointer = vbHourglass

'    Set myDoc2 = myApp2.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\Phase3 HP��֤������ǩ\" & model & ".lab")
'
'    Me.MousePointer = vbDefault
'    Set myVars2 = myDoc2.Variables
'    Set myObjs2 = myDoc2.DocObjects
    Set mydoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\Phase3 HP��֤�·�����ǩ\" & model & ".lab")
    
    Me.MousePointer = vbDefault
    Set myVars = mydoc.Variables
    Set myObjs = mydoc.DocObjects
End Sub

Private Function verifyPB() As Boolean
    If (chkY2.Value + chkY.Value + chkN.Value + chkN4.Value = 1) = False Then
        MsgBox "��ȷ�ϻ�������"
        verifyPB = False
        Exit Function
    End If
    verifyPB = True
End Function


Private Sub UnloadLppx2()
    myApp.Documents.CloseAll False
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Function getDaHuaInformation() As Boolean

      sql = "select * from tblDaHuaNew where Part_Number='" & Mid(txtPart.Text, 4, 8) & "'"
      If rec.State = 1 Then
        rec.Close
      End If
      
      rec.Open sql, conn, adOpenKeyset, adLockReadOnly
      
      If rec.EOF = True Then
          MsgBox "�����к�δά����Ϣ!"
          txtSN.Text = ""
'          txtSN.SetFocus
          rec.Close
          getDaHuaInformation = False
          Exit Function
      Else
          If IsNull(rec.Fields("Part_Number")) Then
              MsgBox ("�����к�δά����Ʒ����!")
              rec.Close
              getDaHuaInformation = False
              Exit Function
'          Else
'              txtPart.Text = rec.Fields("Part_Number")
          End If
    
          If IsNull(rec.Fields("Part_Desc")) Then
              MsgBox ("�����к�δά����Ʒ���!")
              rec.Close
              getDaHuaInformation = False
              Exit Function
          Else
              txtDesc.Text = rec.Fields("Part_Desc")
          End If
          
          If IsNull(rec.Fields("Part_Model")) Then
              MsgBox ("�����к�δά����Ʒ�ͺ�!")
              rec.Close
              getDaHuaInformation = False
              Exit Function
          Else
              txtModel.Text = rec.Fields("Part_Model")
          End If
    
          If IsNull(rec.Fields("ExeStandard")) Then
'              MsgBox ("�����к�δά��ִ�б�׼")
'              rec.Close
'              getDaHuaInformation = False
'              Exit Function
               txtExeStandard.Text = rec.Fields("ExeStandard")
          Else
              txtExeStandard.Text = rec.Fields("ExeStandard")
          End If
          
          If IsNull(rec.Fields("Weight")) Then
              MsgBox ("�����к�δ����!")
              rec.Close
              getDaHuaInformation = False
              Exit Function
          Else
              txtWeight.Text = rec.Fields("Weight")
          End If
          
          If IsNull(rec.Fields("Size")) Then
              MsgBox ("�����к�δά������ߴ�!")
              rec.Close
              getDaHuaInformation = False
              Exit Function
          Else
              txtSize.Text = rec.Fields("Size")
          End If

          If IsNull(rec.Fields("CustomerCode")) Then
              MsgBox ("�����к�δά���ͻ�����!")
              rec.Close
              getDaHuaInformation = False
              Exit Function
          Else
              txtCustomerCode.Text = rec.Fields("CustomerCode")
          End If
          
          If IsNull(rec.Fields("ACSign")) Then
              txtACSign.Text = ""
          Else
              txtACSign.Text = rec.Fields("ACSign")
          End If
          
          
           If IsNull(rec.Fields("Part_ENDesc")) Or rec.Fields("Part_ENDesc") = "/" Then
              txtEN.Text = ""
          Else
              txtEN.Text = rec.Fields("Part_ENDesc")
          End If
        
          If IsNull(rec.Fields("DCSign")) Then
              txtDCSign.Text = ""
          Else
              txtDCSign.Text = rec.Fields("DCSign")
          End If
          
          If UCase(Trim(rec.Fields("CCC"))) = True Then
              chkCCC.Value = 1
              chkNonCCC.Value = 0
          Else
              chkCCC.Value = 0
              chkNonCCC.Value = 1
          End If

'          If GetPbProperty(txtSN.Text) = "Y2" Then
'              chkY2.Value = 1
'              chkY.Value = 0
'              chkN.Value = 0
'              chkN4.Value = 0
'          ElseIf GetPbProperty(txtSN.Text) = "Y*" Then
'              chkY2.Value = 0
'              chkY.Value = 1
'              chkN.Value = 0
'              chkN4.Value = 0
'          ElseIf GetPbProperty(txtSN.Text) = "N4" Then
'              chkY2.Value = 0
'              chkY.Value = 0
'              chkN.Value = 0
'              chkN4.Value = 1
'          ElseIf GetPbProperty(txtSN.Text) = "N*" Then
'              chkY2.Value = 0
'              chkY.Value = 0
'              chkN.Value = 1
'              chkN4.Value = 0
'          Else
'              chkY2.Value = 0
'              chkY.Value = 0
'              chkN.Value = 0
'              chkN4.Value = 0
'              MsgBox ("�����к�Pb����δ����")
'              rec.Close
'              getDaHuaInformation = False
'              Exit Function
'          End If
          
      End If
      getDaHuaInformation = True
End Function

