VERSION 5.00
Begin VB.Form frmDoubleSNHP 
   Caption         =   "HP˫SN�������ߴ�ӡ"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9210
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtHPSN 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtPN 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox txtProduct 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4800
      Width           =   2895
   End
   Begin VB.TextBox txtUPC 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   5520
      Width           =   2895
   End
   Begin VB.TextBox txtDesc 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9015
      Begin VB.Image Image1 
         Height          =   1860
         Left            =   4560
         Picture         =   "frmDoubleSNHP.frx":0000
         Top             =   240
         Width           =   4320
      End
      Begin VB.Image Image3 
         Height          =   1815
         Left            =   120
         Picture         =   "frmDoubleSNHP.frx":27D5
         Top             =   240
         Width           =   4305
      End
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "HP SN��"
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
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
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
      Left            =   720
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
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
      Left            =   720
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
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
      Left            =   720
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
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
      Left            =   720
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmDoubleSNHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql As String
Dim sn As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim hpsn As String
Dim myApp2 As New LabelManager2.Application
Dim myDoc2 As LabelManager2.Document
Dim myVars2 As LabelManager2.Variables
Dim myObjs2 As LabelManager2.DocObjects
Dim strDir As String
Dim rec As New ADODB.Recordset
Dim res2 As New ADODB.Recordset
Dim rec13 As New ADODB.Recordset

Private Sub Form_Load()

    Me.Show

    If conn.State = 0 Then
        conn.ConnectionString = Connect.getConnectionstring
        conn.Open
    End If
    
   txtHPSN.SetFocus
   
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

    Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\��ǩģ��\" & "HP������ǩNEW.lab")
    
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub


Private Sub cmdPrint_HPSN_Click()

    If txtHPSN.Text = "" Then
        MsgBox ("���к�δ���룬���ܴ�ӡ��")
        txtSN.SetFocus
        Exit Sub
    End If

    If txtProduct.Text = "" Then
        MsgBox ("��Ʒ����δ���������ܴ�ӡ��")
        Exit Sub
    End If
    If txtDesc.Text = "" Then
        MsgBox ("��Ʒ����δ���������ܴ�ӡ��")
        Exit Sub
    End If

    OpenLppx

         
    myVars.Item("ID").Value = txtDesc.Text
    myVars.Item("SN2").Value = UCase(txtHPSN.Text)
   
    If Trim(txtPN.Text) <> "" Then
        myVars.Item("PN2").Value = UCase(txtPN.Text)
    Else
        myObjs("bcPN").Top = 10000
    End If

    myVars.Item("Product2").Value = UCase(txtProduct.Text)
    myVars.Item("UPC").Value = Left(Trim(txtUPC.Text), 11)
    

    myDoc.PrintLabel 1
    myDoc.FormFeed
    UnloadLppx
    
End Sub


Private Sub cmdCancel_HPSN_Click()
txtHPSN.Text = ""
txtProduct.Text = ""
txtDesc.Text = ""
txtUPC.Text = ""
txtPN.Text = ""
txtHPSN.SetFocus

End Sub


Private Sub txtHPSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
    
        Dim strModel As String
        Dim strVer As String
        Dim strIII As String
        Dim str2 As String
        Dim strPartNumber As String
        
        Me.txtHPSN.Text = Trim(Me.txtHPSN.Text)
        
        '==================
        Dim con As ADODB.Connection
        Dim rs3 As ADODB.Recordset
        Dim rs4 As ADODB.Recordset
        

        Set con = New ADODB.Connection
        Set rs3 = New ADODB.Recordset
        Set rs4 = New ADODB.Recordset
        
        con.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
        con.ConnectionTimeout = 50
        con.Open
        Dim str As String
        Set rs3.ActiveConnection = con
        rs3.CursorType = adOpenDynamic
        Set rs4.ActiveConnection = con
        rs4.CursorType = adOpenDynamic
        
        str = " select top 1 part_number,part_revision,creation_time from (" & _
        " select part_number,part_revision,creation_time from [10.11.1.17].dsactive.dbo.unit nolock " & _
        " where serial_number='" & Trim(Me.txtHPSN.Text) & "'" & _
        " union" & _
        " select part_number,part_rev as part_revision,creation_time from [10.11.1.17].dsactive.dbo.dc_task_order NOLOCK  " & _
        " where order_number=(select order_number from [10.11.1.17].dsactive.dbo.taskorder_unit NOLOCK" & _
        " where serial_number='" & Trim(Me.txtHPSN.Text) & "')" & _
        " ) as t " & _
        " order by t.creation_time desc"
        
        'str = "select part_number,part_revision from [10.11.1.17].dsactive.dbo.unit nolock where serial_number='" & Trim(Me.txtHPSN.Text) & "' "
        rs3.Open str
        If rs3.EOF = True Then
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
'            str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtHPSN.Text) & "'"
             str = " select top 1 part_number,part_revision,creation_time,order_number from (" & _
            "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtHPSN.Text) & "' union " & _
            "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
            "where b.original_sn_S = '" & Trim(Me.txtHPSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
            com.CommandText = str
            rs13.Open Source:=com
            'rs13.Open str
            If rs13.EOF = True Then
                MsgBox "û�ж�Ӧ���ְ汾��Ϣ", vbOKOnly + vbExclamation, "����"
                cmdCancel_HPSN_Click
                rs13.Close
                Exit Sub
            Else
                strPartNumber = rs13.Fields(0)
                strModel = Mid(Trim(rs13.Fields(0)), 4, 8)
                strVer = rs13.Fields(1)
                strIII = Mid(Trim(txtHPSN.Text), 5, 3)
            
              
                
            Dim con14 As ADODB.Connection
            Dim rs14 As ADODB.Recordset
            Dim com14 As ADODB.Command

            Set con14 = New ADODB.Connection
            Set rs14 = New ADODB.Recordset
            strConn = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
            'con13.ConnectionTimeout = 50
            con14.Open ConnectionString:=strConn
            Set com14 = New ADODB.Command
            com14.ActiveConnection = con14
            'str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtHPSN.Text) & "'"
              str = "select 1 from [H3C_HPWeight] " & _
                " where ((Part_Number = '" & strPartNumber & "' and part_revision = '" & strVer & "') or " & _
                " (Part_Number = '" & strPartNumber & "' and part_revision = 'ALL')) " & _
                " and GrossWeight is not null and NetWeight is not null " & _
                " and GrossWeight <> '' and NetWeight <> '' " & _
                " and is_Valid = 1 "
            com14.CommandText = str
               
               'rs14.Open str
                rs14.Open Source:=com14
                If rs14.EOF = True Then
                    MsgBox "û��ά������", vbOKOnly + vbExclamation, "����"
                    cmdCancel_HPSN_Click
                    rs14.Close
                    Exit Sub
                End If

                Set fs = CreateObject("Scripting.FileSystemObject")
                'Dim fs As New FileSystemObject
    
                strDir = "\\sz-fs01\Public\Manufacture\��ǩģ��\Phase3 HP��֤������ǩ\" & strModel & ".lab"
                If Not fs.FileExists(strDir) Then
                    MsgBox "û�ж�Ӧ���ִ�ӡģ��", vbOKOnly + vbExclamation, "����"
                    cmdCancel_HPSN_Click
                    rs3.Close
                    Exit Sub
                End If
                       

                sql = "select * from HP where h3c_bom_code='" & strModel & "' and hp_sn_iii='" & strIII & "'"
                If conn.State = 0 Then
                    conn.ConnectionString = Connect.getConnectionstring
                    conn.Open
                End If
                rec.Open sql, conn, adOpenKeyset, adLockOptimistic
                If rec.EOF = False Then
                    
                    
                    '====================
                     If IsNull(rec.Fields("hp_desc1")) Then
                        MsgBox ("�����к�δά��������Ϣ!")
                        cmdCancel_HPSN_Click
                        rs3.Close
                        rec.Close
                        Exit Sub
                    Else
                        txtDesc = rec.Fields("hp_desc1")
                    End If
                    
                    If Not IsNull(rec.Fields("hp_desc2")) Then
                        txtDesc = txtDesc & " " & rec.Fields("hp_desc2")
                    End If
                
                    If IsNull(rec.Fields("hp_product")) Then
                        MsgBox ("�����к�δά����Ʒ����!")
                        cmdCancel_HPSN_Click
                        rs.Close
                        rec.Close
                        Exit Sub
                    Else
                        txtProduct = rec.Fields("hp_product")
                    End If
                    
                    Dim res2 As New ADODB.Recordset
                    sql = "select * from singleunit where sn='" & strModel & "'"
                    res2.Open sql, conn, adOpenKeyset, adLockOptimistic
                    If res2.EOF = True Then
                        MsgBox ("�����к�δά����Ʒ�ͺ�!")
                        cmdCancel_HPSN_Click
                        res2.Close
                        rs3.Close
                        rec.Close
                        Exit Sub
                    Else
                        str2 = res2.Fields("type")
                    End If
                    res2.Close
                    
                    If IsNull(rec.Fields("hp_pn")) Then
                        txtPN = ""
                    Else
                        txtPN = rec.Fields("hp_pn")
                    End If
                    
                    If IsNull(rec.Fields("hp_gtin_number")) Then
                        txtUPC = ""
                    Else
                        txtUPC = rec.Fields("hp_gtin_number")
                    End If
    '===============add by ben 2012-02-05 start===============
                    If Connect.reprint = False Then
    '                If reprint = False Then
                        If Connect.isPrintedLabel(Me.txtHPSN.Text, Me.Name) Then
                            MsgBox ("�����к��Ѵ�ӡ��")
                            cmdCancel_HPSN_Click
                            If rs3.State = 1 Then
                                rs3.Close
                            End If
                            If rec.State = 1 Then
                                rec.Close
                            End If
                            Exit Sub
                        End If
                    End If
                    rec.Close
    '===============add by ben 2012-02-05 end=================
                    cmdPrint_HPSN_Click
    
                    cmdPrint_Model_Click strModel, str2, strVer
    '===============add by ben 2012-02-05 start===============
                    Call Connect.addPrintedLabel(Me.txtHPSN.Text, Me.Name)
                End If
            End If
            If rs13.State = 1 Then
                rs13.Close
            End If
            If con13.State = 1 Then
                con13.Close
            End If

'            MsgBox "û�ж�Ӧ���ְ汾��Ϣ", vbOKOnly + vbExclamation, "����"
'            cmdCancel_HPSN_Click
'            rs3.Close
'            Exit Sub
        Else
            strPartNumber = rs3.Fields(0)
            strModel = Mid(Trim(rs3.Fields(0)), 4, 8)
            strVer = rs3.Fields(1)
            strIII = Mid(Trim(txtHPSN.Text), 5, 3)
            
            str = "select 1 from [10.11.1.17].[dsActive].[dbo].[H3C_HPWeight] nolock " & _
            " where ((Part_Number = '" & strPartNumber & "' and part_revision = '" & strVer & "') or " & _
            " (Part_Number = '" & strPartNumber & "' and part_revision = 'ALL')) " & _
            " and GrossWeight is not null and NetWeight is not null " & _
            " and GrossWeight <> '' and NetWeight <> '' " & _
            " and is_Valid = 1 "
            rs4.Open str
            If rs4.EOF = True Then
                MsgBox "û��ά������", vbOKOnly + vbExclamation, "����"
                cmdCancel_HPSN_Click
                rs4.Close
                Exit Sub
            End If

            Set fs = CreateObject("Scripting.FileSystemObject")
            'Dim fs As New FileSystemObject


            strDir = "\\sz-fs01\Public\Manufacture\��ǩģ��\Phase3 HP��֤������ǩ\" & strModel & ".lab"
            If Not fs.FileExists(strDir) Then
                MsgBox "û�ж�Ӧ���ִ�ӡģ��", vbOKOnly + vbExclamation, "����"
                cmdCancel_HPSN_Click
                rs3.Close
                Exit Sub
            End If
            
            
            
            sql = "select * from HP where h3c_bom_code='" & strModel & "' and hp_sn_iii='" & strIII & "'"
            If conn.State = 0 Then
                conn.ConnectionString = Connect.getConnectionstring
                conn.Open
            End If
            rec.Open sql, conn, adOpenKeyset, adLockOptimistic
            If rec.EOF = False Then
                
                
                '====================
                 If IsNull(rec.Fields("hp_desc1")) Then
                    MsgBox ("�����к�δά��������Ϣ!")
                    cmdCancel_HPSN_Click
                    rs3.Close
                    rec.Close
                    Exit Sub
                Else
                    txtDesc = rec.Fields("hp_desc1")
                End If
                
                If Not IsNull(rec.Fields("hp_desc2")) Then
                    txtDesc = txtDesc & " " & rec.Fields("hp_desc2")
                End If
            
                If IsNull(rec.Fields("hp_product")) Then
                    MsgBox ("�����к�δά����Ʒ����!")
                    cmdCancel_HPSN_Click
                    rs.Close
                    rec.Close
                    Exit Sub
                Else
                    txtProduct = rec.Fields("hp_product")
                End If
                
      
                sql = "select * from singleunit where sn='" & strModel & "'"
                res2.Open sql, conn, adOpenKeyset, adLockOptimistic
                If res2.EOF = True Then
                    MsgBox ("�����к�δά����Ʒ�ͺ�!")
                    cmdCancel_HPSN_Click
                    res2.Close
                    rs3.Close
                    rec.Close
                    Exit Sub
                Else
                    str2 = res2.Fields("type")
                End If
                res2.Close
                
                If IsNull(rec.Fields("hp_pn")) Then
                    txtPN = ""
                Else
                    txtPN = rec.Fields("hp_pn")
                End If
                
                If IsNull(rec.Fields("hp_gtin_number")) Then
                    txtUPC = ""
                Else
                    txtUPC = rec.Fields("hp_gtin_number")
                End If
'===============add by ben 2012-02-05 start===============
                If Connect.reprint = False Then
'                If reprint = False Then
                    If Connect.isPrintedLabel(Me.txtHPSN.Text, Me.Name) Then
                        MsgBox ("�����к��Ѵ�ӡ��")
                        cmdCancel_HPSN_Click
                        If rs3.State = 1 Then
                            rs3.Close
                        End If
                        If rec.State = 1 Then
                            rec.Close
                        End If
                        Exit Sub
                    End If
                End If
'===============add by ben 2012-02-05 end=================
                cmdPrint_HPSN_Click

                cmdPrint_Model_Click strModel, str2, strVer
'===============add by ben 2012-02-05 start===============
                Call Connect.addPrintedLabel(Me.txtHPSN.Text, Me.Name)
'===============add by ben 2012-02-05 end=================

                '======================
                
                
            Else
                MsgBox "�˲�Ʒ���δ�ռ��汾!"
                cmdCancel_HPSN_Click
                rec.Close
                rs3.Close
                Exit Sub
            End If
            
            rec.Close
            
        End If
        rs3.Close
        
        cmdCancel_HPSN_Click
        
    End If
End Sub

Private Sub OpenLppx2(model As String)
    Me.MousePointer = vbHourglass

    Set myDoc2 = myApp2.Documents.Open("\\sz-fs01\Public\Manufacture\��ǩģ��\Phase3 HP��֤������ǩ\" & model & ".lab")
    
    Me.MousePointer = vbDefault
    Set myVars2 = myDoc2.Variables
    Set myObjs2 = myDoc2.DocObjects
End Sub

Private Sub cmdPrint_Model_Click(strModel As String, strXingHao As String, strVer As String)

    OpenLppx2 strModel

    myVars2.Item("Model").Value = strXingHao
    myVars2.Item("PN").Value = UCase(strModel)
    myVars2.Item("Rev").Value = UCase(strVer)
    myVars2.Item("SN").Value = UCase(Trim(Me.txtHPSN.Text))
    myDoc2.PrintLabel 1
    myDoc2.FormFeed
    UnloadLppx2
    
End Sub

Private Sub UnloadLppx2()
    myApp2.Documents.CloseAll False
    myApp2.Quit
    Set myApp2 = Nothing
End Sub

