VERSION 5.00
Begin VB.Form frmNEC 
   Caption         =   "NEC��ǩ��ӡ"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   LinkTopic       =   "frmNEC"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8745
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtVer 
      Height          =   450
      Left            =   8280
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ӡ(Print) &p"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����(Return)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(Cancel)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox txtSN 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.Image Image3 
         Height          =   2235
         Left            =   120
         Picture         =   "frmNEC.frx":0000
         Top             =   120
         Width           =   8310
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
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
      Left            =   1800
      TabIndex        =   2
      Top             =   2880
      Width           =   735
   End
End
Attribute VB_Name = "frmNEC"
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
Dim rec As New ADODB.Recordset
Dim myApp2 As New LabelManager2.Application
Dim myDoc2 As LabelManager2.Document
Dim myVars2 As LabelManager2.Variables
Dim myObjs2 As LabelManager2.DocObjects

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

Private Sub cmdCancel_SN_Click()
txtSN.Text = ""
txtSN.SetFocus

End Sub
Public Function get_nextchar(strRemark As String, pipei As String) As String

    If InStr(strRemark, pipei) > 0 Then
        get_nextchar = UCase(Mid(strRemark, InStr(strRemark, pipei) + Len(pipei), 1))
    End If

End Function

Public Function get_ver(strVer As String) As String

    If InStr(strVer, "-") > 1 Then
        get_ver = Mid(strVer, 1, InStr(strVer, "-") - 1)
    Else
        get_ver = strVer
    End If
    

End Function

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If Len(txtSN.Text) < 10 Then
            MsgBox "��Ʒ��ų��Ȳ���С��10!"
            txtSN.SetFocus
            Exit Sub
        End If
        
        Dim snstring As String
        Dim verstring As String
        snstring = Trim(Me.txtSN.Text)
        
        Dim strModel As String
        Dim strVer As String
        Dim strIII As String
        Dim str2 As String
        
        '==================
        Dim con As ADODB.Connection
        Dim rs3 As ADODB.Recordset

        Set con = New ADODB.Connection
        Set rs3 = New ADODB.Recordset
        con.ConnectionString = "Provider=SQLOLEDB;User ID=datasweep;PWD=datasweep;Initial Catalog=dsActive;Data Source=DS-DB"
        con.ConnectionTimeout = 50
        con.Open
        Dim str As String
        Set rs3.ActiveConnection = con
        rs3.CursorType = adOpenDynamic
        
        str = " select top 1 part_number,part_revision,creation_time from (" & _
            " select part_number,part_revision,creation_time from dsactive.dbo.unit nolock " & _
            " where serial_number='" & snstring & "'" & _
            " union" & _
            " select part_number,part_rev as part_revision,creation_time from dsactive.dbo.dc_task_order NOLOCK  " & _
            " where order_number=(select order_number from dsactive.dbo.taskorder_unit NOLOCK" & _
            " where serial_number='" & snstring & "')" & _
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
'            str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtSN.Text) & "'"
             str = " select top 1 part_number,part_revision,creation_time,order_number from (" & _
            "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
            "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
            "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
            com.CommandText = str
            rs13.Open Source:=com
            'rs13.Open str
            If rs13.EOF = True Then
                cmdCancel_SN_Click
                rs13.Close
                rs3.Close
                MsgBox "û�ж�Ӧ���ְ汾��Ϣ", vbOKOnly + vbExclamation, "����"
                Exit Sub
            Else
                strModel = Mid(Trim(rs13.Fields(0)), 4, 8)
                strVer = rs13.Fields(1)
                verstring = strVer
            End If
            If rs13.State = 1 Then
                rs13.Close
            End If
            If con13.State = 1 Then
                con13.Close
            End If
'            MsgBox "û�ж�Ӧ���ְ汾��Ϣ", vbOKOnly + vbExclamation, "����"
'            cmdCancel_SN_Click
'            rs3.Close
'            Exit Sub
        Else
            strModel = Mid(Trim(rs3.Fields(0)), 4, 8)
            strVer = rs3.Fields(1)
            'strIII = Mid(Trim(txtSN.Text), 5, 3)
            verstring = strVer
        End If
        rs3.Close
        Set fs = CreateObject("Scripting.FileSystemObject")
            Dim strDir As String
            strDir = "\\10.11.1.25\Public\Manufacture\��ǩģ��\NEC������ǩ\" & strModel & ".lab"
            If Not fs.FileExists(strDir) Then
                MsgBox "û�ж�Ӧ���ִ�ӡģ��", vbOKOnly + vbExclamation, "����"
                cmdCancel_SN_Click
                rs3.Close
                Exit Sub
            End If
            
            '==============================
            If verstring = "" Then
                MsgBox ("DS�汾δ���������ܴ�ӡ��")
                Exit Sub
            End If
            
            '--------------
            Set con1 = New ADODB.Connection
            con1.CursorLocation = adUseClient
            con1.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
            con1.ConnectionTimeout = 100
                
                
            sql = "select * from tblSoftVersion where model='" & strModel & "'"
            
            If con1.State = 1 Then
                con1.Close
            End If
            con1.Open
            Set rs3 = New ADODB.Recordset
            rs3.ActiveConnection = con
            rs3.Open sql, con1, adOpenKeyset, adLockOptimistic
            If rs3.EOF Then
                MsgBox "�˲�Ʒ���δ���з�����ǩ����汾ά��!"
                txtSN.Text = ""
                txtSN.SetFocus
                rs3.Close
                Exit Sub
            Else
                If rs3.Fields("searchFlag") = "Y" And rs3.Fields("beforeVer") <> "/" And rs3.Fields("nowVer") <> "/" Then
                    Set con2 = New ADODB.Connection
                    con2.CursorLocation = adUseClient
                    con2.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=dataT"
                    con2.ConnectionTimeout = 100
                        
                    sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (rtrim(remark) <> '' and remark is not null AND testtime >= dateadd(month,-1,getdate())) ORDER BY testtime DESC "
'                    sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (ISNULL(remark, '') <> '') ORDER BY testtime DESC "
                    If con2.State = 1 Then
                        con2.Close
                    End If
                    con2.Open
                    Set rs2 = New ADODB.Recordset
                    rs2.ActiveConnection = con2
                    rs2.Open sql, con2, adOpenKeyset, adLockOptimistic
                    If rs2.EOF Then
                        MsgBox "��ѯ����汾����ʱ����!"
                        txtSN.Text = ""
                        txtSN.SetFocus
                        rs2.Close
                        rs3.Close
                        Exit Sub
                    Else
                        Dim stmp As String
                        Dim stmp2 As String
                        Dim stmp3 As String
                        Dim nowver As String
                        Dim beforver As String
                        Dim endDate As String
                        
                        stmp2 = rs3.Fields("nowVer")
                        stmp3 = rs3.Fields("beforeVer")
                        
                        nowver = Mid(stmp2, 2)
                        beforver = Mid(stmp3, 2)
                        nowver = get_ver(nowver)
                        beforver = get_ver(beforver)
                        
                        endDate = rs3.Fields("endDate")
                        
                        stmp = rs2.Fields("remark")
                        
                        If InStr(stmp, nowver) > 0 Then
                            Dim ttt As String
                            ttt = get_nextchar(stmp, nowver)
                            
                            If ttt = "L" Or ttt = "P" Then
                                MsgBox "��ѯ����汾����ʱ����(�汾ƥ�����)!"
                                txtSN.Text = ""
                                txtSN.SetFocus
                                rs2.Close
                                rs3.Close
                                Exit Sub
                            Else
                                txtVer.Text = stmp2
                            End If
                            
                        Else
                            If Trim(beforver) = "" Then
                                MsgBox "��ѯ����汾����ʱ����(�汾ƥ�����)!"
                                txtSN.Text = ""
                                txtSN.SetFocus
                                rs2.Close
                                rs3.Close
                                Exit Sub
                            Else
                                 '***********
                                 
                                If InStr(stmp, beforver) > 0 Then
                                    Dim st As String
                                    st = get_nextchar(stmp, beforver)
                                    If st = "L" Or st = "P" Then
                                        MsgBox "��ѯ����汾����ʱ����(�汾ƥ�����)!"
                                        txtSN.Text = ""
                                        txtSN.SetFocus
                                        rs2.Close
                                        rs3.Close
                                        Exit Sub
                                    Else
                                        If DateDiff("d", Now, CDate(endDate)) < 0 Then
                                            MsgBox "��ѯ����汾����ʱ����(������Ч��)!"
                                            txtSN.Text = ""
                                            txtSN.SetFocus
                                            rs2.Close
                                            rs3.Close
                                            Exit Sub
                                        Else
                                            txtVer.Text = stmp3
                                        End If
                                    End If
                                Else
                                    MsgBox "��ѯ����汾����ʱ����(�汾ƥ�����)!"
                                    txtSN.Text = ""
                                    txtSN.SetFocus
                                    rs2.Close
                                    rs3.Close
                                    Exit Sub
                                End If
                                '**********
                            End If
                            
                        End If
                        
                    End If      'If rs2.EOF Then
                    rs2.Close
                    con2.Close
                ElseIf rs3.Fields("beforeVer") = "/" Or rs3.Fields("nowVer") = "/" Then
                    txtVer.Text = "/"
                Else            'If rs3.Fields("searchFlag") = "Y" Then
                    If rs3.Fields("searchFlag") = "N" Then
'=====================================================================
                        Dim stmp2_2 As String
                        Dim stmp3_2 As String
                        Dim endDate_2 As String
                        Dim nowver_2 As String
                        Dim beforver_2 As String
                        Dim stmp_2 As String
                        
                        stmp2_2 = rs3.Fields("nowVer")
                        stmp3_2 = rs3.Fields("beforeVer")
                        endDate_2 = rs3.Fields("endDate")
                        nowver_2 = Trim(stmp2_2)
                        beforver_2 = Trim(stmp3_2)

                        sql = "select top 1 ver from version where SN='" & txtSN.Text & "' order by testtime desc"
                        rec.Open sql, conn, adOpenKeyset, adLockOptimistic
                        If rec.EOF = True Then
                            MsgBox "�˲�Ʒ���δ�ռ��汾!"
                            txtSN.Text = ""
                            txtSN.SetFocus
                            rec.Close
                            rs3.Close
                            Exit Sub
                        Else
                            Dim rcd As New ADODB.Recordset
                            sql = "select max(testtime) from version where sn='" & Trim(txtSN.Text) & "'"
                            rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
                            If rcd.EOF = True Then
                                MsgBox "�˲�Ʒ���δ�ռ��汾!"
                                txtSN.Text = ""
                                txtSN.SetFocus
                                rcd.Close
                                rec.Close
                                rs3.Close
                                Exit Sub
                            Else
                                Dim rs8 As New ADODB.Recordset
                                sql = "select ver from version where testtime='" & rcd.Fields(0) & "' and sn='" & Trim(txtSN.Text) & "'"
                                rs8.Open sql, conn, adOpenKeyset, adLockOptimistic
                                If rs8.EOF = False Then
'                                   txtVer.Text = rs8.Fields(0)
                                        stmp_2 = rs8.Fields(0)
                                        If checkVersion(stmp_2, beforver_2, nowver_2, endDate_2) Then
                                            txtVer.Text = rs8.Fields(0)
                                        Else
                                            txtSN.Text = ""
                                            txtSN.SetFocus
                                            rs8.Close
                                            rcd.Close
                                            rec.Close
                                            rs3.Close
                                            Exit Sub
                                        End If
                                Else
                                   MsgBox "�˲�Ʒ���δ�ռ��汾!"
                                   txtSN.Text = ""
                                   txtSN.SetFocus
                                   rs8.Close
                                   rcd.Close
                                   rec.Close
                                   rs3.Close
                                   Exit Sub
                                End If
                                rs8.Close
                            End If
                            rcd.Close
                        End If
                        rec.Close
  '==============================================
                    End If
                End If
            End If      'If rs3.Fields("searchFlag") = "Y" Then
            rs3.Close
            con1.Close
            
'===============add by ben 2012-02-05 start===============
                If reprint = False Then
                    If Connect.isPrintedLabel(Me.txtSN.Text, Me.Name) Then
                        MsgBox ("�����к��Ѵ�ӡ��")
                        cmdCancel_SN_Click
                        Exit Sub
                    End If
                End If
'===============add by ben 2012-02-05 end=================
            
            '--------------
            OpenLppx2 strModel
            myVars2.Item("SN").Value = UCase(snstring)
            myVars2.Item("Rev").Value = "R" & UCase(verstring)
            myVars2.Item("soft").Value = txtVer.Text
            myDoc2.PrintLabel 1
            myDoc2.FormFeed
            
'===============add by ben 2012-02-05 start===============
                Call Connect.addPrintedLabel(Me.txtSN.Text, Me.Name)
'===============add by ben 2012-02-05 end=================
            
            UnloadLppx2
        '======================================
        'End If
'
        cmdCancel_SN_Click
    End If
End Sub

Private Sub OpenLppx2(model As String)
    Me.MousePointer = vbHourglass

    Set myDoc2 = myApp2.Documents.Open("\\10.11.1.25\Public\Manufacture\��ǩģ��\NEC������ǩ\" & model & ".lab")
    
    Me.MousePointer = vbDefault
    Set myVars2 = myDoc2.Variables
    Set myObjs2 = myDoc2.DocObjects
End Sub

Private Sub UnloadLppx2()
    myApp2.Documents.CloseAll False
    myApp2.Quit
    Set myApp2 = Nothing
End Sub
