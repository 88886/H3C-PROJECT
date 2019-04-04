VERSION 5.00
Begin VB.Form frmNEC 
   Caption         =   "NEC标签打印"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   LinkTopic       =   "frmNEC"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8745
   StartUpPosition =   2  '屏幕中心
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
      Caption         =   "打印(Print) &p"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "返回(Return)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "取消(Cancel)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "SN："
      BeginProperty Font 
         Name            =   "宋体"
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
         MsgBox "产品序号长度不能小于10!"
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
            MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
            cmdCancel_SN_Click
            rs3.Close
            Exit Sub
        Else
        
            strModel = Mid(Trim(rs3.Fields(0)), 4, 8)
            strVer = rs3.Fields(1)
            

            'strIII = Mid(Trim(txtSN.Text), 5, 3)
            verstring = strVer

            
            Set fs = CreateObject("Scripting.FileSystemObject")

            Dim strDir As String
            strDir = "\\10.11.1.25\Public\Manufacture\标签模板\NEC发货标签\" & strModel & ".lab"
            If Not fs.FileExists(strDir) Then
                MsgBox "没有对应机种打印模板", vbOKOnly + vbExclamation, "警告"
                cmdCancel_SN_Click
                rs3.Close
                Exit Sub
            End If
            
            '==============================


            If verstring = "" Then
                MsgBox ("DS版本未带出，不能打印！")
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
                    MsgBox "此产品序号未进行发货标签软件版本维护!"
                    txtSN.Text = ""
                    txtSN.SetFocus
                    rs3.Close
                    Exit Sub
                Else
                    If rs3.Fields("searchFlag") = "Y" Then
                        Set con2 = New ADODB.Connection
                        con2.CursorLocation = adUseClient
                        con2.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=dataT"
                        con2.ConnectionTimeout = 100
                        
                        sql = "Select top 1  barcode, testtime, remark  FROM test_equ where barcode='" & Trim(txtSN.Text) & "' AND (ISNULL(remark, '') <> '') ORDER BY testtime DESC "
                        If con2.State = 1 Then
                            con2.Close
                        End If
                        con2.Open
                        Set rs2 = New ADODB.Recordset
                        rs2.ActiveConnection = con2
                        rs2.Open sql, con2, adOpenKeyset, adLockOptimistic
                        If rs2.EOF Then
                            MsgBox "查询软件版本资料时错误!"
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
                            Dim enddate As String
                            
                            stmp2 = rs3.Fields("nowVer")
                            stmp3 = rs3.Fields("beforeVer")
                            
                            nowver = Mid(stmp2, 2)
                            beforver = Mid(stmp3, 2)
                            nowver = get_ver(nowver)
                            beforver = get_ver(beforver)
                            
                            enddate = rs3.Fields("endDate")
                            
                            stmp = rs2.Fields("remark")
                            
                            If InStr(stmp, nowver) > 0 Then
                                Dim ttt As String
                                ttt = get_nextchar(stmp, nowver)
                                
                                If ttt = "L" Or ttt = "P" Then
                                    MsgBox "查询软件版本资料时错误(版本匹配错误)!"
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
                                    MsgBox "查询软件版本资料时错误(版本匹配错误)!"
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
                                        MsgBox "查询软件版本资料时错误(版本匹配错误)!"
                                        txtSN.Text = ""
                                        txtSN.SetFocus
                                        rs2.Close
                                        rs3.Close
                                        Exit Sub
                                    Else
                                        If DateDiff("d", Now, CDate(enddate)) < 0 Then
                                            MsgBox "查询软件版本资料时错误(超过有效期)!"
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
                                        MsgBox "查询软件版本资料时错误(版本匹配错误)!"
                                        txtSN.Text = ""
                                        txtSN.SetFocus
                                        rs2.Close
                                        rs3.Close
                                        Exit Sub
                                End If
                                    '**********
                                
                                End If
                                
                            End If
                            
                        End If
                        rs2.Close
                        con2.Close
                        
                    Else
                        If rs3.Fields("searchFlag") = "N" Then
    '=====================================================================
    sql = "select top 1 ver from version where SN='" & txtSN.Text & "' order by testtime desc"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         MsgBox "此产品序号未收集版本!"
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
           MsgBox "此产品序号未收集版本!"
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
             txtVer.Text = rs8.Fields(0)
          Else
             MsgBox "此产品序号未收集版本!"
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
                End If
                
                rs3.Close
                con1.Close
                
                '--------------


            OpenLppx2 strModel

            myVars2.Item("SN").Value = snstring
            myVars2.Item("Rev").Value = verstring
            myVars2.Item("soft").Value = txtVer.Text

            myDoc2.PrintLabel 1
            myDoc2.FormFeed
            UnloadLppx2
            '======================================
            
        End If
        rs3.Close
        
        
        
        cmdCancel_SN_Click
        
        
   End If
End Sub

Private Sub OpenLppx2(model As String)
    Me.MousePointer = vbHourglass

    Set myDoc2 = myApp2.Documents.Open("\\10.11.1.25\Public\Manufacture\标签模板\3Com发货标签\" & model & ".lab")
    
    Me.MousePointer = vbDefault
    Set myVars2 = myDoc2.Variables
    Set myObjs2 = myDoc2.DocObjects
End Sub

Private Sub UnloadLppx2()
    myApp2.Documents.CloseAll False
    myApp2.Quit
    Set myApp2 = Nothing
End Sub
