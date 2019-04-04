VERSION 5.00
Begin VB.Form frmHW3COM 
   Caption         =   "3COM标签打印"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8820
   StartUpPosition =   2  '屏幕中心
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
      Left            =   3600
      TabIndex        =   5
      Top             =   6600
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
      Left            =   6120
      TabIndex        =   4
      Top             =   6600
      Width           =   1815
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
      Left            =   1200
      TabIndex        =   3
      Top             =   6600
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
      Left            =   3000
      TabIndex        =   1
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.Image Image3 
         Height          =   2640
         Left            =   120
         Picture         =   "frmHW3COM.frx":0000
         Top             =   120
         Width           =   8475
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
      Left            =   2040
      TabIndex        =   2
      Top             =   3240
      Width           =   735
   End
End
Attribute VB_Name = "frmHW3COM"
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
Public unit_key As Long
Public HP_pack_label As Boolean

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
        
        
        HP_pack_label = False
        
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
            'str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtSN.Text) & "'"
            'str = " select top 1 a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "'"
            str = " select top 1 part_number,part_revision,creation_time,order_number,unit_key from (" & _
            "select a.part_number,a.part_revision,a.creation_time,a.unit_key,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "' union " & _
            "select top 1 a.part_number,a.part_revision,a.creation_time,a.unit_key,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
            "where b.original_sn_S = '" & Trim(txtSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
            com.CommandText = str
            rs13.Open Source:=com
            'rs13.Open str
            If rs13.EOF = True Then
                MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
'                Me.cmdCancel.Click
                rs13.Close
                Exit Sub
            Else
                strModel = Mid(Trim(rs13.Fields(0)), 4, 8)
                strVer = rs13.Fields(1)
                unit_key = rs13.Fields(4)
            End If
            If rs13.State = 1 Then
                rs13.Close
            End If
            If con13.State = 1 Then
                con13.Close
            End If
            rs3.Close
'            Exit Sub
        Else
             strModel = Mid(Trim(rs3.Fields(0)), 4, 8)
             strVer = rs3.Fields(1)
        End If
        
        If strModel <> "" And strVer <> "" Then
        '926FEDSDAE704
            '+++++++++++++++++++++
            hpsn = ""
            strIII = ""
            Dim checkhp As New ADODB.Recordset
            If con.State = 0 Then
                con.Open
            End If
            
            sql = "SELECT Label,hp_sn FROM H3C_HP with(NOLOCK) where part_number='" & strModel & "'"
            rec.Open sql, con, adOpenKeyset, adLockOptimistic
            If Not rec.EOF Then
                If rec("Label") = "Yes" Then
                    HP_pack_label = True
                    strIII = rec("hp_sn")
                End If
            End If
            If rec.State = 1 Then rec.Close
            
            If HP_pack_label = True Then
      
                
                sql = "select top 1 component_SN from DC_Component_SN where unit_key = " & CStr(unit_key) & " and remark = 'HP' order by last_modified_time DESC"
                con13.Open
                checkhp.Open sql, con13, adOpenKeyset, adLockOptimistic
                If checkhp.EOF = True Then
                    MsgBox ("没有对应的HP条码！")
                    txtSN.Text = ""
                    txtSN.SetFocus
                    checkhp.Close
                    Exit Sub
                Else
                    hpsn = checkhp.Fields(0)
                    checkhp.Close
                End If
                con13.Close
        
            End If
            '+++++++++++++++++++++
            
            
            
            'strIII = Mid(Trim(txtSN.Text), 5, 3)
            verstring = strVer

            
            Set fs = CreateObject("Scripting.FileSystemObject")

            Dim strDir As String
            strDir = "\\10.11.1.25\Public\Manufacture\标签模板\3Com发货标签\" & strModel & ".lab"
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

'===============add by ben 2012-02-05 start===============
                If reprint = False Then
                    If Connect.isPrintedLabel(Me.txtSN.Text, Me.Name) Then
                        MsgBox ("此序列号已打印！")
                        cmdCancel_SN_Click
                        Exit Sub
                    End If
                End If
'===============add by ben 2012-02-05 end=================

            OpenLppx2 strModel

            myVars2.Item("SN").Value = UCase(snstring)
            myVars2.Item("Rev").Value = UCase(verstring)
            myDoc2.PrintLabel 1
            myDoc2.FormFeed
            
'===============add by ben 2012-02-05 start===============
                Call Connect.addPrintedLabel(Me.txtSN.Text, Me.Name)
'===============add by ben 2012-02-05 end=================

            UnloadLppx2
            '======================================
            
        End If
        If (rs.State = 1) Then
            rs3.Close
        End If
        
        
        
        
        cmdCancel_SN_Click
        
        
        If HP_pack_label = True Then
            frmHW3COM.Hide
    
            FormHPFahuo3COM.txtSN = hpsn
            FormHPFahuo3COM.txtModel_hid = strModel
    
    
            FormHPFahuo3COM.Show
            Call FormHPFahuo3COM.cmdMPrint_Click
        End If
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
