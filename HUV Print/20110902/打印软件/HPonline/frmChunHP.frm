VERSION 5.00
Begin VB.Form frmChunHP 
   Caption         =   "纯HP发货在线打印"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9210
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtHPSN 
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
      Left            =   2520
      TabIndex        =   10
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtPN 
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Picture         =   "frmChunHP.frx":0000
         Top             =   240
         Width           =   4320
      End
      Begin VB.Image Image3 
         Height          =   1815
         Left            =   120
         Picture         =   "frmChunHP.frx":27D5
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
      Caption         =   "HP SN："
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
      Left            =   1200
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "产品机种："
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
      Left            =   720
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "产品编号："
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
      Left            =   720
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "产品UPC："
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
      Left            =   720
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Caption         =   "产品描述："
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
      Left            =   720
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmChunHP"
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
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
    Me.MousePointer = vbHourglass

    Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "HP发货标签NEW.lab")
    
    Me.MousePointer = vbDefault
    Set myVars = myDoc.Variables
    Set myObjs = myDoc.DocObjects
End Sub


Private Sub cmdPrint_HPSN_Click()

    If txtHPSN.Text = "" Then
        MsgBox ("序列号未输入，不能打印！")
        txtSN.SetFocus
        Exit Sub
    End If

    If txtProduct.Text = "" Then
        MsgBox ("产品编码未带出，不能打印！")
        Exit Sub
    End If
    If txtDesc.Text = "" Then
        MsgBox ("产品描述未带出，不能打印！")
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
        
        
        '==================
        Dim con As ADODB.Connection
        Dim rs3 As ADODB.Recordset

        Set con = New ADODB.Connection
        Set rs3 = New ADODB.Recordset
        con.ConnectionString = "driver={sql server};server=sz-sql01;uid=sa;pwd=Itadmin1;Database=Print"
        con.ConnectionTimeout = 50
        con.Open
        Dim str As String
        Set rs3.ActiveConnection = con
        rs3.CursorType = adOpenDynamic
        
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
            MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
            cmdCancel_HPSN_Click
            rs3.Close
            Exit Sub
        Else
        
            
            strModel = Mid(Trim(rs3.Fields(0)), 4, 8)
            strVer = rs3.Fields(1)
            strIII = Mid(Trim(txtHPSN.Text), 5, 3)
            

            Set fs = CreateObject("Scripting.FileSystemObject")
            'Dim fs As New FileSystemObject

            Dim strDir As String
            strDir = "\\sz-fs01\Public\Manufacture\标签模板\Phase3 HP认证发货标签\" & strModel & ".lab"
            If Not fs.FileExists(strDir) Then
                MsgBox "没有对应机种打印模板", vbOKOnly + vbExclamation, "警告"
                cmdCancel_HPSN_Click
                rs3.Close
                Exit Sub
            End If
            
            
            Dim rec As New ADODB.Recordset
            sql = "select * from HP where h3c_bom_code='" & strModel & "' and hp_sn_iii='" & strIII & "'"
            If conn.State = 0 Then
                conn.ConnectionString = Connect.getConnectionstring
                conn.Open
            End If
            rec.Open sql, conn, adOpenKeyset, adLockOptimistic
            If rec.EOF = False Then
                
                
                '====================
                 If IsNull(rec.Fields("hp_desc1")) Then
                    MsgBox ("此序列号未维护描述信息!")
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
                    MsgBox ("此序列号未维护产品编码!")
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
                    MsgBox ("此序列号未维护产品型号!")
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
                

                cmdPrint_HPSN_Click
                
                cmdPrint_Model_Click strModel, str2, strVer
                
                '======================
                
                
            Else
                MsgBox "此产品序号未收集版本!"
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

    Set myDoc2 = myApp2.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\Phase3 HP认证发货标签\" & model & ".lab")
    
    Me.MousePointer = vbDefault
    Set myVars2 = myDoc2.Variables
    Set myObjs2 = myDoc2.DocObjects
End Sub

Private Sub cmdPrint_Model_Click(strModel As String, strXingHao As String, strVer As String)

    OpenLppx2 strModel

    myVars2.Item("Model").Value = strXingHao
    myVars2.Item("PN").Value = UCase(strModel)
    myVars2.Item("Rev").Value = UCase(strVer)
   
    myDoc2.PrintLabel 1
    myDoc2.FormFeed
    UnloadLppx2
    
End Sub

Private Sub UnloadLppx2()
    myApp2.Documents.CloseAll False
    myApp2.Quit
    Set myApp2 = Nothing
End Sub

