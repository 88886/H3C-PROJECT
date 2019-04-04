VERSION 5.00
Begin VB.Form F74CAA0190116004 
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   Picture         =   "mainForm.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   13950
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   375
      Left            =   11520
      TabIndex        =   14
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   6045
      Width           =   1815
   End
   Begin VB.TextBox tbGP 
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox tbGTIN 
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox tbASR 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox tbPart 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox tbMAC 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox tbSN 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "GP型号"
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "GTIN型号"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "ASR型号"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "机种名称"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "MAC流水号"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "SN流水号"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   3240
      Width           =   975
   End
End
Attribute VB_Name = "F74CAA0190116004"
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
Dim projectConfig(25, 4) As String
Dim connFTPC As New ADODB.Connection


Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
    
End Sub

Private Sub OpenUnitLppx()
   Me.MousePointer = vbHourglass
   'ASR6026PF,ASR6126PF,GP1026PF-AC,GP1126PF-AC
   'ASR6026PF-DC,ASR6126PF-DC,GP1026PF-DC,GP1126PF-DC

   If Me.tbPart.Text = "ASR6026PF" Or Me.tbPart.Text = "ASR6026PF-BB" Or tbPart.Text = "ASR6126PF" Or tbPart.Text = "ASR6126PF-BB" Or tbPart.Text = "GP1026PF-AC" Or tbPart.Text = "GP1126PF-AC" Or tbPart.Text = "GP1226PF-AC" Or tbPart.Text = "GP1326PF-AC" Or tbPart.Text = "GP1226HFPF-AC" Or tbPart.Text = "GP1326HFPF-AC" Or Me.tbPart.Text = "ASR7024PF-AC" Or Me.tbPart.Text = "ASR7024EPF-AC" Or Me.tbPart.Text = "ASR8048PF-AC" Or Me.tbPart.Text = "ASR8048EPF-AC" Or Me.tbPart.Text = "ASR8024PF-AC" Or Me.tbPart.Text = "ASR8024EPF-AC" Or Me.tbPart.Text = "ASR7048PF-AC" Or Me.tbPart.Text = "ASR7048EPF-AC" Then
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\BU1小项目标签\" & "PFUnitlabel-AC.Lab")
   ElseIf Me.tbPart.Text = "ASR6026PF-DC" Or tbPart.Text = "ASR6126PF-DC" Or tbPart.Text = "GP1026PF-DC" Or tbPart.Text = "GP1126PF-DC" Then
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\BU1小项目标签\" & "PFUnitlabel-DC.Lab")
   Else
        MsgBox "当前机种不在打印列表内", vbInformation + vbOKOnly, "机种不在打印列表内"
        Exit Sub
        Unload Me
   End If
   
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub
Private Sub OpenBoxLppx()
   Me.MousePointer = vbHourglass
   If Me.tbPart.Text = "ASR6026PF" Or Me.tbPart.Text = "ASR6026PF-BB" Or tbPart.Text = "ASR6126PF" Or tbPart.Text = "ASR6126PF-BB" Or tbPart.Text = "GP1026PF-AC" Or tbPart.Text = "GP1126PF-AC" Or tbPart.Text = "GP1226PF-AC" Or tbPart.Text = "GP1326PF-AC" Or tbPart.Text = "GP1226HFPF-AC" Or tbPart.Text = "GP1326HFPF-AC" Or Me.tbPart.Text = "ASR7024PF-AC" Or Me.tbPart.Text = "ASR7024EPF-AC" Or Me.tbPart.Text = "ASR8048PF-AC" Or Me.tbPart.Text = "ASR8048EPF-AC" Or Me.tbPart.Text = "ASR8024PF-AC" Or Me.tbPart.Text = "ASR8024EPF-AC" Or Me.tbPart.Text = "ASR7048PF-AC" Or Me.tbPart.Text = "ASR7048EPF-AC" Or Me.tbPart.Text = "ASR6326PF-AC" Or Me.tbPart.Text = "ASR6226PF-AC" Then
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\BU1小项目标签\" & "PFBoxlabel-AC.Lab")
   ElseIf Me.tbPart.Text = "ASR6026PF-DC" Or tbPart.Text = "ASR6126PF-DC" Or tbPart.Text = "GP1026PF-DC" Or tbPart.Text = "GP1126PF-DC" Then
        Set myDoc = myApp.Documents.Open("\\sz-fs01\Public\Manufacture\标签模板\BU1小项目标签\" & "PFBoxlabel-DC.Lab")
   Else
        MsgBox "当前机种不在打印列表内", vbInformation + vbOKOnly, "机种不在打印列表内"
        Exit Sub
        Unload Me
   End If
   
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub

Private Sub cmdCancel_Click()
    Dim Ctr As Control
    For Each Ctr In Me.Controls
        If TypeOf Ctr Is TextBox Then
            Ctr.Text = ""
        End If
    Next
End Sub

Private Sub cmdExit_Click()
    Unload Me
    UnloadLppx
End Sub

Private Sub cmdPrint_Click()
   If (tbSN.Text = "" Or tbMAC.Text = "" Or tbPart.Text = "" Or tbASR.Text = "" Or tbGTIN.Text = "") Then
      MsgBox "流水码SN或者料号或者ASR料号或者GTIN的料号不为空", vbInformation + vbOKOnly, "打印信息不能为空"
      Exit Sub
   End If

   cmdPrint.Caption = "执行中..."
   cmdPrint.Enabled = False
   If (tbPart.Text = "ASR6026PF" Or tbPart.Text = "ASR6026PF-BB" Or tbPart.Text = "ASR6126PF" Or tbPart.Text = "ASR6126PF-BB" Or tbPart.Text = "ASR6026PF-DC" Or tbPart.Text = "ASR6126PF-DC" Or Me.tbPart.Text = "ASR7024PF-AC" Or Me.tbPart.Text = "ASR7024PF-AC" Or Me.tbPart.Text = "ASR7024EPF-AC" Or Me.tbPart.Text = "ASR8048PF-AC" Or Me.tbPart.Text = "ASR8048EPF-AC" Or Me.tbPart.Text = "ASR8024PF-AC" Or Me.tbPart.Text = "ASR8024EPF-AC" Or Me.tbPart.Text = "ASR7048PF-AC" Or Me.tbPart.Text = "ASR7048EPF-AC" Or Me.tbPart.Text = "ASR6326PF-AC" Or Me.tbPart.Text = "ASR6226PF-AC") Then
    OpenBoxLppx
'        upper case the sn and mac
        myVars.Item("MAC").Value = UCase(tbMAC.Text)
         myVars.Item("Model").Value = tbASR.Text
       ' myVars.Item("Model").Value = tbPart.Text
        myVars.Item("GTN").Value = Me.tbGTIN.Text
        myVars.Item("SN").Value = UCase(tbSN.Text)
        myDoc.PrintLabel 1
        myDoc.FormFeed
    End If
    
    If (tbPart.Text = "GP1026PF-AC" Or tbPart.Text = "GP1126PF-AC" Or tbPart.Text = "GP1026PF-DC" Or tbPart.Text = "GP1126PF-DC" Or tbPart.Text = "GP1226PF-AC" Or tbPart.Text = "GP1326PF-AC" Or tbPart.Text = "GP1226HFPF-AC" Or tbPart.Text = "GP1326HFPF-AC") Then
    'upper case the sn and mac
    OpenUnitLppx
        myVars.Item("MAC").Value = UCase(tbMAC.Text)
        myVars.Item("Model").Value = tbASR.Text
        myVars.Item("SN").Value = UCase(tbSN.Text)
        myDoc.PrintLabel 1
        myDoc.FormFeed
    OpenBoxLppx
        myVars.Item("MAC").Value = UCase(tbMAC.Text)
        myVars.Item("Model").Value = tbASR.Text
        myVars.Item("GTN").Value = getGTN(Me.tbASR.Text)
        myVars.Item("SN").Value = UCase(tbSN.Text)
        myDoc.PrintLabel 1
        myDoc.FormFeed
    OpenUnitLppx
        myVars.Item("MAC").Value = UCase(tbMAC.Text)
        myVars.Item("Model").Value = Me.tbGP.Text
        myVars.Item("SN").Value = UCase(tbSN.Text)
'        myVars.Item("GTN").Value = ""
        myDoc.PrintLabel 1
        myDoc.FormFeed
    OpenBoxLppx
        myVars.Item("MAC").Value = UCase(tbMAC.Text)
        myVars.Item("Model").Value = Me.tbGP.Text
        myVars.Item("GTN").Value = tbGTIN.Text
        myVars.Item("SN").Value = UCase(tbSN.Text)
        myDoc.PrintLabel 1
        myDoc.FormFeed
    
    End If
    
   UnloadLppx
   cmdCancel_Click
   Me.tbSN.SetFocus
   
   cmdPrint.Caption = "打印(Print) &p"
   cmdPrint.Enabled = True
End Sub

Private Sub Form_Load()
    Dim temp As Variant
'    ASR6026PF;ASR6026-AC;;735000801112;
'    ASR6026PF-BB;ASR6026-AC;;735000801112;
'    ASR6126PF;ASR6126-AC;;735000801114;
'    ASR6126PF-BB;ASR6126-AC;;735000801114;
'    ASR6226PF;ASR6226-AC;;735000801127
'    ASR6326PF;ASR6326-AC;;735000801130
'
'    GP1026PF-AC;ASR6026-AC;MS4026-AC;735000801119;
'    GP1126PF-AC;ASR6126-AC;MS4126-AC;735000801120;
'    ASR6026PF-DC;ASR6026-DC;;735000801113;
'    ASR6126PF-DC;ASR6126-DC;;735000801115;
'    GP1026PF-DC;ASR6026-DC;MS4026-DC;735000801234;
'    GP1126PF-DC;ASR6126-DC;MS4126-DC;735000801235;
'    GP1226PF-AC;ASR6226-AC;MS4226-AC;7350008011222
'    GP1326PF-AC;ASR6326-AC;MS4326-AC;735000801125

    temp = Split("ASR6026PF;ASR6026-AC;;735000801112;" & _
    "ASR6026PF-BB;ASR6026-AC;;735000801112;" & _
    "ASR6126PF;ASR6126-AC;;735000801114;" & _
    "ASR6126PF-BB;ASR6126-AC;;735000801114;" & _
    "ASR6226PF;ASR6226-AC;;735000801530;" & _
    "ASR6326PF;ASR6326-AC;;735000801531;" & _
    "ASR6226PF-AC;ASR6226-AC;;735000801530;" & _
    "ASR6326PF-AC;ASR6326-AC;;735000801531;" & _
    "GP1026PF-AC;ASR6026-AC;MS4026-AC;735000801119;" & _
    "GP1126PF-AC;ASR6126-AC;MS4126-AC;735000801120;" & _
    "ASR7024PF-AC;ASR7024-AC;;735000801585;" & _
    "ASR7024EPF-AC;ASR7024E-AC;;735000801584;" & _
    "ASR8048PF-AC;ASR8048-AC;;735000801597;" & _
    "ASR8048EPF-AC;ASR8048E-AC;;735000801596;" & _
    "ASR8024PF-AC;ASR8024-AC;;735000801593;" & _
    "ASR8024EPF-AC;ASR8024E-AC;;735000801592;" & _
    "ASR7048PF-AC;ASR7048-AC;;735000801589;" & _
    "ASR7048EPF-AC;ASR7048E-AC;;735000801590;" & _
    "GP1026PF-DC;ASR6026-DC;MS4026-DC;735000801234;" & _
    "GP1126PF-DC;ASR6126-DC;MS4126-DC;735000801235;" & _
    "GP1226PF-AC;ASR6226-AC;MS4226-AC;735000801532;" & _
    "GP1326PF-AC;ASR6326-AC;MS4326-AC;735000801533;" & _
    "ASR6026PF-DC;ASR6026-DC;;735000801113;" & _
    "GP1226HFPF-AC;ASR6226-AC;MS4226-AC;735000801532;" & _
    "ASR6126PF-DC;ASR6126-DC;;735000801115;", ";")
    
    
    
    '  "GP1326HFPF-AC;ASR6326-AC;MS4326-AC;735000801533"





    Dim i, j, k As Integer
    i = 0
    For j = 0 To 24
        For k = 0 To 3
            projectConfig(j, k) = temp(i)
            i = i + 1
        Next k
    Next j
    If connFTPC.State = 0 Then
      connFTPC.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
      connFTPC.Open
   End If
End Sub

Private Sub tbMAC_Change()
 cmdPrint_Click
End Sub

'Private Sub tbMAC_KeyPress(KeyAscii As Integer)
 '   If (KeyAscii = 13) Then
  '      If (Len(tbMAC.Text) = 12 And Right(tbMAC.Text, 1) = "0") Then
   '         cmdPrint_Click
    '    Else
     '       MsgBox "MAC地址不符合规则", vbOKOnly + vbExclamation, "警告"
      '      tbMAC.SetFocus
       '     Exit Sub
        'End If
   ' End If
    
'End Sub

Private Sub tbSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Dim rs13 As ADODB.Recordset
        Dim com As ADODB.Command
        Set rs13 = New ADODB.Recordset
        Set com = New ADODB.Command
        com.ActiveConnection = connFTPC
        'str = " select top 1 part_number,part_revision,creation_time from unit wint(nolock) where serial_number='" & Trim(txtSN.Text) & "'"
        'str = " select top 1 a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(txtSN.Text) & "'"
        sql = " select top 1 part_number,part_revision,creation_time,order_number from (" & _
        "select a.part_number,a.part_revision,a.creation_time,b.order_number from unit a with(NOLOCK) left outer join work_order b with(NOLOCK) on a.order_key=b.order_key where a.serial_number='" & Trim(tbSN.Text) & "' union " & _
        "select top 1 a.part_number,a.part_revision,a.creation_time,c.order_number from UNIT a left join UDA_Unit b on a.unit_key = b.object_key left join WORK_ORDER c on a.order_key = c.order_key " & _
        "where b.original_sn_S = '" & Trim(tbSN.Text) & "' and b.order_type_S = 'TASK') as t order by t.creation_time desc "
        com.CommandText = sql
        rs13.Open Source:=com
        'rs13.Open str
       If rs13.EOF = True Then
           MsgBox "没有对应机种版本信息", vbOKOnly + vbExclamation, "警告"
            rs13.Close
            connFTPC.Close
           cmdCancel_Click
           tbSN.SetFocus
            Exit Sub
        Else
            'asr,gp,gtin
            tbPart.Text = rs13.Fields(0)
            Dim i As Integer
            For i = 0 To 24
                If projectConfig(i, 0) = tbPart.Text Then
                    tbASR.Text = projectConfig(i, 1)
                    tbGP.Text = projectConfig(i, 2)
                    tbGTIN.Text = projectConfig(i, 3)
                End If
            Next i
        End If
        If rs13.State = 1 Then
            rs13.Close
        End If
        Me.tbASR.Enabled = False
        Me.tbGP.Enabled = False
        Me.tbGTIN.Enabled = False
        Me.tbPart.Enabled = False
        Me.tbMAC.SetFocus
    End If
End Sub
Function getGTN(ASRModel As String) As String
    Dim j As Integer
    For j = 0 To 24
        If projectConfig(j, 1) = ASRModel And Left$(projectConfig(j, 0), 3) = "ASR" Then
            getGTN = projectConfig(j, 3)
            Exit For
        End If
    Next j
End Function

