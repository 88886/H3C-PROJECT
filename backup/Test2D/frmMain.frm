VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Test 2D"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11715
   LinkTopic       =   "2D Test"
   ScaleHeight     =   6585
   ScaleWidth      =   11715
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "高端模块二维码标签验证"
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects

Private Sub Command1_Click()
   OpenLppx
   myVars.Item("Rev").Value = "T0"
   myVars.Item("SN").Value = "21XXXXXXXXXXXXXXXXX"
   myVars.Item("Type").Value = "Test-Test-B0"
   myVars.Item("Rohs").Value = "Y*"
   myDoc.PrintLabel 1
   myDoc.FormFeed
   UnloadLppx
End Sub

Private Sub UnloadLppx()
    myApp.Documents.CloseAll False    '在文档集上使用CloseAll方法来关闭所有文档
    myApp.Quit
    Set myApp = Nothing
End Sub

Private Sub OpenLppx()
   Me.MousePointer = vbHourglass
   
   'sql99 = "select order_number from work_order  where order_key in (select order_key from unit where serial_number='" & txtSN.Text & "')"
   'rec1.Open sql99, conn1, adOpenKeyset, adLockOptimistic
   'If rec1.EOF = True Then
   '     Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "H3C.lab")
   'Else
   '     If Trim(rec1.Fields(0) > "30000000") And Trim(rec1.Fields(0) < "40000000") Then
   '         Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "NPI-H3C.lab")
   '     Else
   '         Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "H3C.lab")
   '     End If
   'End If
   'rec1.Close

   Set myDoc = myApp.Documents.Open("\\Sz-fs01\Public\Manufacture\标签模板\" & "高端模块二维码标签.lab")
   
   Me.MousePointer = vbDefault
   Set myVars = myDoc.Variables
   Set myObjs = myDoc.DocObjects
End Sub
