VERSION 5.00
Begin VB.Form NXPrint 
   Caption         =   "逆向装箱清单打印"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   Icon            =   "Packlist - Copy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10335
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "打印"
      Height          =   615
      Left            =   7560
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtModel 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7080
      TabIndex        =   4
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtVersion 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7080
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txtSN 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7080
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   8535
      Left            =   0
      Picture         =   "Packlist - Copy.frx":0ECA
      ScaleHeight     =   8475
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "版本:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "机种:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SN:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "NXPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myObjs As LabelManager2.DocObjects
Dim con As ADODB.Connection
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim str As String
Dim com As ADODB.Command
Dim status As String
Dim dateStr As String



Private Sub Command1_Click()
       If Trim(txtVersion.Text) = "" Then
         MsgBox "请输入机种版本!"
         txtVersion.SetFocus
         Exit Sub
        End If
      dateStr = Format(Now, "yyyy-mm-dd")
      Dim spath As String
     ' spath = "\\10.11.1.25\Public\Manufacture\逆向标签模板\逆向装箱清单\" + dateStr + "\" + Trim(txtModel.Text) + Trim(txtVersion.Text) + ".doc"
       spath = "\\10.11.1.25\Public\Manufacture\逆向标签模板\逆向装箱清单\" + Trim(txtModel.Text) + Trim(txtVersion.Text) + ".doc"
           
        Set fs = CreateObject("Scripting.FileSystemObject")
            If Not fs.FileExists(spath) Then
                MsgBox "没有对应机种打印模板", vbOKOnly + vbExclamation, "警告"
                txtSN.Text = ""
                txtVersion.Text = ""
                txtModel.Text = ""
                txtSN.SetFocus
                Exit Sub
            End If
            
        Dim wrdObject As Word.Application
        Dim wrdDoc As Word.Document
        Set wrdObject = CreateObject("Word.Application")
        Set wrdDoc = wrdObject.Documents.Open(spath)
        
        wrdObject.Visible = True
        'wrdObject.Selection.TypeText "This is some text."     '写入文字到word中 210231a86ph095000083
        wrdDoc.PrintOut
        'wrdDoc.ClosePrintPreview  '关闭打印预览
        wrdDoc.Close          '关闭word文档
        wrdObject.Quit   'word应用退出
        Set wrdDoc = Nothing '释放内存
        Set wrdObject = Nothing '释放内存
        txtSN.Text = ""
        txtVersion.Text = ""
        txtModel.Text = ""
        txtSN.SetFocus
End Sub


Private Sub Form_Load()
    If conn.State = 0 Then
      conn.ConnectionString = Connect.getConnectionstring
      conn.Open
   End If

   If conn2.State = 0 Then
        conn2.CursorLocation = adUseClient
        conn2.ConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
        conn2.ConnectionTimeout = 100
        conn2.Open
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
   If conn2.State = 2 Then
    conn2.Close
    Set conn2 = Nothing
   End If
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If Len(txtSN.Text) <> 20 Then
         MsgBox "产品序号长度不等于20!"
         txtSN.SetFocus
         Exit Sub
        End If

      
      If Len(Replace(Trim(txtSN.Text), Chr(13) & Chr(10), "")) = 20 Then
        
         txtModel.Text = Mid(Trim(txtSN.Text), 3, 8)
            
      End If
      
    End If
End Sub
