VERSION 5.00
Begin VB.Form frmVersion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "收集版本(Collect Version)"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVersion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(Cancel)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "确定(Confirm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2880
      Width           =   4095
   End
   Begin VB.TextBox txtVer 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox txtModel 
      BackColor       =   &H80000000&
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
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtSN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblPass 
      Caption         =   "确        认(Pass):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label lblVer 
      Caption         =   "版        本(Version):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblModel 
      Caption         =   "机种名称(Model):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblSN 
      Caption         =   "产品条码(Serial Number):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conn As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim myApp As New LabelManager2.Application
Dim myDoc As LabelManager2.Document
Dim myVars As LabelManager2.Variables
Dim myBarcode As LabelManager2.Barcodes
Private Sub cmdCancel_Click()
   txtSN.Text = ""
   txtModel.Text = ""
   txtVer.Text = ""
   txtPass.Text = ""
   txtSN.SetFocus
End Sub

Private Sub cmdConfirm_Click()
   If txtSN.Text = "" Then
       MsgBox "产品条码不能为空!"
       txtSN.SetFocus
       Exit Sub
   End If
   If txtModel.Text = "" Then
      MsgBox "机种名称不能为空!"
      txtModel.SetFocus
      Exit Sub
   End If
   If txtVer.Text = "" Then
      MsgBox "版本不能为空!"
      txtVer.SetFocus
      Exit Sub
   End If
   'sql = "insert into version(sn,model,ver,operator,testtime) values('" & txtSN.Text & "','" & txtModel.Text & "','" & txtVer.Text & "','" & golUSERNAME & "','" & Format(Now, "yyyy-MM-DD HH:MM:SS") & "')"
   sql = "insert into version(sn,model,ver,operator,testtime) values('" & Trim(txtSN.Text) & "','" & Trim(txtModel.Text) & "','" & Trim(txtVer.Text) & "','" & golUSERNAME & "','" & Format(Now, "yyyy-MM-DD HH:MM:SS") & "')"
   conn.Execute sql
   txtSN.Text = ""
   txtModel.Text = ""
   txtVer.Text = ""
   txtPass.Text = ""
   txtSN.SetFocus
End Sub

Private Sub Form_Load()
  golPath = Connect.getConnectionstring
  conn.ConnectionString = golPath
  conn.Open
End Sub

Private Sub Form_Unload(Cancel As Integer)
   conn.Close
   Set conn = Nothing
End Sub

Private Sub txtPass_Change()
   Dim loc As Integer
   loc = InStr(txtPass.Text, Chr(13) & Chr(10))
   If loc > 0 Then
      txtPass.Text = Left(txtPass.Text, loc - 1)
      If UCase(txtPass.Text) = "PASS" Then
        If ver = "" Or txtVer.Text = ver Then
           info = "请确认此机种版本是否为"
           nver = txtVer.Text
           frmInformation.Show 1
           If result = "OK" Then
               Call cmdConfirm_Click
            Else
               txtPass.Text = ""
               txtVer.SetFocus
            End If
        ElseIf txtVer.Text <> ver Then
           info = "请确认此机种版本是否要更新为"
           nver = txtVer.Text
           frmInformation.Show 1
            If result = "OK" Then
               Call cmdConfirm_Click
            Else
               txtPass.Text = ""
               txtVer.SetFocus
            End If
        End If
      Else
        txtPass.Text = ""
        txtPass.SetFocus
      End If
   End If
End Sub

Private Sub txtSN_Change()
   Dim loc As Integer
   Dim rs As New ADODB.Recordset
   loc = InStr(txtSN.Text, Chr(13) & Chr(10))
   If loc > 0 Then
      txtSN.Text = Left(txtSN.Text, loc - 1)
      If Len(txtSN.Text) < 10 Then
         MsgBox "产品条码长度不正确的,请重新输入！"
         txtSN.Text = ""
         txtSN.SetFocus
         Exit Sub
      End If
      txtModel.Text = Mid(txtSN.Text, 3, 8)
      sql = "select ver from version where model='" & txtModel.Text & "'"
      rec.Open sql, conn, adOpenKeyset, adLockOptimistic
      If rec.EOF = True Then
         txtVer.Text = ""
         txtVer.SetFocus
      Else
        Dim rcd As New ADODB.Recordset
        sql = "select max(testtime) from version where model='" & txtModel.Text & "'"
        rcd.Open sql, conn, adOpenKeyset, adLockOptimistic
        If rcd.EOF = True Then
           txtVer.Text = ""
           txtVer.SetFocus
        Else
           sql = "select ver from version where testtime='" & rcd.Fields(0) & "' and model='" & txtModel.Text & "'"
           rs.Open sql, conn, adOpenKeyset, adLockOptimistic
           If rs.EOF = False Then
              ver = rs.Fields(0)
              txtVer.Text = ver
           End If
           rs.Close
           txtPass.SetFocus
        End If
        rcd.Close
      End If
      rec.Close
   End If
End Sub

Private Sub txtVer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'If InStr(txtVer.Text, "PASS") > 0 Or InStr(txtVer.Text, "FAIL") > 0 Or InStr(txtVer.Text, "pass") > 0 Or InStr(txtVer.Text, "fail") > 0 Or InStr(txtVer.Text, "=") > 0 Or InStr(txtVer.Text, "[") > 0 Or InStr(txtVer.Text, "]") > 0 Or InStr(txtVer.Text, "\") > 0 Or InStr(txtVer.Text, ";") > 0 Or InStr(txtVer.Text, "'") > 0 Or InStr(txtVer.Text, ",") > 0 Or InStr(txtVer.Text, ".") > 0 Or InStr(txtVer.Text, "*") > 0 Or InStr(txtVer.Text, "+") > 0 Then
      If InStr(txtVer.Text, "PASS") > 0 Or InStr(txtVer.Text, "FAIL") > 0 Or InStr(txtVer.Text, "pass") > 0 Or InStr(txtVer.Text, "fail") > 0 Or InStr(txtVer.Text, "=") > 0 Or InStr(txtVer.Text, "[") > 0 Or InStr(txtVer.Text, "]") > 0 Or InStr(txtVer.Text, "\") > 0 Or InStr(txtVer.Text, ";") > 0 Or InStr(txtVer.Text, "'") > 0 Or InStr(txtVer.Text, ",") > 0 Or InStr(txtVer.Text, ".") > 0 Or InStr(txtVer.Text, "*") > 0 Or InStr(txtVer.Text, "+") > 0 Or InStr(txtVer.Text, " ") > 0 Then
         MsgBox "版本里含有非法字符,请重新输入！"
         txtVer.Text = ""
         txtVer.SetFocus
         Exit Sub
      Else
         txtPass.SetFocus
      End If
   End If
End Sub
