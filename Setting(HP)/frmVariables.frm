VERSION 5.00
Begin VB.Form frmVariables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Variables Setting"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVariables.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11520
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command7 
      Caption         =   "SRIE�����趨"
      Height          =   735
      Left            =   7920
      TabIndex        =   15
      Top             =   2040
      Width           =   2145
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "NEC�����趨"
      Height          =   735
      Left            =   7920
      TabIndex        =   14
      Top             =   1080
      Width           =   2145
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00008000&
      Caption         =   "�󻪱����趨"
      Height          =   735
      Left            =   5040
      TabIndex        =   13
      Top             =   1080
      Width           =   2145
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���������趨"
      Height          =   735
      Left            =   5040
      TabIndex        =   12
      Top             =   3000
      Width           =   2145
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UNIS �����趨"
      Height          =   735
      Left            =   1560
      TabIndex        =   11
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton bthHP 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HP�����趨"
      Height          =   735
      Left            =   5040
      TabIndex        =   10
      Top             =   2040
      Width           =   2145
   End
   Begin VB.CommandButton cmdH3C 
      Caption         =   "H3C �����趨"
      Height          =   735
      Left            =   1560
      TabIndex        =   9
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmd3COM 
      Caption         =   "3COM��NEC��Դ�����趨"
      Height          =   735
      Left            =   1560
      TabIndex        =   8
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdHUAWEI 
      BackColor       =   &H00404040&
      Caption         =   "HUAWEI �����趨"
      Height          =   735
      Left            =   1560
      TabIndex        =   7
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton cmdH3C_3COM 
      Caption         =   "H3C-3COM �����趨"
      Height          =   735
      Left            =   1560
      TabIndex        =   6
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "HUV �����趨"
      Height          =   735
      Left            =   5040
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton cmdH3C2D 
      Caption         =   "H3C 2D ��ӡ�����趨"
      Height          =   735
      Left            =   5040
      TabIndex        =   4
      Top             =   3960
      Width           =   2145
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "��H3C�����趨"
      Height          =   735
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   2145
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "��    ��(Return)"
      Height          =   735
      Left            =   7560
      TabIndex        =   1
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "2013.09.03"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "��ǩ�����趨(Variables Setting)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bthHP_Click()
  frmHPSetting.Show 1
End Sub

Private Sub cmd3COM_Click()
    frmOthers.Show 1
End Sub

Private Sub cmdH3C_3COM_Click()
   frmH3C_3COMSetting.Show 1
End Sub

Private Sub cmdH3C_Click()
   frmH3CSetting.Show 1
End Sub

Private Sub cmdH3C2D_Click()
    frmH3C2DSetting.Show 1
End Sub

Private Sub cmdHUAWEI_Click()
   frmHUAWEISetting.Show 1
End Sub

Private Sub cmdReturn_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   frmHUVSetting.Show 1
End Sub

Private Sub Command2_Click()
    frmNewH3CSetting.Show 1
End Sub

Private Sub img3COM_Click()
   cmd3COM_Click
End Sub

Private Sub imgH3C_3COM_Click()
   cmdH3C_3COM_Click
End Sub

Private Sub imgH3C_Click()
   cmdH3C_Click
End Sub

Private Sub imgHUAWEI_Click()
   cmdHUAWEI_Click
End Sub

Private Sub Command3_Click()
    frmNewUNISSetting.Show 1
End Sub

Private Sub Command4_Click()
    frmNewConsenSetting.Show 1
End Sub

Private Sub Command5_Click()
    frmNewDaHuaSetting.Show 1
End Sub

Private Sub Command6_Click()
    frmNewNECSetting.Show 1
End Sub

Private Sub Command7_Click()
    frmSRIESetting.Show 1
End Sub

Private Sub Command8_Click()
   frmMobileCommunicationSetting.Show 1
End Sub
