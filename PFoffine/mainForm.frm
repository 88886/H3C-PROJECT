VERSION 5.00
Begin VB.Form mainForm 
   Caption         =   "PacketFront Label Print"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   7710
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdMAC 
      Caption         =   "MAC"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdMBSN 
      Caption         =   "MB SN"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "��ӡ��¼����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   4800
      Width           =   3855
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "PF TRACK��ǩ��ӡ(��ӡ����)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CommandButton cmdUnit 
      Caption         =   "PF Unit��ǩ��ӡ(��ӡ����)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "PacketFront ��ӡ���İ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExport_Click()
    frmReport.Show
End Sub

Private Sub cmdMAC_Click()
    frmMAC.Show
End Sub

Private Sub cmdMBSN_Click()
    frmMBSN.Show
End Sub

Private Sub cmdTrack_Click()
    formTrack.Show
End Sub

Private Sub cmdUnit_Click()
    formUnit.Show
End Sub

