VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "��ӡ���İ�"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   4770
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdfrmMain 
      Caption         =   "BU1��ӡ"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdSetting 
      Caption         =   "�����趨"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdfrmMain_Click()
    frmBU1Print.Show
End Sub

Private Sub cmdSetting_Click()
    FrmMaintain.Show
End Sub


