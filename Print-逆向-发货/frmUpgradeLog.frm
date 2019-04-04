VERSION 5.00
Begin VB.Form frmUpgradeLog 
   Caption         =   "打印程序升级日志"
   ClientHeight    =   9525
   ClientLeft      =   3015
   ClientTop       =   1020
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   14895
   Begin VB.ListBox lstMain 
      Height          =   8880
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   14295
   End
End
Attribute VB_Name = "frmUpgradeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.lstMain.AddItem ("HUV pring program(生产,备件,备件良品) upgraded by allen.yan on 6/9/2014 for setting upgrading(adding 4 columns) requirement from Shun.Huang")
    Me.lstMain.AddItem ("DaHua pring program(生产,备件,备件良品) upgraded by Robin.Huang on 9/3/2019 for setting upgrading(adding 5 columns) requirement from GuoDong.Ma")
End Sub

