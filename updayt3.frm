VERSION 5.00
Begin VB.Form updayt3 
   Caption         =   "updayt lnst"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdMethod2 
      Caption         =   "����"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "updayt3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Me.Caption = "Mote.exe ��ġ��...."
Label1.Caption = "���� �ٿ�ε� : Mote.exe"
Kill App.Path & "\Mote.exe"
DownloadFileFromWeb "http://motepg.net78.net/note_for_get/up/Mote.dat", App.Path & "\Mote.exe"
MsgBox "��ġ�� �Ϸ� �Ǿ����ϴ�. ���α׷��� �����մϴ�.", vbOKOnly + vbInformation, "Updayt"
Shell "mote.exe", vbNormalFocus
End
End Sub
