VERSION 5.00
Begin VB.Form index 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  '없음
   Caption         =   "DConvert"
   ClientHeight    =   4185
   ClientLeft      =   5070
   ClientTop       =   -135
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5160
      Top             =   0
   End
   Begin VB.TextBox tup3 
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   5
      Text            =   "index.frx":0000
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton end 
      Caption         =   "닫기"
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton updayt 
      Caption         =   "업데이트 체크"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label tup 
      BackStyle       =   0  '투명
      Caption         =   "◎업데이트◎"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label thank 
      BackStyle       =   0  '투명
      Caption         =   "이용해 주셔서 감사합니다."
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label N1 
      BackStyle       =   0  '투명
      Caption         =   "D C o n v e r t"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub end_Click()
Me.Hide

End Sub

Private Sub Form_Load()
Open "Updayt.txt" For Output As #1
Write #1, "일부 오류를 해결 했습니다. 오류신고는 gnyontu39@gmail.com로 보내주시면 빠른 시일내에 업데이트를 진행하도록 하겠습니다."
Close

Open "boot.ini" For Output As #1
Write #1, "DConvert 초기화 중입니다."
Close

Open "setting.ini" For Output As #1
Write #1, ""
Write #1, ""
Write #1, "[기타 설정]"
Write #1, ""
Write #1, "skin : 1"
Write #1, ""
Write #1, "ver : 2.41"
Write #1, ""
Write #1, "sever : http://mote.site88.net"
Write #1, ""
Write #1, "updayt : 2011.04.09"
Write #1, ""
Write #1, "avi"
Write #1, ""
Write #1, "240x240"
Write #1, ""
Write #1, "app.path\convert"
Write #1, ""
Write #1, ""
Write #1, "[언어]"
Write #1, "korea"
Close

Dim up32 As String
Open "Updayt.txt" For Input As #2
Input #2, up32

tup3.Text = up32
Close

Unload TRY
Open "boot.ini" For Output As #1
Write #1, "초기화 완료"
Close
Load TRY

End Sub

Private Sub N1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub Timer1_Timer()
TRY.Show
main.Show
Timer1.Enabled = False

End Sub

Private Sub updayt_Click()
Unload TRY
Open "boot.ini" For Output As #1
Write #1, "본 기능은 인터넷을 요구합니다."
Close
Load TRY
Shell "updayt.exe", vbNormalFocus
End

End Sub
