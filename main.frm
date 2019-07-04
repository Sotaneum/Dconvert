VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form main 
   BackColor       =   &H80000009&
   BorderStyle     =   0  '없음
   Caption         =   "D C o N v e r t 2.41v"
   ClientHeight    =   5505
   ClientLeft      =   4005
   ClientTop       =   4020
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   367
   ScaleMode       =   0  '사용자
   ScaleWidth      =   538
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   6000
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5280
      Top             =   840
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1320
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "음악 선택"
   End
   Begin VB.ListBox List1 
      Height          =   1140
      ItemData        =   "main.frx":0000
      Left            =   240
      List            =   "main.frx":0002
      TabIndex        =   8
      Top             =   1200
      Width           =   7455
   End
   Begin VB.PictureBox WebBrowser1 
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   6195
      TabIndex        =   6
      Top             =   3960
      Width           =   6255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "설정"
      Height          =   495
      Left            =   6600
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   7455
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   13150
      _cy             =   873
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Caption         =   "http://dmote.hostoi.com/hi.php"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "제거"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label plus1 
      BackStyle       =   0  '투명
      Caption         =   "추가"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   6720
      Picture         =   "main.frx":0004
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   536
      Y1              =   368
      Y2              =   368
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "Copyrigt ⓒ 2011 Gnyontu39@gmail.com"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Line Line2 
      X1              =   16
      X2              =   512
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "파일을 선택해 주십시오."
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   7455
   End
   Begin VB.Image Image3 
      Height          =   4815
      Left            =   0
      Picture         =   "main.frx":1DF8
      Stretch         =   -1  'True
      Top             =   720
      Width           =   8040
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   528
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   7200
      Picture         =   "main.frx":8370C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "D C o n v e r t 2.41v"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   0
      Picture         =   "main.frx":85500
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   8085
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Sub Command4_Click()
Unload TRY
Open "boot.ini" For Output As #1
Write #1, "설정모드로 전환합니다."
Close
Load TRY

End Sub

Private Sub Form_Load()
index.Show

    SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 530, 350, 30, 30), True
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub Image2_Click()
Unload TRY
Open "boot.ini" For Output As #1
Write #1, "프로그램을 종료 합니다."
Close
Load TRY

Kill App.Path & "\boot.ini"
Kill App.Path & "\setting.ini"
Kill App.Path & "\Updayt.txt"

End

End Sub

Private Sub Image4_Click()
Unload TRY
Open "boot.ini" For Output As #1
Write #1, "트래이 모드로 전환합니다. 다시 이 아이콘을 클릭하시면 창을 띄울 수 있습니다."
Close
Load TRY

Me.Hide

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub list1_Change()
WindowsMediaPlayer1.URL = List1.Text

End Sub

Private Sub Label4_Click()
Unload TRY
Open "boot.ini" For Output As #1
Write #1, "음악을 제거합니다."
Close
Load TRY

List1.Text = ""
List1.Clear

End Sub

Private Sub Label6_Click()
Unload TRY
Open "boot.ini" For Output As #1
Write #1, "경고 : 이 기능은 인터넷이 연결이 가능해야만 이용이 가능합니다."
Close
Load TRY

main.Show

Shell "explorer http://www.dmote.hostoi.com/hi.php/"

End Sub

Private Sub List1_Click()
WindowsMediaPlayer1.URL = List1.Text
Unload TRY
Open "boot.ini" For Output As #1
Write #1, "재생합니다." & WindowsMediaPlayer1.playState
Close
Load TRY

End Sub

Private Sub plus1_Click()
Unload TRY
Open "boot.ini" For Output As #1
Write #1, "음악을 추가합니다."
Close
Load TRY

    CD1.ShowOpen
    If CD1.FileName = "" Then
    Else
        List1.AddItem (CD1.FileName)
    End If
    
End Sub

Private Sub Timer1_Timer()
Label2.Caption = WindowsMediaPlayer1.Status & "  즐거운 감상 되세요." & "hWnd : " & List1.hWnd

End Sub

Private Sub Timer2_Timer()
If Not (WindowsMediaPlayer1.Status = "") Then
Unload TRY
Open "boot.ini" For Output As #1
Write #1, WindowsMediaPlayer1.Status
Close
Load TRY
End If

End Sub
