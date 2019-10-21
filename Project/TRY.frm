VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form TRY 
   BorderStyle     =   0  '없음
   Caption         =   "D C onvert"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   705
   ClientWidth     =   4155
   Icon            =   "TRY.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1680
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   1080
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   1085
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.PictureBox P 
      Height          =   330
      Left            =   1920
      ScaleHeight     =   270
      ScaleWidth      =   240
      TabIndex        =   0
      ToolTipText     =   "fghfgfgb"
      Top             =   360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Menu msys 
      Caption         =   "msys"
      Visible         =   0   'False
      Begin VB.Menu mclose 
         Caption         =   "닫기 (&Close)"
      End
      Begin VB.Menu mopen 
         Caption         =   "열기 (&Open)"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mabout 
         Caption         =   "정보 (&About)"
      End
   End
End
Attribute VB_Name = "TRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutOrVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type
  

Private Const NIIF_WARNING = 2
Private Const NIIF_ERROR = 3
Private Const NIIF_INFO = 1

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Dim SysTrayT As NOTIFYICONDATA

Private Sub Form_Load()
Dim boomt As String
Open "boot.ini" For Input As #1
Input #1, boomt
TRY.Caption = boomt
Close

Dim kgo As String
Me.Hide
index.Hide
    With SysTrayT
        .cbSize = Len(SysTrayT)
        .hWnd = P.hWnd
        .uID = 1
        .uFlags = &H2 Or &H1 Or &H10 Or &H4
        .hIcon = Me.Icon
        .uCallbackMessage = &H200
        
        .szInfo = boomt & Chr(0)  ' 풍선 메세지
        .uTimeoutOrVersion = 10000 '풍선을 보여줌 (1000 = 1초)
        .dwInfoFlags = 1 ' 풍선 아이콘 : 1 = 정보, 2 = 주의, 3 = 에러
    End With
    
        Shell_NotifyIcon &H0, SysTrayT
Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell_NotifyIcon &H2, SysTrayT

End Sub

Private Sub Form_Resize()
Dim kgo As String
kgo = TRY.Caption
    With SysTrayT
        .cbSize = Len(SysTrayT)
        .hWnd = P.hWnd
        .uID = 1
        .uFlags = &H2 Or &H1 Or &H10 Or &H4
        .hIcon = Me.Icon
        .uCallbackMessage = &H200
        
        .szInfo = kgo & Chr(0)  ' 풍선 메세지
        .uTimeoutOrVersion = 30000 '풍선을 보여줌 (1000 = 1초)
        .dwInfoFlags = 1 ' 풍선 아이콘 : 1 = 정보, 2 = 주의, 3 = 에러
            End With
    
        Shell_NotifyIcon &H0, SysTrayT
End Sub

Private Sub mabout_Click()
Unload TRY
Open "boot.ini" For Output As #1
Write #1, "정보를 읽어 냅니다."
Close
Load TRY

index.Show
End Sub

Private Sub mclose_Click()

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

Private Sub mopen_Click()
Unload TRY
Open "boot.ini" For Output As #1
Write #1, "프로그램을 전환합니다."
Close
Load TRY

main.Show

End Sub

Private Sub P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Static rec As Boolean, msg As Long
    msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case msg
            Case &H202: main.Show '왼쪽마우스 클릭
            Unload TRY
Open "boot.ini" For Output As #1
Write #1, "프로그램을 전환합니다."
Close
Load TRY
            Case &H205: PopupMenu msys ' 오른쪽 마우스 클릭
        End Select
        rec = False
    End If
End Sub

Private Sub Timer1_Timer()
Me.Hide

End Sub

