VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form updayt2 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "DCont Updayt"
   ClientHeight    =   2025
   ClientLeft      =   5925
   ClientTop       =   4725
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Cd1 
      Caption         =   "상태 확인 중....."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5160
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '단일 고정
      Caption         =   "현제 버전 : "
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Caption         =   "최신 버전 : "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "updayt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cd1_Click()
If Cd1.Caption = "업데이트" Then
Dim Msg
Msg = MsgBox("업데이트가 발견되었습니다. 업데이트를 진행하시겠습니까?", vbQuestion & vbYesNo, "Updayt")

If Msg = vbYes Then

ShowInfo "서버와 연결을 재시도 합니다."
ShowInfo "기존 버전을 제거 합니다."
Kill App.Path & "\DConvert.exe"
ShowInfo "기존 버전을 제거 완료."
ShowInfo "새로운 버전을 임시 폴더로 다운로드 받습니다. 이 과정에서 약간의 렉이 발생 할 수 있으며, 업데이트에 실패 할 수 있습니다."
MkDir App.Path & "\updayt"
ShowInfo "임시 폴더를 생성합니다."
DownloadFileFromWeb "http://dmote.hoseoi.com/updayt/up34.txt", App.Path & "\updayt\sup.exe"
ShowInfo "최신 버전을 임시폴더에 설치 완료"
FileCopy App.Path & "\updayt\sup.exe", App.Path & "\DConvert.exe"
ShowInfo "업데이트 완료."
Kill App.Path & "\updayt"
ShowInfo "임시 폴더 삭제 완료"
ShowInfo "실행합니다."
Shell "DConvert.exe", vbNormalFocus
End If

If Msg = vbNo Then
End If
Else
Shell "explorer DConvert.exe"

End If


End Sub

Private Sub Form_Load()
Dim Upday As String
ShowInfo "서버와 연결 중입니다.."
Upday = "http://dmote.hostoi.com/Updayt/DCU.ini"
Label1.Caption = Inet1.OpenURL(Upday)
Label2.Caption = App.Major & "." & App.Minor

ShowInfo "버전을 확인 중입니다."
If Label1.Caption <= Label2.Caption Then
ShowInfo " 최신버전이 없습니다."
Cd1.Enabled = True
Cd1.Caption = "프로그램 재실행"

ElseIf Label1.Caption > Label2.Caption Then
ShowInfo "업데이트가 발견되었습니다. 업데이트를 위해 준비합니다."
Cd1.Enabled = True
Cd1.Caption = "업데이트"

End If

End Sub

Function ShowInfo(body As String)
    Text1 = Text1 & body & vbCrLf
    Text1.SelStart = Len(Text1)
End Function
