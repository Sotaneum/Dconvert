VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form updayt2 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
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
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Cd1 
      Caption         =   "���� Ȯ�� ��....."
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
      ScrollBars      =   2  '����
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
      BorderStyle     =   1  '���� ����
      Caption         =   "���� ���� : "
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '���� ����
      Caption         =   "�ֽ� ���� : "
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
If Cd1.Caption = "������Ʈ" Then
Dim Msg
Msg = MsgBox("������Ʈ�� �߰ߵǾ����ϴ�. ������Ʈ�� �����Ͻðڽ��ϱ�?", vbQuestion & vbYesNo, "Updayt")

If Msg = vbYes Then

ShowInfo "������ ������ ��õ� �մϴ�."
ShowInfo "���� ������ ���� �մϴ�."
Kill App.Path & "\DConvert.exe"
ShowInfo "���� ������ ���� �Ϸ�."
ShowInfo "���ο� ������ �ӽ� ������ �ٿ�ε� �޽��ϴ�. �� �������� �ణ�� ���� �߻� �� �� ������, ������Ʈ�� ���� �� �� �ֽ��ϴ�."
MkDir App.Path & "\updayt"
ShowInfo "�ӽ� ������ �����մϴ�."
DownloadFileFromWeb "http://dmote.hoseoi.com/updayt/up34.txt", App.Path & "\updayt\sup.exe"
ShowInfo "�ֽ� ������ �ӽ������� ��ġ �Ϸ�"
FileCopy App.Path & "\updayt\sup.exe", App.Path & "\DConvert.exe"
ShowInfo "������Ʈ �Ϸ�."
Kill App.Path & "\updayt"
ShowInfo "�ӽ� ���� ���� �Ϸ�"
ShowInfo "�����մϴ�."
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
ShowInfo "������ ���� ���Դϴ�.."
Upday = "http://dmote.hostoi.com/Updayt/DCU.ini"
Label1.Caption = Inet1.OpenURL(Upday)
Label2.Caption = App.Major & "." & App.Minor

ShowInfo "������ Ȯ�� ���Դϴ�."
If Label1.Caption <= Label2.Caption Then
ShowInfo " �ֽŹ����� �����ϴ�."
Cd1.Enabled = True
Cd1.Caption = "���α׷� �����"

ElseIf Label1.Caption > Label2.Caption Then
ShowInfo "������Ʈ�� �߰ߵǾ����ϴ�. ������Ʈ�� ���� �غ��մϴ�."
Cd1.Enabled = True
Cd1.Caption = "������Ʈ"

End If

End Sub

Function ShowInfo(body As String)
    Text1 = Text1 & body & vbCrLf
    Text1.SelStart = Len(Text1)
End Function
