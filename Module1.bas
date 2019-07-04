Attribute VB_Name = "Module1"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Sub DoDrag(TheForm As Form)
' TheForm:  The form you want to start dragging
    
    ReleaseCapture
    SendMessage TheForm.hWnd, &HA1, 2, 0&
End Sub
