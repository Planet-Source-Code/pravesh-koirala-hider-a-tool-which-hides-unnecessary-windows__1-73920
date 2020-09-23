Attribute VB_Name = "ModMain"
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal Id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal Id As Long) As Long
Private Const WM_HOTKEY = &H312
Private Const WM_MOUSEMOVE = &H200
Public Const WM_APP = &H8000
Public Const WM_CUSTOM = WM_APP + 25  'Arbitary number
Private Const WM_LBUTTONDBLCLICK = 515
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Const GWL_WNDPROC = (-4)
Private lastWndProc As Long
Private HotKey As Long

Public Sub SubclassWindow(ByVal hwnd As Long)
    If lastWndProc <> 0 Then Exit Sub   'Already subclassed
    HotKey = RegisterHotKey(hwnd, 1, MOD_ALT Or MOD_CONTROL, vbKeySpace)
    
    lastWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SubWindowProc)
End Sub

Public Sub StopSubclass(ByVal hwnd As Long)
    ' Remove the Forms subclassing
    If lastWndProc = 0 Then Exit Sub    'Not Subclassed
    SetWindowLong lWindowHandle, GWL_WNDPROC, lWindowProc
    UnregisterHotKey hwnd, 1
End Sub

Private Function SubWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim X As Long
    If Msg = WM_HOTKEY Then
        X = GetForegroundWindow()
        Debug.Print Err.LastDllError
        Form1.Text1.text = Hex(X)
        Form1.Command1_Click
   End If
   If Msg = WM_MOUSEMOVE Then
         If lParam = WM_LBUTTONDBLCLICK Then Form1.Form1_Mouse 1, CInt(wParam), 0, 0
    
End If
    SubWindowProc = CallWindowProc(lastWndProc, hwnd, Msg, wParam, lParam)
End Function

