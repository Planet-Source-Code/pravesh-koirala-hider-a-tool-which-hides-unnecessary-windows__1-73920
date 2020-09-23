VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const WM_GETICON = &H7F
Private Const ICON_SMALL = 0
Private Const ICON_SMALL2 = 2
Private Const ICON_BIG = 1
Private Const SW_SHOW = 5
Private Const SW_HIDE = 0
Private Const MAX = 10

Public Id As Long
Public hwnd As Long
Public Hiding As Boolean

Public Sub Show(ByVal bShow As Boolean)
    If hwnd = 0 Then Exit Sub

    If bShow Then
        ShowWindow hwnd, SW_SHOW
        
        Me.Hiding = False
        ModBalloon.ShellTrayRemove Id
        'Form1.List1.RemoveItem Id
    Else
        ShowWindow hwnd, SW_HIDE
        
        'TODO: ADD Notification to down
        DoHidingStuffs
        
        Me.Hiding = True
    End If

End Sub

Private Sub DoHidingStuffs()
    Dim X As WinData, i As Long
    With X
        GetWindowText hwnd, .Title, 50
        .Id = Id
        .Icon = GetIconFromHwnd()
    End With
    ModBalloon.ShellTrayAdd X
    
End Sub

Private Function GetIconFromHwnd() As Long
    Dim Icon As Long, Dllerror As Long
    'Small Icon
    Icon = SendMessage(hwnd, WM_GETICON, ICON_BIG, 0)
    If Icon <> 0 Then GoTo AssignAndReturn
    'Big Icon
    Icon = SendMessage(hwnd, WM_GETICON, ICON_SMALL, 0)
    If Icon <> 0 Then GoTo AssignAndReturn
    'Alternatively
    Icon = GetClassLong(hwnd, -34)
    If Icon = 0 Then Icon = Form1.Icon
AssignAndReturn:
    GetIconFromHwnd = Icon
'    Debug.Print
End Function
