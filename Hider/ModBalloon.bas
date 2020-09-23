Attribute VB_Name = "ModBalloon"
'In this module we add tray icon and also used to display a balloon

Option Explicit

Private Const APP_SYSTRAY_ID = 420 'unique identifier

Private Const WM_USER = &H400

Public Const NIN_BALLOONSHOW = (WM_USER + 2)
Public Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
Private Const WM_APP As Long = &H8000&
Private Const WM_MOUSEMOVE = &H200
Private Const NOTIFYICON_VERSION = &H3
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_INFO = &H10
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETVERSION = &H4
Private Const NIS_SHAREDICON = &H2
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)


'NOTIFIYICONDATA  size
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  '-5.0 structure
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 '-6.0 structure
Private NOTIFYICONDATA_SIZE As Long

Private Type GUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uID As Long
uFlags As Long
uCallbackMessage As Long
hIcon As Long
szTip As String * 128
dwState As Long
dwStateMask As Long
szInfo As String * 256
uTimeoutAndVersion As Long
szInfoTitle As String * 64
dwInfoFlags As Long
guidItem As GUID
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" _
   Alias "Shell_NotifyIconA" _
  (ByVal dwMessage As Long, _
   lpData As NOTIFYICONDATA) As Long

Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lpBuffer As Any, _
   nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Public Type WinData
    Title As String * 50
    Id As Long
    Icon As Long
End Type

Public Sub ShellTrayAdd(Data As WinData)

Dim nid As NOTIFYICONDATA
If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
'set up the type members
With nid
    .cbSize = NOTIFYICONDATA_SIZE
    .hwnd = Form1.hwnd
    .uID = Data.Id
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .dwState = NIS_SHAREDICON
    .hIcon = Data.Icon
    .uCallbackMessage = WM_MOUSEMOVE
    .szTip = Data.Title & vbNullChar
    .uTimeoutAndVersion = NOTIFYICON_VERSION
End With
    'add the icon ...
    Call Shell_NotifyIcon(NIM_ADD, nid)
    'Tell the system about the version of NOTIFYICON in use
Call Shell_NotifyIcon(NIM_SETVERSION, nid)
End Sub


Public Sub ShellTrayRemove(ByVal Id As Long)
    Dim nid As NOTIFYICONDATA
    If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
    
    With nid
        .cbSize = NOTIFYICONDATA_SIZE
        .hwnd = Form1.hwnd
        .uID = Id
    End With
    
    Call Shell_NotifyIcon(NIM_DELETE, nid)
End Sub


Public Sub ShellTrayModifyTip(ByVal text As String)
    Dim nid As NOTIFYICONDATA
    If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
    
    With nid
        .cbSize = NOTIFYICONDATA_SIZE
        .hwnd = Form1.hwnd
        .uID = APP_SYSTRAY_ID
        .uFlags = NIF_INFO
        .dwInfoFlags = 1
        .hIcon = Form1.Icon
        .szInfoTitle = "Immunize!!!" & vbNullChar
        .szInfo = text & vbNullChar
    End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, nid)
End Sub


Private Sub SetShellVersion()
    'Here we find the major version of shell32.dll and use the NotifyIconData
    'accordingly
    
    Dim BufferSize As Long
    Dim Unused As Long
    Dim Bufferx As Long
    Dim VerMajor As Integer
    Dim Buffer() As Byte
    
    Const DLLFile As String = "shell32.dll"
    
    BufferSize = GetFileVersionInfoSize(DLLFile, Unused)
    If BufferSize > 0 Then
        ReDim Buffer(BufferSize - 1) As Byte
        Call GetFileVersionInfo(DLLFile, 0&, BufferSize, Buffer(0))
            If VerQueryValue(Buffer(0), "\", Bufferx, Unused) = 1 Then
                CopyMemory VerMajor, ByVal Bufferx + 10, 2
            End If  'VerQueryValue
    End If  'BufferSize
    

    If VerMajor = 6 Then
        NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE    '6 or +6 Structure
    ElseIf VerMajor = 5 Then
        NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE    '-6 Structure
    Else
        NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE    '-5 Structure
    End If
    
End Sub


