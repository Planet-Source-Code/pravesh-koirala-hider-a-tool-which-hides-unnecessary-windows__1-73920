VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hider"
   ClientHeight    =   1980
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton dummy 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select a Window To Hide"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "HideWindow"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Or just press CTRL + ALT + SPACE in the window which you want to hide"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Hwnd"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "Show"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAX = 10
Dim Windows(MAX) As Window
Dim NextFree As Long    'Counts next free class


Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Sub Command1_Click()
    Dim szHwnd As String, hwnd As Long
    szHwnd = Text1.text
    
    If szHwnd = "" Then Exit Sub
    If NextFree = -1 Then
        MsgBox "Sorry, quota full!!", vbOKOnly Or vbExclamation
        Exit Sub
    End If
   
    
    hwnd = Val("&H" & szHwnd)
    If IsWindow(hwnd) = 0 Then
        MsgBox "Not Valid window", vbOKOnly Or vbCritical
        Exit Sub
    End If
    Windows(NextFree).hwnd = hwnd
    Windows(NextFree).Show (False)
    
    FindNextFree
End Sub

Public Sub Form1_Mouse(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Windows(Shift).Show True
    End If
End Sub

Private Sub Form_Load()
    Dim X As Long
    For X = 0 To MAX
        Set Windows(X) = New Window
        Windows(X).Id = X
    Next
    ModMain.SubclassWindow Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Long
    For X = 0 To MAX
        If Windows(X).Hiding = True Then Windows(X).Show True   'Show all hidden windows
        Set Windows(X) = Nothing
       
    Next
    ModMain.StopSubclass Me.hwnd
End Sub


Private Sub FindNextFree()
    Dim Y As Long
    For Y = 0 To MAX
        If Windows(Y).Hiding = False Then
           NextFree = Y
           Exit Sub
        End If
    Next
    NextFree = -1
End Sub

Private Sub Text1_Change()
    If Text1.text = "" Then Command1.Enabled = False Else Command1.Enabled = True
End Sub
