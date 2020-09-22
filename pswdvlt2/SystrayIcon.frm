VERSION 5.00
Begin VB.Form SystrayIcon 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   7110
   ControlBox      =   0   'False
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SystrayIcon.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Label lblToolTip 
      Caption         =   "Password Vault!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label lblPresent 
      Caption         =   $"SystrayIcon.frx":0442
      Height          =   1695
      Left            =   1485
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "SystrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
  
Dim SysIcon As NOTIFYICONDATA

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        ' Put here your code to emulate systray icons events.
        Select Case x
            Case 7680   ' MouseMove
                
            Case 7695   ' Left MouseDown
                MsgBox x
            Case 7710   ' Left MouseUp
                MsgBox x
            Case 7725   ' Left DoubleClick
                MsgBox x
            Case 7740   ' Right MouseDown
                MsgBox x
            Case 7755   ' Right MouseUp
                MsgBox x
            Case 7770   ' Right DoubleClick
                MsgBox x
        End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Remove the icon when this form unload. Don't forget to unload this form!
    RemoveIcon
End Sub

Public Sub ShowIcon()
    ' Show the systray icon. Use from another form : "SystrayIcon.ShowIcon"
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = Me.hwnd
    SysIcon.uID = vbNull
    SysIcon.uFlags = 7
    SysIcon.uCallbackMessage = 512
    SysIcon.hIcon = Me.Icon
    SysIcon.szTip = lblToolTip.Caption + Chr(0)
    Shell_NotifyIcon 0, SysIcon
End Sub

Public Sub RemoveIcon()
    ' Remove the systray icon. Use from another form : "SystrayIcon.RemoveIcon"
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = Me.hwnd
    SysIcon.uID = vbNull
    SysIcon.uFlags = 7
    SysIcon.uCallbackMessage = vbNull
    SysIcon.hIcon = Me.Icon
    SysIcon.szTip = Chr(0)
    Shell_NotifyIcon 2, SysIcon
End Sub

