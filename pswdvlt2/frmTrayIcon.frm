VERSION 5.00
Begin VB.Form frmTrayIcon 
   Caption         =   "Tray Icon..."
   ClientHeight    =   1335
   ClientLeft      =   1515
   ClientTop       =   4560
   ClientWidth     =   2700
   Icon            =   "frmTrayIcon.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1335
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox pichook 
      Height          =   555
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnuBar 
      Caption         =   "PopupMenu"
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Quit"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA

Private Sub Form_Load()
    Me.Icon = LoadResPicture(101, vbResIcon)
    t.cbSize = Len(t)
    t.hWnd = pichook.hWnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
    t.szTip = "Password Vault" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, t
    Me.Hide
    App.TaskVisible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    t.cbSize = Len(t)
    t.hWnd = pichook.hWnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub


Private Sub mnuAbout_Click()
    frmSplash.Show
End Sub

Private Sub mnuMain_Click(Index As Integer)
    frmMain.cmdExit_Click
End Sub

Private Sub mnuMinimize_Click()
    frmMain.WindowState = 1
End Sub

Private Sub mnuRestore_Click()
    frmMain.WindowState = 0
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static rec As Boolean, msg As Long
    msg = x / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case msg
            Case WM_LBUTTONDBLCLK:
                If frmMain.WindowState = 0 Then
                    mnuMinimize_Click
                Else
                    mnuRestore_Click
                End If
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                Me.PopupMenu mnuBar
        End Select
        rec = False
    End If
End Sub
