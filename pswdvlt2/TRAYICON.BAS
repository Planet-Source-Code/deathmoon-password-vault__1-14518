Attribute VB_Name = "basTrayIcon"
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
       
'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click


Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nid As NOTIFYICONDATA
       
Sub Initialise(mee As Form)  'Place in form load
    With nid
        .cbSize = Len(nid)
        .hwnd = mee.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = mee.Icon
        '.szTip = " Click Right Mouse Button " & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    mee.Hide
    App.TaskVisible = False

End Sub

Sub PopMenu(mee As Form, x As Single)  'Place in form mouse move
    Dim Msg As Long
    Msg = x / Screen.TwipsPerPixelX

    Select Case Msg
        Case WM_LBUTTONDBLCLK:
           
        Case WM_LBUTTONDOWN:
                
        Case WM_LBUTTONUP:
            mee.PopupMenu mee.mnuPopMenu
        Case WM_RBUTTONDBLCLK:
            
        Case WM_RBUTTONDOWN:
            
        Case WM_RBUTTONUP:
            mee.PopupMenu mee.mnuPopMenu
        
        End Select
End Sub

Sub CloseApp() 'Place in form unload
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Sub Down(mee As Form)  'Place in form resize
    If mee.WindowState = vbMinimized Then mee.Hide
End Sub

