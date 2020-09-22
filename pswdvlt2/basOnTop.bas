Attribute VB_Name = "basOnTop"
Option Explicit

'Declare API Functions
Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hwndinsertafter As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wflags As Long) As Long
    
'Declare our constants
'SWP stands for SetWindowPos
'SWP_NoSize tells SetWindowPos to ignore the cx and cy arguments
Private Const SWP_NoSize = &H1
'SWP_NoMove tells SetWindowPos to ignore the x and y arguments
Private Const SWP_NoMove = &H2
'hwnd_Top_Most is passed to SetWindowPos to set the target window Always on Top.
Private Const HWnd_TopMost = -1
'Hwn_NoTopMost is passed to SetWindowPos to remove Always On Top
Private Const hwnd_NoTopMost = -2
'Declare variables
Dim x As Byte
Public Sub SetFormOnTop(myForm As Object)
    SetWindowPos myForm.hwnd, HWnd_TopMost, 0, 0, 0, 0, SWP_NoMove Or SWP_NoSize
End Sub
Public Sub UnSetFormOnTop(myForm As Object)
    SetWindowPos myForm.hwnd, hwnd_NoTopMost, 0, 0, 0, 0, SWP_NoMove Or SWP_NoSize
End Sub

'Private Sub cmdOnTopOrNot_Click()
'    If x = 0 Then
'        SetFormOnTop Me
'        x = 1
'        cmdOnTopOrNot.Caption = "Always On Top - Enabled"
'    ElseIf x = 1 Then
'        UnSetFormOnTop Me
'        x = 0
'        cmdOnTopOrNot.Caption = "Always On Top - Disabled"
'    End If
'End Sub

'Private Sub Form_Load()
'    cmdOnTopOrNot.Caption = "Always On Top - Disabled"
'End Sub




