VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   25000
      Left            =   840
      Top             =   480
   End
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "New User"
      Height          =   390
      Left            =   1320
      TabIndex        =   6
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2520
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private booExitApp As Boolean
Private bytMaxAttempts As Byte

Private Sub cmdAddUser_Click()
    Me.Hide
    frmAddUsers.Show
End Sub

Private Sub cmdCancel_Click()
    booExitApp = True
    Me.Hide
    Unload Me
End Sub

Private Sub Decrypt()
  
End Sub

Private Sub cmdOK_Click()
On Error GoTo LocalTrap:
    
    Decrypt
    
    strSQL = "SELECT user_name, user_password FROM adm_users WHERE user_name='" & Me.txtUserName.Text & "' AND user_password='" & Me.txtPassword.Text & "';"
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)
        Do Until rst.EOF
            rst.MoveNext
        Loop
        rst.MoveFirst
    rst.Close
    Me.Timer1.Enabled = False
    strUserName = Me.txtUserName.Text
    strPassword = Me.txtPassword.Text
    frmMain.Show
    booExitApp = False
    Me.Hide
    Unload Me
    Exit Sub
LocalTrap:
    
    If bytMaxAttempts < 3 Then
        MsgBox "Incorrect Password and/or User Name!"
        bytMaxAttempts = bytMaxAttempts + 1
    Else
        MsgBox "You have entered an incorrect Password and/or User Name 3 times!  Exiting Application!", vbOKOnly + vbExclamation, "Exit"
        db.Close
        wks.Close
        End
    End If
    Me.txtPassword.Text = ""
    Me.txtUserName.Text = ""
    Me.txtUserName.SetFocus
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture(102, vbResIcon)
    LicAgreement
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If booExitApp = True Then
        basMain.ExitApp
        End
    End If
End Sub

Private Sub Timer1_Timer()
    MsgBox "Time out!", vbCritical + vbOKOnly, "Time Out"
    db.Close
    wks.Close
    End
End Sub
