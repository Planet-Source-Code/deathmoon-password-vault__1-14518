VERSION 5.00
Begin VB.Form frmAddUsers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Users"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtConfirm 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtUserName 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm Password"
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private booExit As Boolean
Private UserName As String
Private Password As String

Private Sub cmdAdd_Click()
    'Check to see if user name is already in database.  If so, add a 1 to the
    'end of it.
    If Me.txtUserName.Text = "" Then
        MsgBox "You must enter a User Name!", vbCritical + vbOKOnly, "User Name Required"
        Exit Sub
    End If
    If Me.txtConfirm.Text = Me.txtPassword.Text Then
        basCreateLogin.CreateLogin Me.txtUserName.Text
        Me.txtUserName.Text = NewLoginID
        AddUser
        booExit = False
        Unload Me
    Else
        MsgBox "Passwords do not match!", vbExclamation + vbOKOnly, "User Not Added"
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim nResponse As Integer
    
    basMain.LicAgreement
    
    If bytCount = 0 Then
        nResponse = MsgBox("You cannot use this program until you have entered a User Name and Password!", vbCritical + vbOKCancel, "Error")
        If nResponse = vbCancel Then
            booExit = True
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    If booExit = True Then
        basMain.ExitApp
        End
    Else
        frmLogin.Show
    End If
End Sub

Private Sub Encrypt()
    UserName = Me.txtUserName.Text
    Password = Me.txtPassword.Text
End Sub

Private Sub AddUser()
    
On Error GoTo LocalErr:
    'Encrypt the users of the database...
    Encrypt
    
    strSQL = "INSERT INTO adm_users " _
        & "(user_name, user_password) VALUES " _
        & "('" & UserName & "','" & Password & "');"
     db.Execute strSQL
     MsgBox "User Added!"
     basMain.ClearData Me
    Exit Sub
LocalErr:
    basMain.DBErrors
End Sub
