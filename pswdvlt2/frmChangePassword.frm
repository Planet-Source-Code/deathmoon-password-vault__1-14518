VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Password"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCurrentPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm Password"
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "New Password"
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Current Password"
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public UsersPassword As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    ChangePassword
    basMain.ClearData Me
    Unload Me
End Sub

Private Sub Decrypt()

End Sub

Private Sub Encrypt()
    
End Sub

Private Sub ChangePassword()
    'Get user's current password.
    strSQL = "SELECT user_password FROM adm_users WHERE user_name='" & strUserName & "'"
    Set rst = db.OpenRecordset(strSQL, dbOpenDynaset)
        UsersPassword = rst.Fields(0)
    rst.Close
        
    Decrypt
        
    'Check to see if old password matches what the user typed.
    If UsersPassword <> Me.txtCurrentPassword.Text Then
        MsgBox "Invalid password!", vbCritical + vbOKOnly, "Incorrect Password"
    Else
        If Me.txtNewPassword.Text <> Me.txtConfirmPassword.Text Then
            Me.txtNewPassword.Text = ""
            Me.txtConfirmPassword.Text = ""
            MsgBox "New Password does not match verification.  Please re-type.", vbExclamation + vbOKOnly, "Verification"
        Else
            Encrypt
            strSQL = "UPDATE adm_users " _
                & "SET user_password ='" & Me.txtNewPassword.Text & "'" _
                & "WHERE user_name='" & strUserName & "';"
            db.Execute strSQL
            MsgBox "Password successfully changed", vbInformation + vbOKOnly, "Password Changed"
        End If
    End If
    UsersPassword = ""
End Sub
