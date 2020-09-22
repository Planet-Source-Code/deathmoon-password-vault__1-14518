Attribute VB_Name = "basCreateLogin"
Option Explicit

Public NewLoginID As String

Public Sub CreateLogin(strLogin As String)
    Dim Login As String
    Dim m_tmp As String
    Dim x As String
    Dim y As Byte
    'Login has been created
    Login = strLogin
    'check employee table to see if login is in table
    strSQL = "SELECT user_name FROM adm_users;"
    y = 1
    Set rst = db.OpenRecordset(strSQL)
SearchAgain:
        Do While Not rst.EOF
            x = rst.Fields("user_name")
       
            x = Trim(x)
            If x = Login Then
                'Create new login
                Login = Login & y
                y = y + 1
                GoTo SearchAgain
                'Must Search through records again
                rst.MoveFirst
            End If
            rst.MoveNext
        Loop
    rst.Close
    If strLogin <> Login Then
        MsgBox "The username you entered was already taken.  Your new username is " & Login & ".", vbInformation + vbOKOnly, "New User Added"
    End If
    NewLoginID = Login
End Sub
