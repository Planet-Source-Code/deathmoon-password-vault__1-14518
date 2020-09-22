Attribute VB_Name = "basMain"
Option Explicit

'Database Vars:
Public wks As Workspace
Public db As Database
Public rst As Recordset
Public strSQL As String

'Name / Path / Passwords
Private dbPath As String
Private dbName As String
Private dbPass As String

Public strUserName As String
Public strPassword As String
Public booLoadOnce As Boolean
Public bytCount As Byte

'Converted String
Public NewString As String

'basAlgorithm Variables
Public strToConvert As String
Public strEnDecrypted As String

'This allows you to "SLEEP" your application
Public Declare Sub Sleep Lib "kernel32" (ByVal _
    dwMilliseconds As Long)

'Used to get the logged in User's Name
Public Declare Function WNetGetUserA _
    Lib "mpr" (ByVal lpName As String, ByVal lpUserName As String, _
    lpnLength As Long) As Long

'Used for the GetUser routine
Public Function GetUser() As String
    Dim sUserNameBuff As String * 255
    sUserNameBuff = Space(255)
    Call WNetGetUserA(vbNullString, sUserNameBuff, 255&)
    GetUser = Left$(sUserNameBuff, InStr(sUserNameBuff, vbNullChar) - 1)
End Function

Public Sub Main()
    frmSplash.Show vbModal
    
    App.HelpFile = App.Path & "\Psswrd_Vlt.chm"
        
    dbName = "pwdata.mdb"
    dbPath = App.Path & "\" & dbName
    dbPass = ";pwd=KlduyDrVcN"
    
On Error GoTo RestoreBackup:
OpenDatabase:
    Set wks = DBEngine.Workspaces(0)
    Set db = wks.OpenDatabase(dbPath, False, False, dbPass)
    
    'If ZERO Users show the frmAddUsers Screen...
    LicAgreement
    If bytCount <> 0 Then
        frmLogin.Show
        Unload frmSplash
    End If
    Exit Sub
RestoreBackup:
    wks.Close
    GetRegistrySettings
    GoTo OpenDatabase:
End Sub

Public Sub FillCombo(ctl As Control, dbName As Database, _
    FieldName As String, TableName As String)
    ''Fills a combo box, list box, etc. with values that are passed to the
    'SQL statement below.
    strSQL = "SELECT " & FieldName & " FROM " & TableName & ""
    Set rst = dbName.OpenRecordset(strSQL)
    Do While Not rst.EOF
        ctl.AddItem rst.Fields(0).Value
        rst.MoveNext
    Loop
    rst.Close
End Sub

Public Sub FillComboWhere(ctl As Control, dbName As Database, _
    FieldName As String, TableName As String, _
    FilterByField As String, TypeOfFilter As String, FilterString As String)
    'Fills a combo box, list box, etc. with values that are passed to the
    'SQL statement below.
    'This routine has the WHERE Statement included.
    strSQL = "SELECT " & FieldName & " FROM " & TableName & " WHERE " & FilterByField & " " & TypeOfFilter & " '" & FilterString & "' GROUP BY " & FieldName
    Set rst = dbName.OpenRecordset(strSQL)
    Do While Not rst.EOF
        ctl.AddItem rst.Fields(0).Value
        rst.MoveNext
    Loop
    rst.Close
End Sub

Public Sub FillComboDoubleWhere(ctl As Control, dbName As Database, _
    FieldName As String, TableName As String, _
    FilterByField As String, TypeOfFilter As String, FilterString As String, _
    FilterByField2 As String, TypeOfFilter2 As String, FilterString2 As String)
    'Fills a combo box, list box, etc. with values that are passed to the
    'SQL statement below.
    'This routine has the WHERE Statement included.
    strSQL = "SELECT " & FieldName & " FROM " & TableName & " WHERE " & FilterByField & " " & TypeOfFilter & " '" & FilterString & "' AND " & FilterByField2 & " " & TypeOfFilter2 & " '" & FilterString2 & "' GROUP BY " & FieldName
    
    Set rst = dbName.OpenRecordset(strSQL)
    Do While Not rst.EOF
        ctl.AddItem rst.Fields(0).Value
        rst.MoveNext
    Loop
    rst.Close
End Sub

Public Sub ClearData(ByRef frm As Form)
    Dim ctl As Control
    Dim i As Byte
    For i = 0 To (frm.Controls.Count - 1)
        Set ctl = frm.Controls(i)
        'If it's a textbox, clear it
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next i
End Sub

Public Sub ValidateData(ByRef frm As Form)
    'If any of the controls on a form are NULL then
    'this places a 0 in that control.
    Dim ctl As Control
    Dim i As Byte
    For i = 0 To (frm.Controls.Count - 1)
        Set ctl = frm.Controls(i)
        'If it's a textbox, clear it
        If TypeOf ctl Is TextBox Then
            If ctl.Text = "" Then
                ctl.Text = "0"
            End If
        End If
    Next i
End Sub

Public Sub DBErrors()
    'Shows in detail ALL errors that occur to the program when
    'connecting to the database, executing queries, etc.
    Dim MyError As Error
    'shows a msgbox of how many errors occured
    MsgBox Errors.Count
    'shows a msgbox for each error that occured
    For Each MyError In DBEngine.Errors
        With MyError
            MsgBox .Number & " " & .Description
        End With
    Next MyError
End Sub

Sub SaveRegistrySettings(strDBBackUp As String)
    Dim strAppName As String    'name of the application in the registry
    Dim strSection As String    'section of the setting under AppName
    Dim strKey As String        'name of the key under which the setting
                                'will be saved
    Dim strRegValue As String   'Registry Value of the Key
    Dim lUsed As Long           '# of times program has been ran
    '
    strAppName = "Password Vault"
    strSection = "HH Data Solutions"
    '
    strKey = "AppUsed"
On Error GoTo LocalTrap:

    lUsed = GetSetting(strAppName, strSection, strKey)
    strRegValue = Val(lUsed) + 1
TryAgain:
    SaveSetting strAppName, strSection, strKey, strRegValue
    '
    'strKey = "UserID_Installed"
    'strRegValue = "Thomas Hudson"
    'SaveSetting strAppName, strSection, strKey, strRegValue
    '
    strKey = "LastRunDate"
    strRegValue = Date
    SaveSetting strAppName, strSection, strKey, strRegValue
    '
    strKey = "UserID_LastRun"
    strRegValue = GetUser
    SaveSetting strAppName, strSection, strKey, strRegValue
    '
    strKey = "DBBackUp"
    strRegValue = strDBBackUp
    SaveSetting strAppName, strSection, strKey, strRegValue
Exit Sub
LocalTrap:
    strRegValue = 1
    GoTo TryAgain
End Sub

Sub GetRegistrySettings()
    Dim strAppName As String    'name of the application in the registry
    Dim strSection As String    'section of the setting under AppName
    Dim strKey As String        'name of the key under which the setting
                                'will be saved
    Dim strRegValue As String   'Registry Value of the Key
    
    strAppName = "Password Vault"
    strSection = "HH Data Solutions"
    strKey = "DBBackUp"
    
    strRegValue = GetSetting(strAppName, strSection, strKey)
    
    MsgBox "Database is corrupted.  Will attempt to restore.", vbCritical + vbOKOnly, "Corrupt Database"

    FileCopy strRegValue, "pwdata.mdb"
    
    MsgBox "Database restored!", vbCritical + vbOKOnly, "Database Restored"
End Sub

Public Sub LicAgreement()
    strSQL = "SELECT Count(user_id) FROM adm_users"
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)
        bytCount = rst.Fields(0)
    rst.Close
    Set rst = Nothing
    strSQL = ""
    
    If bytCount = 0 Then
        frmAddUsers.Show
    End If
    
    'This is where you can place limitations on the
    '# of users the program will allow.
        'If bytCount = 1 Then
            'This is a single user license.
        '    Me.cmdAddUser.Visible = False
        'End If
        'If bytCount = 5 Then
            'This is a 5 user license.
        '    Me.cmdAddUser.Visible = False
        'End If
        'If bytCount = 10 Then
            'This is a 10 user license.
        '    Me.cmdAddUser.Visible = False
        'End If
End Sub

Sub BackUpDatabase()
    'This routine will backup the database in the following format
    'and save this file name in the registry.  It will also delete
    'the previous backup.
    'The naming convention of the backup is:  mmddyy.mdb
    Dim strBackup As String
    Dim strOldBackup As String
    Dim strAppName As String
    Dim strSection As String
    Dim strKey As String
    
    strAppName = "Password Vault"
    strSection = "HH Data Solutions"
    strKey = "DBBackUp"
    
    strOldBackup = GetSetting(strAppName, strSection, strKey)
    
    'Delete the Old Backup database...
On Error GoTo LocalTrap:
    Kill (App.Path & "\" & strOldBackup)
    
    strBackup = Format(Now(), "mmddyy") & ".mdb"
    FileCopy dbName, strBackup
    
    SaveRegistrySettings (strBackup)
    Exit Sub
LocalTrap:
    If Err.Number = 53 Then Resume Next
End Sub

Public Function fConvert(ByVal sStr As String) As String
    Dim i As Integer
    Dim sBadChar As String
    'List all illegal / unwanted characters
    sBadChar = "'"
    'Loop through all the characters of the string
    'checking whether each is an illegal character
    NewString = ""
    For i = 1 To Len(sStr)
        If InStr(sBadChar, Mid(sStr, i, 1)) = 0 Then
            NewString = NewString & Mid(sStr, i, 1)
        End If
    Next i
End Function

Sub ExitApp()
    'Close all connections to the database
    db.Close
    wks.Close
    'Backup database
    BackUpDatabase
End Sub
