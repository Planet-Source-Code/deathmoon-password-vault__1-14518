VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Password Vault"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotes 
      Height          =   855
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ComboBox cboTypes 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3948
      Width           =   3375
   End
   Begin VB.TextBox txtUrl 
      Height          =   300
      Left            =   1080
      TabIndex        =   11
      Top             =   3591
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      Left            =   1080
      TabIndex        =   10
      Top             =   3234
      Width           =   3375
   End
   Begin VB.TextBox txtLogin 
      Height          =   300
      Left            =   1080
      TabIndex        =   9
      Top             =   2877
      Width           =   3375
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   2520
      Width           =   3375
   End
   Begin VB.ListBox lisPssEntries 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.ComboBox cboPasswordTypes 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Notes"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4560
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label5 
      Caption         =   "Login ID"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2916
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3969
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "URL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Double Click to open URL"
      Top             =   3618
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3267
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label lblpwTypes 
      Caption         =   "Password Types"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1335
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsAddTypes 
         Caption         =   "A&dd Types"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOptionsAlwaysOnTop 
         Caption         =   "Always On &Top"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShowPasswords 
         Caption         =   "&Show Password"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "&Change Password"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Contents..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOptionsAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strEncrypt As String
Private dblEntryID As Double
Private booFilterByTwo As Boolean

Private x As Byte

Private sLogin As String
Private sPassword As String
Private sURL As String
Private sNotes As String

Private Sub FilterbyTwo()
    Dim FieldName As String
    Dim TableName As String
    Dim FilterByField As String
    Dim FilterByField2 As String
    Dim TypeOfFilter As String
    Dim FilterString As String
    Dim FilterString2 As String

    FieldName = "title"
    TableName = "pw_data"
    FilterByField = "user_id"
    TypeOfFilter = "="
    FilterString = strUserName
    
    FilterByField2 = "pw_types_id"
    FilterString2 = Me.cboPasswordTypes.Text

    basMain.FillComboDoubleWhere lisPssEntries, db, FieldName, TableName, _
        FilterByField, TypeOfFilter, FilterString, FilterByField2, _
        TypeOfFilter, FilterString2

End Sub

Private Sub cboPasswordTypes_Click()
    If Me.cboPasswordTypes.Text <> "All" Then
        booFilterByTwo = True
        basMain.ClearData Me
        RefreshControls
        Me.cmdAdd.Enabled = True
        Me.cmdUpdate.Enabled = False
    Else
        booFilterByTwo = False
        basMain.ClearData Me
        RefreshControls
        Me.cmdAdd.Enabled = True
        Me.cmdUpdate.Enabled = False
    End If
End Sub

Private Sub StripCharacters()
    'This routine will check for bad characters before converting them
    'or updating the database with them.
    basMain.fConvert (Me.txtLogin.Text)
    Me.txtLogin.Text = NewString
    
    fConvert (Me.txtNotes.Text)
    Me.txtNotes.Text = NewString
    
    fConvert (Me.txtPassword.Text)
    Me.txtPassword.Text = NewString
    
    fConvert (Me.txtTitle.Text)
    Me.txtTitle.Text = NewString
    
    fConvert (Me.txtUrl.Text)
    Me.txtUrl.Text = NewString
End Sub

Private Sub cmdAdd_Click()
    
On Error GoTo LocalErr:
    If Me.txtTitle.Text = "" Then
        MsgBox "You must give a record a title!", vbCritical + vbOKOnly, "Error"
        Exit Sub
    Else
        'This will make all controls have a value in them.
        basMain.ValidateData Me
        If cboTypes.Text = "" Then
            MsgBox "You must enter a category for this entry!", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
        'After validating controls you can encrypt the data
        'to be stored.
        StripCharacters
        Encrypt
        'Update database with encrypted data.
        strSQL = "INSERT INTO pw_data " _
            & "(title, login_id, password, url, notes, pw_types_id, user_id) VALUES " _
            & "('" & Me.txtTitle.Text & "','" & sLogin & "','" & sPassword & "','" & sURL & "','" & sNotes & "','" & Me.cboTypes.Text & "','" & strUserName & "');"
         
        db.Execute strSQL
        MsgBox "Record Saved!"
        basMain.ClearData Me
        RefreshControls
        Me.txtTitle.SetFocus
    End If
    Exit Sub
LocalErr:
    basMain.DBErrors
End Sub

Private Function Encrypt()
    'Encrypt the following strings.
    'Login
    sLogin = Me.txtLogin.Text
    
    'Password
    sPassword = Me.txtPassword.Text
    
    'URL
    sURL = Me.txtUrl.Text
    
    'Notes
    sNotes = Me.txtNotes.Text
End Function

Private Function Decrypt()
   
End Function

Public Sub cmdExit_Click()
    Dim nResponse As Integer
    nResponse = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Exit Application")
    If nResponse = vbNo Then
        Exit Sub
    Else
        basMain.ExitApp
        With Me
            .WindowState = 1
            .Hide
        End With
        Unload frmTrayIcon
    End If
    End
End Sub

Private Sub cmdUpdate_Click()
    'Update Database
    StripCharacters
    Encrypt
On Error GoTo LocalTrap:
    strSQL = "UPDATE pw_data SET " _
        & "title='" & Me.txtTitle.Text & "'," _
        & "login_id='" & sLogin & "'," _
        & "password='" & sPassword & "'," _
        & "url='" & sURL & "'," _
        & "notes='" & sNotes & "'," _
        & "pw_types_id='" & Me.cboTypes.Text & "'," _
        & "user_id='" & strUserName & "'" _
        & "WHERE pw_data_id=" & dblEntryID & ";"
    
    db.Execute strSQL
    MsgBox "Record Updated!", vbInformation + vbOKOnly, "Updated"
    'Clear ALL controls
    ClearData Me
    'Refresh ALL controls
    RefreshControls
    Me.cmdUpdate.Enabled = False
    Me.cmdAdd.Enabled = True
    Exit Sub
LocalTrap:
    basMain.DBErrors
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture(101, vbResIcon)
    Load frmTrayIcon
    mnuShowPasswords_Click
    'mnuOptionsAlwaysOnTop_Click
On Error GoTo LocalTrap:
    Me.cmdUpdate.Enabled = False
    booLoadOnce = True
        LoadData
    booLoadOnce = False
    Exit Sub
LocalTrap:
    basMain.DBErrors
End Sub

Private Sub RefreshControls()
    With Me
        .lisPssEntries.Clear
        .cboTypes.Clear
    End With
    LoadData
End Sub

Private Sub LoadData()
    Dim FieldName As String
    Dim TableName As String
    Dim FilterByField As String
    Dim TypeOfFilter As String
    Dim FilterString As String
    
    Dim FilterByField2 As String
    Dim FilterString2 As String
    Dim TypeOfFilter2 As String
    
    If booFilterByTwo = False Then
        FieldName = "title"
        TableName = "pw_data"
        FilterByField = "user_id"
        TypeOfFilter = "="
        FilterString = strUserName
        If Me.cboPasswordTypes.Text = "All" Then
            basMain.FillComboWhere lisPssEntries, db, FieldName, _
                TableName, FilterByField, TypeOfFilter, FilterString
        Else
            FilterByField2 = "pw_types_id"
            FilterString2 = "Deleted"
            TypeOfFilter2 = "<>"
            
            basMain.FillComboDoubleWhere lisPssEntries, db, FieldName, TableName, _
                FilterByField, TypeOfFilter, FilterString, FilterByField2, _
                TypeOfFilter2, FilterString2
        End If
    Else
        FilterbyTwo
    End If
    FieldName = "password_type"
    TableName = "pw_types"
    If booLoadOnce = True Then
        basMain.FillCombo cboPasswordTypes, db, FieldName, TableName
    End If
    basMain.FillCombo cboTypes, db, FieldName, TableName
End Sub

Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then
        'If the window is minimzed then set
        'Restore to Enabled and Minimized to
        'disabled.
        frmTrayIcon.mnuMinimize.Enabled = False
        frmTrayIcon.mnuRestore.Enabled = True
    ElseIf Me.WindowState = vbNormal Then
        'if the window is restored then set
        'restore to disabled and minimized to
        'enabled.
        frmTrayIcon.mnuMinimize.Enabled = True
        frmTrayIcon.mnuRestore.Enabled = False
        Me.Width = 4725
        Me.Height = 6405
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim nResponse As Integer
    nResponse = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Exit Application")
    If nResponse = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        basMain.ExitApp
        With Me
            .WindowState = 1
            .Hide
        End With
        Unload frmTrayIcon
    End If
    End
End Sub

Private Sub Label3_DblClick()
    Dim sWebSite As String
    Dim sTmp As String
        
    sTmp = Mid(Me.txtUrl.Text, 1, 7)
        
    If sTmp = "http://" Then
        sWebSite = Me.txtUrl.Text
    Else
        sWebSite = "http://" & Me.txtUrl.Text
    End If
    
    Shell ("explorer " & sWebSite), vbNormalNoFocus
End Sub

Private Sub lisPssEntries_DblClick()
    Me.cmdUpdate.Enabled = True
    'When user double clicks on item in listbox have it
    'populate text boxes.
    Me.txtTitle.Text = Me.lisPssEntries.Text
    strSQL = "SELECT * FROM pw_data WHERE title='" & Me.lisPssEntries.Text & "';"
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)
        Me.txtLogin.Text = rst.Fields("login_id")
        Me.txtNotes.Text = rst.Fields("notes")
        Me.txtPassword.Text = rst.Fields("password")
        Me.txtUrl.Text = rst.Fields("url")
        Me.cboTypes.Text = rst.Fields("pw_types_id")
        dblEntryID = rst.Fields("pw_data_id")
    rst.Close
    Decrypt
    Me.cmdAdd.Enabled = False
End Sub

Private Sub mnuChangePassword_Click()
    frmChangePassword.Show vbModal
End Sub

Private Sub mnuOptionsAbout_Click()
    frmSplash.Show
End Sub

Private Sub mnuOptionsAddTypes_Click()
    frmTypes.Show vbModal
End Sub

Private Sub mnuOptionsAlwaysOnTop_Click()
    If x = 0 Then
        SetFormOnTop Me
        Me.mnuOptionsAlwaysOnTop.Checked = True
        x = 1
    ElseIf x = 1 Then
        UnSetFormOnTop Me
        Me.mnuOptionsAlwaysOnTop.Checked = False
        x = 0
    End If
End Sub

Private Sub mnuShowPasswords_Click()
    If Me.txtPassword.PasswordChar = "*" Then
        Me.txtPassword.PasswordChar = ""
        Me.mnuShowPasswords.Caption = "Hide Password"
        Me.mnuShowPasswords.Checked = False
    ElseIf Me.txtPassword.PasswordChar = "" Then
        Me.txtPassword.PasswordChar = "*"
        Me.mnuShowPasswords.Caption = "Show Password"
        Me.mnuShowPasswords.Checked = True
    End If
End Sub
