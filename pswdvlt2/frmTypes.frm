VERSION 5.00
Begin VB.Form frmTypes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Types"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtNewType 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   2895
   End
   Begin VB.ListBox lisTypes 
      Height          =   2400
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Current Types"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If Me.txtNewType.Text <> "" Then
        Dim nResponse As Integer
        nResponse = MsgBox("Are you sure you want to add this type?", vbQuestion + vbYesNo, "Add Type")
        If nResponse = vbNo Then
            Exit Sub
        Else
            strSQL = "INSERT INTO pw_types (password_type) VALUES ('" & Me.txtNewType.Text & "');"
            db.Execute strSQL
            MsgBox "Type added", vbInformation + vbOKOnly, "Record Added"
            Me.txtNewType.Text = ""
            Me.lisTypes.Clear
            LoadData
        End If
    Else
        MsgBox "You must enter a type before adding a record.", vbExclamation + vbOKOnly, "Error"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LoadData
End Sub

Private Sub LoadData()
    Dim FieldName As String
    Dim TableName As String
    
    FieldName = "password_type"
    TableName = "pw_types"
    
    basMain.FillCombo lisTypes, db, FieldName, TableName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Refresh frmMain controls
    frmMain.cboPasswordTypes.Clear
    frmMain.cboTypes.Clear
    frmMain.lisPssEntries.Clear
    basMain.FillCombo frmMain.cboPasswordTypes, db, "password_type", "pw_types"
    basMain.FillCombo frmMain.cboTypes, db, "password_type", "pw_types"
    basMain.FillComboWhere frmMain.lisPssEntries, db, "title", "pw_data", "user_id", "=", strUserName
    
    Unload Me
End Sub
