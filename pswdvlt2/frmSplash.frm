VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2760
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2670
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7140
      Begin VB.Timer Timer1 
         Interval        =   3500
         Left            =   2760
         Top             =   1560
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   675
         Width           =   2175
      End
      Begin VB.Label lblUrl 
         BackColor       =   &H00E0E0E0&
         Caption         =   "All Rights Reserved!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4560
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   240
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblAppName 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   270
         TabIndex        =   7
         Top             =   675
         Width           =   1665
      End
      Begin VB.Label lblFAX 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Made in the USA"
         Height          =   240
         Index           =   4
         Left            =   4560
         TabIndex        =   6
         Top             =   1200
         Width           =   2265
      End
      Begin VB.Label lblEMail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New England"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   4560
         MousePointer    =   4  'Icon
         TabIndex        =   5
         Top             =   1440
         Width           =   2490
      End
      Begin VB.Label lblCityS 
         BackColor       =   &H00E0E0E0&
         Caption         =   "www.planet-source-code.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   4560
         TabIndex        =   4
         Top             =   930
         Width           =   2265
      End
      Begin VB.Label lblPO 
         BackColor       =   &H00E0E0E0&
         Caption         =   "email: deathmoon91@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   4560
         TabIndex        =   3
         Top             =   705
         Width           =   2505
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Password Vault"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   270
         TabIndex        =   2
         Tag             =   "CompanyProduct"
         Top             =   270
         Width           =   2685
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Warning..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   120
         TabIndex        =   1
         Tag             =   "Warning"
         Top             =   1920
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Image1.Picture = LoadResPicture(101, vbResIcon)
    Me.lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub fraMainFrame_Click()
    Unload Me
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub lblCityS_Click(Index As Integer)
    Shell ("explorer http://www.planet-source-code.com"), vbNormalNoFocus
End Sub

Private Sub lblCompanyProduct_Click()
    Unload Me
End Sub

Private Sub lblProductName_Click(Index As Integer)
    Unload Me
End Sub

Private Sub lblPO_Click(Index As Integer)
    Shell ("explorer mailto:deathmoon91@yahoo.com"), vbNormalNoFocus
End Sub

Private Sub lblWarning_Click()
    Unload Me
End Sub

Private Sub picLogo_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
