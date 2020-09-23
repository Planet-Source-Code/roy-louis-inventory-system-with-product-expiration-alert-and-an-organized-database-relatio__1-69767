VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Log On"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login.frx":0000
   ScaleHeight     =   3960
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      Picture         =   "login.frx":E44F2
      ScaleHeight     =   2865
      ScaleWidth      =   5865
      TabIndex        =   2
      Top             =   960
      Width           =   5895
      Begin VB.TextBox username 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   6
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox password 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdLogon 
         Caption         =   "Log On"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   1080
         TabIndex        =   10
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please provide your username and password below."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5445
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forgot Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   3720
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "login.frx":1C89E4
      ScaleHeight     =   705
      ScaleWidth      =   5865
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5895
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log On"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   720
         TabIndex        =   1
         Top             =   120
         Width           =   1020
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "login.frx":1CAD8A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdLogon_Click()
    If Authenticate = True Then
        Unload Me
        Load MDIForm1
        MDIForm1.Show
        Load expirationAlert
    Else
        MsgBox "Invalid Username or Password!", vbInformation, appTitle
    End If
End Sub


'==================================================================================
'                     *
'      *                     ~~~~~~~~~~~~~~~~~~~~~~           *
'                          /                        \
'                  ~~~~~~~~$    F U N C T I O N S   $~~~~~~~~              *
'                          \                        /
'         *                  ~~~~~~~~~~~~~~~~~~~~~~                 *
'
'==================================================================================

' CHECK USERNAME AND PASSWORD (AUTHENTICATION)
Public Function Authenticate() As Boolean
    Dim i As Long
    Dim holder_username As String
    Dim holder_password As String
    Dim holder_account_type As Integer
    Dim checker As Integer

    'CHECK FROM DATABASE
    RS.Open "SELECT * FROM users", DB, adOpenStatic, adLockOptimistic
        For i = 1 To RS.RecordCount Step 1
            With RS
                holder_username = .Fields("username").Value
                holder_password = .Fields("password").Value
                holder_account_type = .Fields("accountType").Value
                
                If Me.username.Text = holder_username And Me.password.Text = holder_password Then
                    Authenticate = True
                    checker = checker + 1
                    'Hold Current User's data
                    updateUserInfoVariable Trim(Me.username.Text)
                    Exit For
                End If
            End With
            RS.MoveNext
        Next i
    RS.Close


    'CHECK DEFAULT IF ACCT. DOES NOT EXIST IN THE DATABASE
    If checker = 0 Then
        If Me.username.Text = DEFAULT_USERNAME And Me.password.Text = DEFAULT_PASSWORD Then
            Authenticate = True
            'Hold Current User's data
            user.firstname = DEFAULT_FIRSTNAME
            user.lastname = DEFAULT_LASTNAME
            user.username = DEFAULT_USERNAME
            user.password = DEFAULT_PASSWORD
            user.level = DEFAULT_LEVEL
            checker = checker + 1
        End If
    End If


    'INVALID USER AND PASS
    If checker = 0 Then
        Authenticate = False
    End If
End Function


Private Sub Form_activate()
    Me.username.SetFocus
End Sub

Private Sub Form_Load()
    Me.Show
End Sub

Private Sub Label2_Click()
    Unload Me
    Load forgotPass
    forgotPass.Show
End Sub

Private Sub username_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdLogon_Click
End Sub

Private Sub password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdLogon_Click
End Sub
