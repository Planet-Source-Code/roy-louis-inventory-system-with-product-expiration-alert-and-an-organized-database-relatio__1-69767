VERSION 5.00
Begin VB.Form forgotPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forgot Password"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "forgotPass.frx":0000
   ScaleHeight     =   5025
   ScaleWidth      =   6345
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "forgotPass.frx":E44F2
      ScaleHeight     =   705
      ScaleWidth      =   6105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   6135
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "forgotPass.frx":E6898
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forgot Password"
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
         TabIndex        =   7
         Top             =   120
         Width           =   2400
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      Picture         =   "forgotPass.frx":E6C22
      ScaleHeight     =   3945
      ScaleWidth      =   6105
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   6135
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   120
         Picture         =   "forgotPass.frx":1CB114
         ScaleHeight     =   3105
         ScaleWidth      =   5865
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   5895
         Begin VB.TextBox password 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   240
            Locked          =   -1  'True
            PasswordChar    =   "•"
            TabIndex        =   3
            Top             =   2520
            Width           =   2655
         End
         Begin VB.TextBox username 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   240
            TabIndex        =   1
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox secA 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "•"
            TabIndex        =   2
            Top             =   1800
            Width           =   2655
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Your password is:"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   2280
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "You are required to provide your username and secret answer to retrieve your password."
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
            Height          =   585
            Left            =   240
            TabIndex        =   11
            Top             =   120
            Width           =   4305
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   840
            Width           =   795
         End
         Begin VB.Label secQ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Secret Question"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   1560
            Width           =   1140
         End
      End
      Begin VB.CommandButton Command2 
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
         Left            =   4560
         TabIndex        =   5
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Retrieve Password"
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
         Left            =   2760
         TabIndex        =   4
         Top             =   3360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "forgotPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Me.Command1.Caption = "Proceed to Log on" Then
        Load login
        login.username = Me.username
        login.password = Me.password
        MsgBox "You can change your password once you've logged on.", vbInformation, appTitle
        login.Show
        Unload Me
    Else
        RS.Open "SELECT * FROM users WHERE username = '" & Trim(Me.username.Text) & "'", DB, adOpenStatic, adLockOptimistic
            If RS.RecordCount > 0 Then
                With RS
                    If Trim(Me.secA.Text) = !secretA Then
                        Me.password = !password
                        Me.username.Enabled = False
                        Me.secA.Enabled = False
                        Me.password.Enabled = False
                        Me.Command1.Caption = "Proceed to Log on"
                    Else
                        MsgBox "Invalid secret answer. Please try again.", vbExclamation, appTitle
                        Me.secA.SetFocus
                    End If
                End With
            Else
                MsgBox "Unregistered username.", vbExclamation, appTitle
                Me.username.SetFocus
            End If
        RS.Close
    End If

End Sub

Private Sub Command2_Click()
    Unload Me
    Load login
    login.Show
End Sub

Private Sub secA_Change()
    Call username_Change
End Sub

Private Sub username_Change()
    If Me.username = Empty And Me.secA = Empty Then
        Me.Command1.Enabled = False
    Else
        Me.Command1.Enabled = True
    End If
    
    RS.Open "SELECT * FROM users WHERE username = '" & Trim(Me.username.Text) & "'", DB, adOpenStatic, adLockOptimistic
        If RS.RecordCount > 0 Then
            With RS
                Me.secQ = Questions(!secretQ + 1)
            End With
        Else
            Me.secQ.Caption = "Secret Question"
        End If
    RS.Close
End Sub
