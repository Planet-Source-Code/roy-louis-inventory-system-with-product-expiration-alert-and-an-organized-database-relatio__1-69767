VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form userAcct 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Accounts"
   ClientHeight    =   8190
   ClientLeft      =   2310
   ClientTop       =   1095
   ClientWidth     =   9270
   Icon            =   "userAcct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "userAcct.frx":1272
   ScaleHeight     =   8190
   ScaleWidth      =   9270
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "userAcct.frx":9642
      ScaleHeight     =   705
      ScaleWidth      =   8985
      TabIndex        =   55
      Top             =   120
      Width           =   9015
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "userAcct.frx":B9E8
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Accounts"
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
         Left            =   600
         TabIndex        =   56
         Top             =   120
         Width           =   2070
      End
   End
   Begin VB.PictureBox frameMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      Picture         =   "userAcct.frx":BD72
      ScaleHeight     =   2745
      ScaleWidth      =   8985
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   9015
      Begin VB.Shape box 
         Height          =   735
         Index           =   2
         Left            =   240
         Top             =   2280
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Shape box 
         Height          =   735
         Index           =   1
         Left            =   240
         Top             =   1560
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Shape box 
         Height          =   735
         Index           =   0
         Left            =   240
         Top             =   840
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label task 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   2
         Left            =   1200
         TabIndex        =   34
         Top             =   2520
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label task 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modify Existing Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   1
         Left            =   1200
         TabIndex        =   33
         Top             =   1800
         Width           =   2310
      End
      Begin VB.Label task 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create New Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   0
         Left            =   1200
         TabIndex        =   32
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Image menuIcon 
         Height          =   480
         Index           =   2
         Left            =   480
         Picture         =   "userAcct.frx":F0264
         Stretch         =   -1  'True
         Top             =   2400
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image menuIcon 
         Height          =   480
         Index           =   1
         Left            =   480
         Picture         =   "userAcct.frx":F05EE
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image menuIcon 
         Height          =   480
         Index           =   0
         Left            =   480
         Picture         =   "userAcct.frx":F0978
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick a task"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.PictureBox frameEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   120
      Picture         =   "userAcct.frx":F0D02
      ScaleHeight     =   7065
      ScaleWidth      =   8985
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   960
      Width           =   9015
      Begin VB.PictureBox EDIT_FRAME_confirmation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   1680
         Picture         =   "userAcct.frx":1D51F4
         ScaleHeight     =   2745
         ScaleWidth      =   5505
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5535
         Begin VB.CommandButton EDIT_cmdCANCEL1 
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
            Left            =   3720
            TabIndex        =   16
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton EDIT_cmdProceed 
            Caption         =   "Proceed"
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
            Left            =   2400
            TabIndex        =   15
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox confirmationPassword 
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
            Left            =   1680
            MaxLength       =   15
            PasswordChar    =   "•"
            TabIndex        =   14
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label Label12 
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
            Left            =   720
            TabIndex        =   41
            Top             =   1200
            Width           =   750
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "For security reason, you are required to provide your password below before you can proceed."
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
            Height          =   390
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   4845
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox EDIT_FRAME_userInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   120
         Picture         =   "userAcct.frx":2B96E6
         ScaleHeight     =   6225
         ScaleWidth      =   8625
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   600
         Width           =   8655
         Begin VB.OptionButton edit_accountType 
            BackColor       =   &H00CDC5B8&
            Caption         =   "Limited - User"
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
            Height          =   315
            Index           =   1
            Left            =   600
            TabIndex        =   27
            Top             =   5400
            Width           =   1455
         End
         Begin VB.OptionButton edit_accountType 
            BackColor       =   &H00CDC5B8&
            Caption         =   "Administrator"
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
            Height          =   315
            Index           =   0
            Left            =   600
            TabIndex        =   26
            Top             =   5040
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.TextBox edit_username 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDC5B8&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   330
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   64
            Text            =   "USERNAME"
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton EDIT_cmdCANCEL2 
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
            Left            =   7080
            TabIndex        =   29
            Top             =   5760
            Width           =   1335
         End
         Begin VB.CommandButton EDIT_cmdUpdate 
            Caption         =   "Update Account"
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
            Left            =   5520
            TabIndex        =   28
            Top             =   5760
            Width           =   1455
         End
         Begin VB.TextBox edit_secA 
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
            Left            =   600
            MaxLength       =   15
            PasswordChar    =   "•"
            TabIndex        =   25
            Top             =   4080
            Width           =   2655
         End
         Begin VB.ComboBox edit_secQ 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   3360
            Width           =   5655
         End
         Begin VB.TextBox edit_firstname 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   21
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox edit_lastname 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   600
            MaxLength       =   15
            TabIndex        =   20
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox edit_password2 
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
            Left            =   3600
            MaxLength       =   15
            PasswordChar    =   "•"
            TabIndex        =   23
            Top             =   2640
            Width           =   2655
         End
         Begin VB.TextBox edit_password 
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
            Left            =   600
            MaxLength       =   15
            PasswordChar    =   "•"
            TabIndex        =   22
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Level of Access"
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
            Left            =   480
            TabIndex        =   69
            Top             =   4680
            Width           =   1725
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Security Information"
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
            Left            =   480
            TabIndex        =   68
            Top             =   2040
            Width           =   1770
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
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
            Left            =   480
            TabIndex        =   67
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Secret Answer:"
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
            Left            =   600
            TabIndex        =   63
            Top             =   3840
            Width           =   1110
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Secret Question:"
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
            Left            =   600
            TabIndex        =   62
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name:"
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
            Left            =   3600
            TabIndex        =   61
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name:"
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
            Left            =   600
            TabIndex        =   60
            Top             =   1200
            Width           =   810
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password:"
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
            Left            =   3600
            TabIndex        =   59
            Top             =   2400
            Width           =   1350
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
            Left            =   600
            TabIndex        =   58
            Top             =   2400
            Width           =   750
         End
      End
      Begin VB.PictureBox EDIT_FRAME_usersList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   120
         Picture         =   "userAcct.frx":39DBD8
         ScaleHeight     =   4065
         ScaleWidth      =   8625
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   600
         Width           =   8655
         Begin VB.CommandButton EDIT_cmdPropoerties 
            Caption         =   "Properties"
            Enabled         =   0   'False
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
            Left            =   5400
            TabIndex        =   18
            Top             =   3480
            Width           =   1455
         End
         Begin VB.CommandButton EDIT_cmdCANCEL3 
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
            Left            =   6960
            TabIndex        =   19
            Top             =   3480
            Width           =   1335
         End
         Begin MSComctlLib.ListView edit_usersList 
            Height          =   2415
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   4260
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Use the list below to change password or other settings."
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
            Left            =   240
            TabIndex        =   66
            Top             =   240
            Width           =   4740
         End
      End
      Begin VB.Image Image4 
         Height          =   360
         Left            =   120
         Picture         =   "userAcct.frx":4820CA
         Stretch         =   -1  'True
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Account"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   600
         TabIndex        =   38
         Top             =   120
         Width           =   1725
      End
   End
   Begin VB.PictureBox frameAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   120
      Picture         =   "userAcct.frx":482454
      ScaleHeight     =   5625
      ScaleWidth      =   8985
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   960
      Width           =   9015
      Begin VB.PictureBox ADD_FRAME_UserInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   120
         Picture         =   "userAcct.frx":566946
         ScaleHeight     =   4785
         ScaleWidth      =   8625
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   600
         Width           =   8655
         Begin VB.CommandButton ADD_cmdCANCEL1 
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
            Left            =   7080
            TabIndex        =   8
            Top             =   4200
            Width           =   1335
         End
         Begin VB.CommandButton ADD_cmdNEXT 
            Caption         =   "Next >>"
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
            Left            =   5520
            TabIndex        =   7
            Top             =   4200
            Width           =   1455
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
            Left            =   480
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1800
            Width           =   2655
         End
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
            Left            =   480
            MaxLength       =   15
            PasswordChar    =   "•"
            TabIndex        =   4
            ToolTipText     =   "Alphanumeric values"
            Top             =   2400
            Width           =   2655
         End
         Begin VB.TextBox password2 
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
            Left            =   3480
            MaxLength       =   15
            PasswordChar    =   "•"
            TabIndex        =   30
            ToolTipText     =   "Alphanumeric values"
            Top             =   2400
            Width           =   2655
         End
         Begin VB.TextBox lastname 
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
            Left            =   480
            MaxLength       =   15
            TabIndex        =   1
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox firstname 
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
            Left            =   3480
            MaxLength       =   15
            TabIndex        =   2
            Top             =   1080
            Width           =   2655
         End
         Begin VB.ComboBox secQ 
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
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   3120
            Width           =   5775
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
            Left            =   480
            MaxLength       =   15
            PasswordChar    =   "•"
            TabIndex        =   6
            Top             =   3720
            Width           =   2655
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please provide the following information needed below."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   57
            Top             =   240
            Width           =   4665
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Left            =   480
            TabIndex        =   50
            Top             =   1560
            Width           =   840
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Left            =   480
            TabIndex        =   49
            Top             =   2160
            Width           =   750
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3480
            TabIndex        =   48
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   47
            Top             =   840
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3480
            TabIndex        =   46
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Secret Question:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   45
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Secret Answer:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   44
            Top             =   3480
            Width           =   1110
         End
      End
      Begin VB.PictureBox ADD_FRAME_ChooseAccType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   960
         Picture         =   "userAcct.frx":64AE38
         ScaleHeight     =   4665
         ScaleWidth      =   6825
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   720
         Width           =   6855
         Begin VB.CommandButton ADD_cmdBack 
            Caption         =   "<< Back"
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
            Left            =   2400
            TabIndex        =   11
            Top             =   4080
            Width           =   1215
         End
         Begin VB.CommandButton ADD_cmdCANCEL2 
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
            Left            =   5280
            TabIndex        =   13
            Top             =   4080
            Width           =   1335
         End
         Begin VB.CommandButton ADD_cmdCreateAcct 
            Caption         =   "Create Account"
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
            Left            =   3720
            TabIndex        =   12
            Top             =   4080
            Width           =   1455
         End
         Begin VB.OptionButton accountType 
            BackColor       =   &H00CDC5B8&
            Caption         =   "Limited - User"
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
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   10
            Top             =   2280
            Width           =   1695
         End
         Begin VB.OptionButton accountType 
            BackColor       =   &H00CDC5B8&
            Caption         =   "Administrator"
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
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   9
            Top             =   720
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"userAcct.frx":72F32A
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
            Height          =   1095
            Left            =   960
            TabIndex        =   54
            Top             =   2640
            Width           =   5415
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"userAcct.frx":72F411
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
            Height          =   1095
            Left            =   960
            TabIndex        =   53
            Top             =   1080
            Width           =   5415
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "What level of access do you want to grant this user?"
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
            Left            =   240
            TabIndex        =   52
            Top             =   240
            Width           =   4410
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create New Account"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   600
         TabIndex        =   36
         Top             =   120
         Width           =   2775
      End
      Begin VB.Image Image6 
         Height          =   360
         Left            =   120
         Picture         =   "userAcct.frx":72F4FF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   360
      End
   End
End
Attribute VB_Name = "userAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'=================================================================
'               INITIALIZE COMPONENTS' VALUES
'=================================================================
Private Sub Form_Load()
    'Show the form before loading contents
    Me.Show
    
    'Load Secret Questions
    loadSecretQ Me.secQ
    loadSecretQ Me.edit_secQ
        
    'Set FRAMES on their initial state
    initFrames
End Sub




'=================================================================
'                         TASK SELECTION
'                  (  CREATE NEW, EDIT, DELETE   )
'=================================================================
Private Sub task_Click(Index As Integer)
    Select Case Index
        'CREATE
        Case 0
            If user.level = 1 Then
                toggleFrame Me.frameADD
                clearADD
                Me.lastname.SetFocus
            Else
                MsgBox "You are not logged on as an Administrator. This action is not allowed.", vbInformation, appTitle
            End If
        
        'EDIT
        Case 1
            toggleFrame Me.frameEDIT
            clearEDIT
            Me.confirmationPassword.SetFocus
        
        'DELETE
        Case 2
    End Select
End Sub


'=================================================================
'                FRAME: CREATE NEW ACCOUNT
'=================================================================
Private Sub ADD_cmdNEXT_Click()
    'Check EMPTY fields
    If Me.lastname <> Empty And Me.firstname <> Empty _
    And Me.username <> Empty And Me.password <> Empty _
    And Me.secA <> Empty Then
        'Check USERNAME's availabilty
        If UserNameIsAvailable(Me.username) Then
            'Check PASSWORD confirmation
            If Trim(Me.password) = Trim(Me.password2) Then
                'PROCEED (Next >>)
                Me.ADD_FRAME_UserInfo.Visible = False
                Me.ADD_FRAME_ChooseAccType.Visible = True
            Else
                MsgBox "Password did not match.", vbExclamation, appTitle
            End If
        Else
            MsgBox "Username is no longer available", vbInformation, appTitle
        End If
    Else
        MsgBox "Please fill up all fields.", vbExclamation, appTitle
    End If
End Sub

Private Sub ADD_cmdBack_Click()
    Me.ADD_FRAME_ChooseAccType.Visible = False
    Me.ADD_FRAME_UserInfo.Visible = True
End Sub

Private Sub ADD_cmdCANCEL1_Click()
    If Me.lastname = Empty And Me.firstname = Empty _
    And Me.username = Empty And Me.password = Empty And Me.password2 = Empty _
    And Me.secA = Empty Then
        'do nothing, just cancel the transaction
        toggleFrame Me.frameMENU
    Else
        If MsgBox("Are you sure you want to cancel creating a new account?", vbQuestion + vbYesNo, appTitle) = vbYes Then toggleFrame Me.frameMENU
    End If
End Sub

Private Sub ADD_cmdCANCEL2_Click()
    If MsgBox("Are you sure you want to cancel creating a new account?", vbQuestion + vbYesNo, appTitle) = vbYes Then toggleFrame Me.frameMENU
End Sub

Private Sub ADD_cmdCreateAcct_Click()
    saveAccount
    MsgBox "New account has been successfully saved!" & vbNewLine _
    & "Username: " & Trim(Me.username) & vbNewLine _
    & "Account Type: " & IIf(Me.accountType(0).Value = True, "Adminstrator", "Limited - User") _
    , vbInformation, appTitle
    toggleFrame Me.frameMENU
End Sub


'=================================================================
'                FRAME: EDIT AN EXISTING ACCOUNT
'=================================================================
Private Sub EDIT_cmdCANCEL1_Click()
    toggleFrame Me.frameMENU
End Sub

Private Sub EDIT_cmdProceed_Click()
    If CStr(Me.confirmationPassword.Text) = CStr(user.password) Then
    
        If user.level = 1 Then
            Me.EDIT_FRAME_confirmation.Visible = False
            Me.EDIT_FRAME_userInfo.Visible = False
            Me.EDIT_FRAME_usersList.Visible = True
            Me.EDIT_FRAME_usersList.ZOrder vbBringToFront
            
            retrieveUsers
            
            If Me.edit_usersList.ListItems.Count > 0 Then Me.EDIT_cmdPropoerties.Enabled = True Else: Me _
            .EDIT_cmdPropoerties.Enabled = False
            
        Else
            Me.EDIT_FRAME_confirmation.Visible = False
            Me.EDIT_FRAME_usersList.Visible = False
            Me.EDIT_FRAME_userInfo.Visible = True
            Me.EDIT_FRAME_userInfo.ZOrder vbBringToFront
            
            retrieveUserInfo (user.username)
        End If
    Else
        MsgBox "Unauthorized Access!", vbExclamation, appTitle
        Me.confirmationPassword.SetFocus
    End If
End Sub

Private Sub EDIT_cmdCANCEL3_Click()
    clearEDIT
    Me.confirmationPassword.SetFocus
End Sub

Private Sub EDIT_cmdPropoerties_Click()
    Dim selectedUsername As String
    
    selectedUsername = Trim(Me.edit_usersList.SelectedItem.Text)
    Me.EDIT_FRAME_confirmation.Visible = False
    Me.EDIT_FRAME_usersList.Visible = False
    Me.EDIT_FRAME_userInfo.Visible = True
    
    retrieveUserInfo (selectedUsername)
    Me.edit_accountType(0).Enabled = True
    Me.edit_accountType(1).Enabled = True
    Me.edit_lastname.SetFocus
End Sub

Private Sub EDIT_cmdCANCEL2_Click()
    If user.level = 1 Then
        Me.EDIT_FRAME_confirmation.Visible = False
        Me.EDIT_FRAME_usersList.Visible = True
        Me.EDIT_FRAME_userInfo.Visible = False
    Else
        clearEDIT
    End If
End Sub

Private Sub EDIT_cmdUpdate_Click()
    'Check EMPTY fields
    If Me.edit_lastname <> Empty And Me.edit_firstname <> Empty _
    And Me.edit_username <> Empty And Me.edit_password <> Empty _
    And Me.edit_secA <> Empty Then
        'Check PASSWORD confirmation
        If Trim(Me.edit_password) = Trim(Me.edit_password2) Then
            'Update account (save changes)
            updateAccount Trim(Me.edit_username.Text)
            If user.username = Trim(Me.edit_username.Text) Then _
            updateUserInfoVariable Trim(Me.edit_username.Text)
            MsgBox "Account has been successfully updated.", vbInformation, appTitle
            
            retrieveUsers
            
            Me.EDIT_FRAME_confirmation.Visible = False
            Me.EDIT_FRAME_usersList.Visible = True
            Me.EDIT_FRAME_userInfo.Visible = False
        Else
            MsgBox "Password did not match.", vbExclamation, appTitle
        End If
    Else
        MsgBox "Please fill up all fields.", vbExclamation, appTitle
    End If
End Sub








'----------------------------------------------------------------------------------
'   Highlight Selected Task (onMouseOver)
'----------------------------------------------------------------------------------
'Show Box
Private Sub task_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Me.box(Index).Visible = True
End Sub
Private Sub menuIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call task_MouseMove(Index, Button, Shift, x, Y)
End Sub

'Hide Box
Private Sub frameMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Me.box(0).Visible = False
    Me.box(1).Visible = False
    Me.box(2).Visible = False
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

'Set frames on their initial state
Private Function initFrames()
    Me.frameMENU.Visible = True
    Me.frameADD.Visible = False
    Me.frameEDIT.Visible = False
    
    clearADD
    clearEDIT
End Function

'Show/Hide MAIN FRAMES (ADD,EDIT,DELETE ACCOUNT)
Private Function toggleFrame(showF As Object, Optional stillShowF As Object)
On Error Resume Next
    Dim FRAME_ As PictureBox
    
    Me.frameMENU.Visible = False
    Me.frameADD.Visible = False
    Me.frameEDIT.Visible = False
    
    Me.box(0).Visible = False
    Me.box(1).Visible = False
    Me.box(2).Visible = False
    
    Set FRAME_ = showF
    FRAME_.Visible = True
    FRAME_.ZOrder vbBringToFront
    
    Set FRAME_ = stillShowF
    FRAME_.Visible = True
    FRAME_.ZOrder vbBringToFront
End Function

'Retrieve Users and place it on the ListView Control
Private Function retrieveUsers()
    Dim item_ As ListItem
    Dim i As Long
    
    'Clear Headers and Lists
    Me.edit_usersList.ColumnHeaders.Clear
    Me.edit_usersList.ListItems.Clear
    
    'Set Column Headers
    With Me.edit_usersList.ColumnHeaders
        .Add Text:="Username", Width:=2300
        .Add Text:="Account Name", Width:=2800
        .Add Text:="Account Type", Width:=1500
    End With
    
    RS.Open "SELECT * FROM users", DB, adOpenStatic, adLockOptimistic
        With RS
            For i = 1 To .RecordCount Step 1
                Set item_ = Me.edit_usersList.ListItems.Add(Text:=!username)
                item_.SubItems(1) = !firstname & " " & !lastname
                item_.SubItems(2) = IIf(!accountType = 1, "Administrator", "Limited-User")
                .MoveNext
            Next i
        End With
    RS.Close
End Function

'Check if the text passed is already in the database (users!username)
Private Function UserNameIsAvailable(strUsername As String) As Boolean
    RS.Open "SELECT * FROM users WHERE username = '" & Trim(strUsername) & "'", DB, adOpenStatic, adLockOptimistic
        If RS.RecordCount > 0 Then UserNameIsAvailable = False Else UserNameIsAvailable = True
    RS.Close
    
    If strUsername = DEFAULT_USERNAME Then UserNameIsAvailable = False
End Function

'Save new account
Private Function saveAccount()
    RS.Open "SELECT * FROM users", DB, adOpenStatic, adLockOptimistic
        With RS
            .AddNew
                !lastname = Trim(Me.lastname)
                !firstname = Trim(Me.firstname)
                !username = Trim(Me.username)
                !password = Trim(Me.password)
                !secretQ = Me.secQ.ListIndex
                !secretA = Trim(Me.secA)
                !accountType = IIf(Me.accountType(0).Value = True, 1, 0)
            .Update
        End With
    RS.Close
End Function


'Update account
Private Function updateAccount(strUsername As String)
    RS.Open "SELECT * FROM users WHERE username ='" & strUsername & "'", DB, adOpenStatic, adLockOptimistic
        With RS
            .Update
                !lastname = Trim(Me.edit_lastname)
                !firstname = Trim(Me.edit_firstname)
                !password = Trim(Me.edit_password)
                !secretQ = Me.edit_secQ.ListIndex
                !secretA = Trim(Me.edit_secA)
                !accountType = IIf(Me.edit_accountType(0).Value = True, 1, 0)
            .Update
        End With
    RS.Close
End Function

'Load Secret Questions
Private Function loadSecretQ(onComboBox As Object)
On Error Resume Next
    Dim container As ComboBox
    Dim i As Integer
    
    Set container = onComboBox
    
    container.Clear
    For i = 1 To 5 Step 1
        container.AddItem Questions(i)
    Next i
    container.ListIndex = 0
End Function

'Retrieve user information (USED IN EDIT FRAME)
Private Function retrieveUserInfo(strUsername As String)
    RS.Open "SELECT * FROM users WHERE username = '" & strUsername & "'", DB, adOpenStatic, adLockOptimistic
        With RS
            Debug.Print .RecordCount
            Me.edit_username = !username
            Me.edit_lastname = !lastname
            Me.edit_firstname = !firstname
            Me.edit_password = !password
            Me.edit_password2 = !password
            Me.edit_secQ.ListIndex = !secretQ
            Me.edit_secA = !secretA
            Me.edit_accountType(0).Value = IIf(!accountType = 1, True, False)
            Me.edit_accountType(1).Value = IIf(!accountType = 0, True, False)
        End With
    RS.Close
End Function




'---------------------------------------
'  RESET FIELDS
'---------------------------------------
Private Function clearADD()
On Error Resume Next
    Me.lastname = Empty
    Me.firstname = Empty
    Me.username = Empty
    Me.password = Empty
    Me.password2 = Empty
    Me.secQ.ListIndex = 0
    Me.secA = Empty
    
    Me.accountType(0).Value = True
    Me.ADD_FRAME_ChooseAccType.Visible = False
    Me.ADD_FRAME_UserInfo.Visible = True
    Me.ADD_FRAME_UserInfo.ZOrder vbBringToFront
End Function
Private Function clearEDIT()
On Error Resume Next
    Me.edit_lastname = Empty
    Me.edit_firstname = Empty
    Me.edit_username = Empty
    Me.edit_password = Empty
    Me.edit_password2 = Empty
    Me.edit_secQ.ListIndex = 0
    Me.edit_secA = Empty
    Me.confirmationPassword = Empty
    Me.edit_usersList.ColumnHeaders.Clear
    Me.edit_usersList.ListItems.Clear
    Me.edit_accountType(0).Value = True
    Me.edit_accountType(0).Enabled = False
    Me.edit_accountType(1).Enabled = False
    
    Me.EDIT_FRAME_confirmation.Visible = True
    Me.EDIT_FRAME_userInfo.Visible = False
    Me.EDIT_FRAME_usersList.Visible = False
    Me.EDIT_FRAME_confirmation.ZOrder vbBringToFront
    
    Me.EDIT_cmdPropoerties.Enabled = False
End Function



