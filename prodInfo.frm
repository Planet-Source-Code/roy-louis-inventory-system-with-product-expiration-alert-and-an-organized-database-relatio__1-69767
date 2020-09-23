VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form prodInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product and Inventory Management"
   ClientHeight    =   7605
   ClientLeft      =   4290
   ClientTop       =   2505
   ClientWidth     =   13380
   Icon            =   "prodInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "prodInfo.frx":1272
   ScaleHeight     =   7605
   ScaleWidth      =   13380
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "prodInfo.frx":9642
      ScaleHeight     =   705
      ScaleWidth      =   13065
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product and Inventory Management"
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
         TabIndex        =   45
         Top             =   120
         Width           =   5190
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "prodInfo.frx":656C4
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox frameEDIT 
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
      Height          =   5055
      Left            =   120
      Picture         =   "prodInfo.frx":65A4E
      ScaleHeight     =   5025
      ScaleWidth      =   6945
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   960
      Width           =   6975
      Begin VB.PictureBox EDIT_FRAME_confirmation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   720
         Picture         =   "prodInfo.frx":6DE1E
         ScaleHeight     =   2745
         ScaleWidth      =   5505
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5535
         Begin VB.TextBox edit_confirmationPassword 
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
            Left            =   1680
            MaxLength       =   15
            PasswordChar    =   "•"
            TabIndex        =   8
            Top             =   1200
            Width           =   3015
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
            Left            =   2280
            TabIndex        =   9
            Top             =   2040
            Width           =   1215
         End
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
            Left            =   3600
            TabIndex        =   10
            Top             =   2040
            Width           =   1095
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
            TabIndex        =   69
            Top             =   240
            Width           =   4845
            WordWrap        =   -1  'True
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
            TabIndex        =   68
            Top             =   1200
            Width           =   750
         End
      End
      Begin VB.PictureBox EDIT_FRAME_ProdInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   120
         Picture         =   "prodInfo.frx":761EE
         ScaleHeight     =   3705
         ScaleWidth      =   6585
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   600
         Width           =   6615
         Begin VB.TextBox edit_prodName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   14
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox edit_prodNameHolder 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   102
            Top             =   720
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox edit_prodUnitPrice 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   17
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox edit_prodSize 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   16
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox edit_prodDesc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   1320
            Width           =   3615
         End
         Begin VB.CommandButton EDIT_cmdUpdate 
            Caption         =   "Update Product Information"
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
            TabIndex        =   18
            Top             =   3240
            Width           =   2295
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
            Left            =   5160
            TabIndex        =   19
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox edit_prodID 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Text            =   "PRODUCT_ID"
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product ID:"
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
            Left            =   480
            TabIndex        =   110
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name:"
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
            Left            =   480
            TabIndex        =   76
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Price:"
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
            Left            =   480
            TabIndex        =   75
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size:"
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
            Left            =   480
            TabIndex        =   74
            Top             =   1920
            Width           =   345
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
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
            Left            =   480
            TabIndex        =   73
            Top             =   1320
            Width           =   855
         End
      End
      Begin VB.PictureBox EDIT_FRAME_prodList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   120
         Picture         =   "prodInfo.frx":7E5BE
         ScaleHeight     =   4065
         ScaleWidth      =   6585
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   600
         Width           =   6615
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
            Left            =   5040
            TabIndex        =   13
            Top             =   3600
            Width           =   1335
         End
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
            Left            =   3480
            TabIndex        =   12
            Top             =   3600
            Width           =   1455
         End
         Begin MSComctlLib.ListView edit_prodList 
            Height          =   2775
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   4895
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
            Caption         =   "Use the list below to modify the name and other product information."
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
            TabIndex        =   71
            Top             =   240
            Width           =   5850
         End
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modify Product Information"
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
         TabIndex        =   72
         Top             =   120
         Width           =   3675
      End
      Begin VB.Image Image4 
         Height          =   360
         Left            =   120
         Picture         =   "prodInfo.frx":8698E
         Stretch         =   -1  'True
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox frameADD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   120
      Picture         =   "prodInfo.frx":86D18
      ScaleHeight     =   4785
      ScaleWidth      =   7425
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   960
      Width           =   7455
      Begin VB.PictureBox ADD_FRAME_ADDPRODUCT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   120
         Picture         =   "prodInfo.frx":8F0E8
         ScaleHeight     =   3945
         ScaleWidth      =   7065
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   600
         Width           =   7095
         Begin VB.TextBox add_prodDesc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1560
            Width           =   3495
         End
         Begin VB.TextBox add_prodSize 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   4
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox add_prodUnitPrice 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   5
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox add_prodName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   2
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox add_prodID 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   840
            Width           =   3495
         End
         Begin VB.CommandButton add_cmd_save 
            Caption         =   "Save New Product"
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
            TabIndex        =   6
            Top             =   3360
            Width           =   1695
         End
         Begin VB.CommandButton add_cmd_cancel 
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
            Left            =   5520
            TabIndex        =   7
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
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
            Left            =   480
            TabIndex        =   63
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size:"
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
            Left            =   480
            TabIndex        =   62
            Top             =   2400
            Width           =   345
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Price:"
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
            Left            =   480
            TabIndex        =   61
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product ID No.:"
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
            Left            =   480
            TabIndex        =   60
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name:"
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
            Left            =   480
            TabIndex        =   54
            Top             =   1200
            Width           =   1065
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
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   240
            TabIndex        =   53
            Top             =   240
            Width           =   4665
         End
      End
      Begin VB.Image Image6 
         Height          =   360
         Left            =   120
         Picture         =   "prodInfo.frx":974B8
         Stretch         =   -1  'True
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Product"
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
         TabIndex        =   55
         Top             =   120
         Width           =   2340
      End
   End
   Begin VB.PictureBox frameVIEWSTOCKS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   120
      Picture         =   "prodInfo.frx":97842
      ScaleHeight     =   6465
      ScaleWidth      =   13065
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   960
      Width           =   13095
      Begin VB.CommandButton VIEWSTOCKS_cmdClose 
         Caption         =   "Close"
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
         Left            =   11520
         TabIndex        =   38
         Top             =   5880
         Width           =   1335
      End
      Begin VB.PictureBox VIEWSTOCKS_FRAME_inventoryList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   120
         Picture         =   "prodInfo.frx":9FC12
         ScaleHeight     =   5145
         ScaleWidth      =   12705
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   600
         Width           =   12735
         Begin VB.TextBox VIEWSTOCKS_allStocks 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox VIEWSTOCKS_nonExpiredStocks 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox VIEWSTOCKS_expiredStocks 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   405
            Left            =   9600
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox VIEWSTOCKS_prodDesc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   3600
            Width           =   4815
         End
         Begin MSComctlLib.ListView VIEWSTOCKS_inventoryList 
            Height          =   2535
            Left            =   5520
            TabIndex        =   37
            Top             =   2280
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   4471
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
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
         Begin MSComctlLib.ListView VIEWSTOCKS_prodList 
            Height          =   2295
            Left            =   240
            TabIndex        =   32
            Top             =   960
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   4048
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
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
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Stocks:"
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
            Left            =   5520
            TabIndex        =   109
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Non-Expired Stocks:"
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
            Left            =   7560
            TabIndex        =   108
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expired Stocks:"
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
            Left            =   9600
            TabIndex        =   107
            Top             =   1080
            Width           =   1110
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inventory Stocks Summary"
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
            Left            =   5520
            TabIndex        =   106
            Top             =   720
            Width           =   2340
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Description"
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
            TabIndex        =   105
            Top             =   3360
            Width           =   1665
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Products List"
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
            TabIndex        =   104
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inventory Stocks Details"
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
            Left            =   5520
            TabIndex        =   103
            Top             =   2040
            Width           =   2100
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check the inventory stocks of each product by clicking an item below."
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
            TabIndex        =   100
            Top             =   240
            Width           =   5880
         End
      End
      Begin VB.Image Image5 
         Height          =   360
         Left            =   120
         Picture         =   "prodInfo.frx":A7FE2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Stocks in Inventory"
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
         TabIndex        =   101
         Top             =   120
         Width           =   3345
      End
   End
   Begin VB.PictureBox frameMENU 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      Picture         =   "prodInfo.frx":A856C
      ScaleHeight     =   4305
      ScaleWidth      =   10185
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   960
      Width           =   10215
      Begin VB.Shape box 
         Height          =   735
         Index           =   4
         Left            =   5160
         Top             =   1920
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label task 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Inventory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   4
         Left            =   6120
         TabIndex        =   59
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Image menuIcon 
         Height          =   480
         Index           =   4
         Left            =   5400
         Picture         =   "prodInfo.frx":B093C
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Management"
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
         Left            =   5160
         TabIndex        =   58
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Information"
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
         Left            =   240
         TabIndex        =   57
         Top             =   840
         Width           =   1965
      End
      Begin VB.Shape box 
         Height          =   735
         Index           =   3
         Left            =   5160
         Top             =   1200
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label task 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Stocks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   3
         Left            =   6120
         TabIndex        =   56
         Top             =   1440
         Width           =   945
      End
      Begin VB.Image menuIcon 
         Height          =   480
         Index           =   3
         Left            =   5400
         Picture         =   "prodInfo.frx":B0CC6
         Stretch         =   -1  'True
         Top             =   1320
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
         TabIndex        =   50
         Top             =   240
         Width           =   1395
      End
      Begin VB.Image menuIcon 
         Height          =   480
         Index           =   0
         Left            =   480
         Picture         =   "prodInfo.frx":B1050
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image menuIcon 
         Height          =   480
         Index           =   1
         Left            =   480
         Picture         =   "prodInfo.frx":B13DA
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image menuIcon 
         Height          =   480
         Index           =   2
         Left            =   480
         Picture         =   "prodInfo.frx":B1764
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   480
      End
      Begin VB.Label task 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Product"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   0
         Left            =   1200
         TabIndex        =   49
         Top             =   1440
         Width           =   1470
      End
      Begin VB.Label task 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modify Existing Product Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   1
         Left            =   1200
         TabIndex        =   48
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label task 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remove an Existing Product"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   2
         Left            =   1200
         TabIndex        =   47
         Top             =   2880
         Width           =   2370
      End
      Begin VB.Shape box 
         Height          =   735
         Index           =   0
         Left            =   240
         Top             =   1200
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Shape box 
         Height          =   735
         Index           =   1
         Left            =   240
         Top             =   1920
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Shape box 
         Height          =   735
         Index           =   2
         Left            =   240
         Top             =   2640
         Visible         =   0   'False
         Width           =   4695
      End
   End
   Begin VB.PictureBox frameADDSTOCKS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   120
      Picture         =   "prodInfo.frx":B1AEE
      ScaleHeight     =   5625
      ScaleWidth      =   8025
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   960
      Width           =   8055
      Begin VB.PictureBox ADDSTOCKS_FRAME_addStock 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   120
         Picture         =   "prodInfo.frx":B9EBE
         ScaleHeight     =   4785
         ScaleWidth      =   7665
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   600
         Width           =   7695
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   111
            Top             =   1560
            Width           =   1455
         End
         Begin VB.ComboBox ADDSTOCKS_supplier 
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
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   2880
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker ADDSTOCKS_expiration 
            Height          =   375
            Left            =   1800
            TabIndex        =   29
            Top             =   3240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   57540609
            CurrentDate     =   39430
         End
         Begin MSComCtl2.DTPicker ADDSTOCKS_deliveryDate 
            Height          =   375
            Left            =   1800
            TabIndex        =   26
            Top             =   2040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   57540609
            CurrentDate     =   39430
         End
         Begin VB.TextBox ADDSTOCKS_Quantity 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   27
            Text            =   "1"
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox ADDSTOCKS_currentStock 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   3360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CommandButton ADDSTOCKS_cmdCancel2 
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
            Left            =   6120
            TabIndex        =   31
            Top             =   4200
            Width           =   1335
         End
         Begin VB.CommandButton ADDSTOCKS_cmdAddStock 
            Caption         =   "Add Stock"
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
            Left            =   4320
            TabIndex        =   30
            Top             =   4200
            Width           =   1695
         End
         Begin VB.TextBox ADDSTOCKS_prodID 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox ADDSTOCKS_prodName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox ADDSTOCKS_prodPrice 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1560
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expiration Date:"
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
            Left            =   480
            TabIndex        =   96
            Top             =   3240
            Width           =   1170
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier:"
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
            Left            =   480
            TabIndex        =   95
            Top             =   2880
            Width           =   630
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Date:"
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
            Left            =   480
            TabIndex        =   94
            Top             =   2040
            Width           =   1035
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity:"
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
            Left            =   480
            TabIndex        =   93
            Top             =   2520
            Width           =   690
         End
         Begin VB.Label Label29 
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
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   240
            TabIndex        =   89
            Top             =   240
            Width           =   4665
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name:"
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
            Left            =   480
            TabIndex        =   88
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product ID No.:"
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
            Left            =   480
            TabIndex        =   87
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Price:"
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
            Left            =   480
            TabIndex        =   86
            Top             =   1560
            Width           =   735
         End
      End
      Begin VB.PictureBox ADDSTOCKS_FRAME_prodList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   120
         Picture         =   "prodInfo.frx":C228E
         ScaleHeight     =   4185
         ScaleWidth      =   7665
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   600
         Width           =   7695
         Begin VB.CommandButton ADDSTOCKS_cmdCancel1 
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
            Left            =   6120
            TabIndex        =   22
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton ADDSTOCKS_cmdUpdateStock 
            Caption         =   "Update Stock"
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
            Left            =   4440
            TabIndex        =   21
            Top             =   3600
            Width           =   1575
         End
         Begin MSComctlLib.ListView ADDSTOCKS_prodList 
            Height          =   2655
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   4683
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
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select the item you want to add stocks in the inventory."
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
            TabIndex        =   92
            Top             =   240
            Width           =   4740
         End
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Stocks"
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
         TabIndex        =   90
         Top             =   120
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   360
         Left            =   120
         Picture         =   "prodInfo.frx":CA65E
         Stretch         =   -1  'True
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox frameDELETE 
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
      Height          =   4935
      Left            =   120
      Picture         =   "prodInfo.frx":CA9E8
      ScaleHeight     =   4905
      ScaleWidth      =   7065
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   960
      Width           =   7095
      Begin VB.PictureBox DELETE_FRAME_prodLIST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   120
         Picture         =   "prodInfo.frx":D2DB8
         ScaleHeight     =   4065
         ScaleWidth      =   6705
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   600
         Width           =   6735
         Begin VB.CommandButton DELETE_cmdDelete 
            Caption         =   "Remove Product"
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
            Left            =   3600
            TabIndex        =   43
            Top             =   3600
            Width           =   1575
         End
         Begin VB.CommandButton DELETE_cmdCancel2 
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
            TabIndex        =   44
            Top             =   3600
            Width           =   1335
         End
         Begin MSComctlLib.ListView DELETE_prodList 
            Height          =   2775
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   4895
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
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select the product you want to remove."
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
            TabIndex        =   82
            Top             =   240
            Width           =   3360
         End
      End
      Begin VB.PictureBox DELETE_FRAME_confirmation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   840
         Picture         =   "prodInfo.frx":DB188
         ScaleHeight     =   2745
         ScaleWidth      =   5505
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5535
         Begin VB.CommandButton DELETE_cmdCancel1 
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
            Left            =   3600
            TabIndex        =   41
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton DELETE_cmdProceed 
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
            Left            =   2280
            TabIndex        =   40
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox DELETE_confirmationPassword 
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
            Left            =   1680
            MaxLength       =   15
            PasswordChar    =   "•"
            TabIndex        =   39
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label Label20 
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
            TabIndex        =   80
            Top             =   1200
            Width           =   750
         End
         Begin VB.Label Label17 
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
            TabIndex        =   79
            Top             =   240
            Width           =   4845
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   120
         Picture         =   "prodInfo.frx":1BF67A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remove Product"
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
         TabIndex        =   83
         Top             =   120
         Width           =   2220
      End
   End
End
Attribute VB_Name = "prodInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub add_prodSize_Change()
    If Me.add_prodSize = Empty Then Me.add_prodSize = 1
End Sub
Private Sub add_prodSize_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 58       'Numerci [0-9]
        Case 8              'Backspace
        Case 13             'Enter Key
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub add_prodUnitPrice_Change()
    If Me.add_prodUnitPrice = Empty Then Me.add_prodUnitPrice = 0
    Me.add_prodUnitPrice = Format(Me.add_prodUnitPrice, "##,##0.00")
End Sub
Private Sub add_prodUnitPrice_GotFocus()
    Me.add_prodUnitPrice.SelStart = 0
    Me.add_prodUnitPrice.SelLength = Len(Me.add_prodUnitPrice)
End Sub
Private Sub add_prodUnitPrice_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 58       'Numerci [0-9]
        Case 8              'Backspace
        Case 13             'Enter Key
        Case Else
            KeyAscii = 0
    End Select
End Sub



Private Sub ADDSTOCKS_prodPrice_Change()
    Me.Text1 = "Php " & Format(ADDSTOCKS_prodPrice.Text, "##0.00")
End Sub

Private Sub ADDSTOCKS_Quantity_Change()
    If Me.ADDSTOCKS_Quantity = Empty Then Me.ADDSTOCKS_Quantity = 1
End Sub

Private Sub ADDSTOCKS_Quantity_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 58       'Numerci [0-9]
        Case 8              'Backspace
        Case 13             'Enter Key
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub edit_prodSize_Change()
    If Me.edit_prodSize = Empty Then Me.edit_prodSize = 1
End Sub
Private Sub edit_prodSize_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 58       'Numerci [0-9]
        Case 8              'Backspace
        Case 13             'Enter Key
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub edit_prodUnitPrice_Change()
    If Me.edit_prodUnitPrice = Empty Then Me.edit_prodUnitPrice = 0
    Me.edit_prodUnitPrice = Format(Me.edit_prodUnitPrice, "##,##0.00")
End Sub
Private Sub edit_prodUnitPrice_GotFocus()
    Me.edit_prodUnitPrice.SelStart = 0
    Me.edit_prodUnitPrice.SelLength = Len(Me.edit_prodUnitPrice)
End Sub
Private Sub edit_prodUnitPrice_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 58       'Numerci [0-9]
        Case 8              'Backspace
        Case 13             'Enter Key
        Case Else
            KeyAscii = 0
    End Select
End Sub



'=================================================================
'               INITIALIZE COMPONENTS' VALUES
'=================================================================
Private Sub Form_Load()
    'Show the form before loading contents
    Me.Show
    
    'Set initial state of frames
    initFrames
    
    'Load suppliers
    loadSuppliers ADDSTOCKS_supplier
End Sub












'=================================================================
'                         TASK SELECTION
'                  (  XXXXXXXXXXXXXXXXXXXXXXXX  )
'=================================================================
Private Sub task_Click(Index As Integer)
    Select Case Index
        'CREATE NEW PRODUCT
        Case 0
            If user.level = 1 Then
                clearADD
                toggleFrame Me.frameADD
                Me.add_prodID = generateProdID
                Me.add_prodName.SetFocus
            Else
                MsgBox "You are not logged on as an Administrator. This action is prohibited.", vbInformation, appTitle
            End If
            
            
        'MODIFY EXISTING PROD
        Case 1
            If user.level = 1 Then
                clearEDIT
                toggleFrame Me.frameEDIT
                Me.edit_confirmationPassword.SetFocus
            Else
                MsgBox "You are not logged on as an Administrator. This action is prohibited.", vbInformation, appTitle
            End If
        
        'DELETE EXISTING PROD
        Case 2
            If user.level = 1 Then
                clearDELETE
                toggleFrame Me.frameDELETE
                Me.DELETE_confirmationPassword.SetFocus
            Else
                MsgBox "You are not logged on as an Administrator. This action is prohibited.", vbInformation, appTitle
            End If
            
        'ADD STOCKS
        Case 3
            clearADDSTOCKS
            toggleFrame Me.frameADDSTOCKS
            retrieveProducts3
            Me.ADDSTOCKS_cmdUpdateStock.Enabled = _
            IIf(Me.ADDSTOCKS_prodList.ListItems.Count = 0, False, True)
            Me.ADDSTOCKS_prodList.SetFocus
        
        'VIEW STOCKS
        Case 4
            clearVIEWSTOCKS
            toggleFrame Me.frameVIEWSTOCKS
            retrieveProducts4
            Me.VIEWSTOCKS_prodList.SetFocus
    End Select
End Sub











'=================================================================
'                FRAME: CREATE NEW PRODUCT
'=================================================================
Private Sub add_cmd_save_Click()
    'Check Fields
    If Me.add_prodName <> Empty And _
    Me.add_prodDesc <> Empty And _
    Me.add_prodSize <> Empty And _
    Me.add_prodUnitPrice <> Empty Then
        'Check Unit Price
        If CDbl(Me.add_prodUnitPrice) > 0 Then
            'Check product name
            If isIdentical(Trim(Me.add_prodName)) = False Then
                saveProduct
                MsgBox "New product has been successfully saved!", vbInformation, appTitle
                toggleFrame Me.frameMENU
            Else
                If MsgBox("The product name """ & Trim(Me.add_prodName) & """ already exists in the database. Continue anyway?", vbInformation + vbYesNo) = vbYes Then
                    saveProduct
                    MsgBox "New product has been successfully saved!", vbInformation, appTitle
                    toggleFrame Me.frameMENU
                Else
                    'Provide a new one hehehe
                End If
            End If
        Else
            MsgBox "Please specify a valid unit price for this product.", vbExclamation, appTitle
        End If
    Else
        MsgBox "Please fill up all fields.", vbExclamation, appTitle
    End If
End Sub

Private Sub add_cmd_cancel_Click()
    If Me.add_prodName <> Empty Or _
    Me.add_prodDesc <> Empty Or _
    Me.add_prodSize <> 1 Or _
    Me.add_prodUnitPrice <> 0 Then
        If MsgBox("Are you sure you want to cancel creating a new product information?", vbQuestion + vbYesNo, appTitle) = vbYes Then toggleFrame Me.frameMENU
    Else
        toggleFrame Me.frameMENU
    End If
End Sub









'=================================================================
'                FRAME: MODIFY EXISTING PROD. INFO.
'=================================================================
Private Sub EDIT_cmdProceed_Click()
    If CStr(Me.edit_confirmationPassword) = CStr(user.password) Then
        Me.EDIT_FRAME_confirmation.Visible = False
        Me.EDIT_FRAME_prodList.Visible = True
        Me.EDIT_FRAME_ProdInfo.Visible = False
        Me.EDIT_FRAME_prodList.ZOrder vbBringToFront

        retrieveProducts
        
        If Me.edit_prodList.ListItems.Count > 0 Then Me.EDIT_cmdPropoerties.Enabled = True Else: Me _
        .EDIT_cmdPropoerties.Enabled = False
        Me.edit_prodList.SetFocus
    Else
        MsgBox "Unauthorized Access!", vbExclamation, appTitle
        Me.edit_confirmationPassword.SetFocus
    End If
End Sub

Private Sub EDIT_cmdCANCEL1_Click()
    toggleFrame Me.frameMENU
End Sub

Private Sub EDIT_cmdPropoerties_Click()
    Dim selectedProd As String
    
    selectedProd = Trim(Me.edit_prodList.SelectedItem.Text)
    Me.EDIT_FRAME_confirmation.Visible = False
    Me.EDIT_FRAME_prodList.Visible = False
    Me.EDIT_FRAME_ProdInfo.Visible = True
    
    retrieveProdInfo (selectedProd)
    Me.edit_prodName.SetFocus
End Sub

Private Sub EDIT_cmdUpdate_Click()
    'Check Fields
    If Me.edit_prodName <> Empty And _
    Me.edit_prodDesc <> Empty And _
    Me.edit_prodSize <> Empty And _
    Me.edit_prodUnitPrice <> Empty Then
        'Check Unit Price
        If CDbl(Me.edit_prodUnitPrice) > 0 Then
            'Check product name
            If isIdentical(Trim(Me.edit_prodName), Trim(Me.edit_prodNameHolder)) = False Then
                updateProduct Me.edit_prodID
                MsgBox "Product information has been successfully updated!", vbInformation, appTitle
                retrieveProducts
                Me.EDIT_FRAME_confirmation.Visible = False
                Me.EDIT_FRAME_ProdInfo.Visible = False
                Me.EDIT_FRAME_prodList.Visible = True
                Me.edit_prodList.SetFocus
            Else
                If MsgBox("The product name """ & Trim(Me.edit_prodName) & """ already exists in the database. Continue anyway?", vbInformation + vbYesNo, appTitle) = vbYes Then
                    updateProduct Me.edit_prodID
                    MsgBox "Product information has been successfully updated!", vbInformation, appTitle
                    retrieveProducts
                    Me.EDIT_FRAME_confirmation.Visible = False
                    Me.EDIT_FRAME_ProdInfo.Visible = False
                    Me.EDIT_FRAME_prodList.Visible = True
                    Me.edit_prodList.SetFocus
                Else
                    'Provide a new one hehehe
                End If
            End If
        Else
            MsgBox "Please specify a valid unit price for this product.", vbExclamation, appTitle
        End If
    Else
        MsgBox "Please fill up all fields.", vbExclamation, appTitle
    End If
End Sub


Private Sub EDIT_cmdCANCEL2_Click()
    Me.EDIT_FRAME_confirmation.Visible = False
    Me.EDIT_FRAME_ProdInfo.Visible = False
    Me.EDIT_FRAME_prodList.Visible = True
    Me.edit_prodList.SetFocus
End Sub

Private Sub EDIT_cmdCANCEL3_Click()
    clearEDIT
    Me.edit_confirmationPassword.SetFocus
End Sub










'=================================================================
'                FRAME: REMOVE PRODUCT (DELETE)
'=================================================================
Private Sub DELETE_cmdProceed_Click()
    If CStr(Me.DELETE_confirmationPassword) = CStr(user.password) Then
        Me.DELETE_FRAME_confirmation.Visible = False
        Me.DELETE_FRAME_prodLIST.Visible = True
        Me.DELETE_FRAME_prodLIST.ZOrder vbBringToFront

        retrieveProducts2
        
        If Me.DELETE_prodList.ListItems.Count > 0 Then Me.DELETE_cmdDelete.Enabled = True Else: Me _
        .DELETE_cmdDelete.Enabled = False
        Me.DELETE_prodList.SetFocus
    Else
        MsgBox "Unauthorized Access!", vbExclamation, appTitle
        Me.DELETE_confirmationPassword.SetFocus
    End If
End Sub

Private Sub DELETE_cmdCancel1_Click()
    toggleFrame Me.frameMENU
End Sub

Private Sub DELETE_cmdCancel2_Click()
    clearDELETE
    Me.DELETE_confirmationPassword.SetFocus
End Sub

Private Sub DELETE_cmdDelete_Click()
    Dim selectedProd As String
    
    selectedProd = Trim(Me.DELETE_prodList.SelectedItem.Text)
    
    If MsgBox("Are you sure you want to remove this product from the database? All stocks in the inventory and all transactions connected to this product will also be deleted.", vbQuestion + vbYesNo, appTitle) = vbYes Then
        DB.Execute "DELETE FROM tblProducts WHERE prodID = '" & selectedProd & "'"
        DB.Execute "DELETE FROM tblInventory WHERE prodID = '" & selectedProd & "'"
        DB.Execute "DELETE FROM tblTransactions WHERE prodID = '" & selectedProd & "'"
        MsgBox "Product has been successfully removed.", vbInformation, appTitle
        retrieveProducts2
        If Me.DELETE_prodList.ListItems.Count > 0 Then Me.DELETE_cmdDelete.Enabled = True Else: Me _
        .DELETE_cmdDelete.Enabled = False
        Me.DELETE_prodList.SetFocus
    End If
End Sub










'=================================================================
'                FRAME: ADD STOCKS
'=================================================================
Private Sub ADDSTOCKS_cmdCancel1_Click()
    toggleFrame Me.frameMENU
End Sub

Private Sub ADDSTOCKS_cmdCancel2_Click()
    clearADDSTOCKS
    retrieveProducts3
    Me.ADDSTOCKS_cmdUpdateStock.Enabled = _
    IIf(Me.ADDSTOCKS_prodList.ListItems.Count = 0, False, True)
    Me.ADDSTOCKS_prodList.SetFocus
End Sub

Private Sub ADDSTOCKS_cmdUpdateStock_Click()
    Dim selectedProduct As String
    
    selectedProduct = Me.ADDSTOCKS_prodList.SelectedItem.Text
    retrieveProdInfo2 selectedProduct
    Me.ADDSTOCKS_FRAME_addStock.Visible = True
    Me.ADDSTOCKS_FRAME_prodList.Visible = False
    Me.ADDSTOCKS_FRAME_addStock.ZOrder vbBringToFront
End Sub

Private Sub ADDSTOCKS_cmdAddStock_Click()

    If Me.ADDSTOCKS_deliveryDate.Value > Date Then MsgBox "Delivered in the future? It is not allowed here.", vbExclamation, appTitle: Exit Sub
    If Me.ADDSTOCKS_expiration <= Date Then MsgBox "The product is already expired. The system won't accept it.", vbExclamation, appTitle: Exit Sub
    If IsNumeric(Me.ADDSTOCKS_Quantity) = False Then MsgBox "Quantity is not valid. Please provide a numeric value.", vbExclamation, appTitle: Exit Sub
    addToInventory Me.ADDSTOCKS_prodID, Me.ADDSTOCKS_Quantity, Me.ADDSTOCKS_expiration.Value, Me.ADDSTOCKS_supplier.ListIndex, Me.ADDSTOCKS_supplier.List(Me.ADDSTOCKS_supplier.ListIndex)
    MsgBox "Inventory has been successfully updated!", vbInformation, appTitle
    toggleFrame Me.frameMENU
    clearADDSTOCKS
    retrieveProducts3
    Me.ADDSTOCKS_cmdUpdateStock.Enabled = _
    IIf(Me.ADDSTOCKS_prodList.ListItems.Count = 0, False, True)
    Me.ADDSTOCKS_FRAME_prodList.Visible = True
    Me.ADDSTOCKS_FRAME_addStock.Visible = False
    'Me.ADDSTOCKS_prodList.SetFocus
End Sub





'=================================================================
'                FRAME: VIEW STOCKS
'=================================================================
Private Sub VIEWSTOCKS_cmdClose_Click()
    toggleFrame Me.frameMENU
End Sub

Private Sub VIEWSTOCKS_prodList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    retrieveInventory (Me.VIEWSTOCKS_prodList.SelectedItem.Text)
    Me.VIEWSTOCKS_prodDesc = getProdDesc(Me.VIEWSTOCKS_prodList.SelectedItem.Text)
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
    Me.box(3).Visible = False
    Me.box(4).Visible = False
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

'Generate Product ID
Private Function generateProdID() As String
    RS.Open "SELECT prodID FROM tblProducts", DB, adOpenStatic, adLockOptimistic
        With RS
            If .RecordCount = 0 Then
                generateProdID = Format(1, "00000")
            Else
                .MoveLast
                generateProdID = Format(CInt(!prodID) + 1, "00000")
            End If
        End With
    RS.Close
End Function

'Check if product name is identical to some existing records
Private Function isIdentical(strProductName As String, Optional oldNameToOmit As String) As Boolean
On Error Resume Next
    RS.Open "SELECT prodName from tblProducts WHERE prodName LIKE '" & strProductName & "'", DB, adOpenStatic, adLockOptimistic
        If RS.RecordCount = 0 Then isIdentical = False Else: isIdentical = True
    RS.Close
    
    If strProductName = oldNameToOmit Then isIdentical = False
End Function

'Save new Product Information
Private Function saveProduct()
    RS.Open "SELECT * FROM tblProducts", DB, adOpenStatic, adLockOptimistic
        With RS
            .AddNew
                !prodID = Me.add_prodID
                !prodName = Me.add_prodName
                !prodDesc = Me.add_prodDesc
                !prodPrice = Me.add_prodUnitPrice
                !prodSize = Me.add_prodSize
            .Update
        End With
    RS.Close
End Function

'Update Product Information
Private Function updateProduct(strProductID As String)
    RS.Open "SELECT * FROM tblProducts WHERE prodID = '" & strProductID & "'", DB, adOpenStatic, adLockOptimistic
        With RS
            .Update
                '!prodID = Me.edit_prodID
                !prodName = Me.edit_prodName
                !prodDesc = Me.edit_prodDesc
                !prodPrice = Me.edit_prodUnitPrice
                !prodSize = Me.edit_prodSize
            .Update
        End With
    RS.Close
End Function




'Set frames on their initial state
Private Function initFrames()
    Me.frameMENU.Visible = True
    Me.frameADD.Visible = False
    Me.frameEDIT.Visible = False
    Me.frameDELETE.Visible = False
    Me.frameADDSTOCKS.Visible = False
    Me.frameVIEWSTOCKS.Visible = False
    
    clearADD
    clearEDIT
    clearDELETE
    clearADDSTOCKS
    clearVIEWSTOCKS
End Function

'Show/Hide MAIN FRAMES (ADD,EDIT,DELETE ACCOUNT)
Private Function toggleFrame(showF As Object, Optional stillShowF As Object)
On Error Resume Next
    Dim FRAME_ As PictureBox
    
    Me.frameMENU.Visible = False
    Me.frameADD.Visible = False
    Me.frameEDIT.Visible = False
    Me.frameDELETE.Visible = False
    Me.frameADDSTOCKS.Visible = False
    Me.frameVIEWSTOCKS.Visible = False

    Me.box(0).Visible = False
    Me.box(1).Visible = False
    Me.box(2).Visible = False
    Me.box(3).Visible = False
    Me.box(4).Visible = False
    
    Set FRAME_ = showF
    FRAME_.Visible = True
    FRAME_.ZOrder vbBringToFront
    
    Set FRAME_ = stillShowF
    FRAME_.Visible = True
    FRAME_.ZOrder vbBringToFront
End Function

'Retrieve PRODUCTS and place it on the ListView Control
Private Function retrieveProducts()
    Dim item_ As ListItem
    Dim i As Long
    
    'Clear Headers and Lists
    Me.edit_prodList.ColumnHeaders.Clear
    Me.edit_prodList.ListItems.Clear
    
    'Set Column Headers
    With Me.edit_prodList.ColumnHeaders
        .Add Text:="Product ID", Width:=1100
        .Add Text:="Product Name", Width:=2500
        .Add Text:="Product Size", Width:=1100
        .Add Text:="Unit Price", Width:=1000
    End With
    
    RS.Open "SELECT * FROM tblProducts", DB, adOpenStatic, adLockOptimistic
        With RS
            For i = 1 To .RecordCount Step 1
                Set item_ = Me.edit_prodList.ListItems.Add(Text:=!prodID)
                item_.SubItems(1) = !prodName
                item_.SubItems(2) = !prodSize
                item_.SubItems(3) = "Php " & Format(!prodPrice, "##,##0.00")
                .MoveNext
            Next i
        End With
    RS.Close
End Function
Private Function retrieveProducts2()
    Dim item_ As ListItem
    Dim i As Long
    
    'Clear Headers and Lists
    Me.DELETE_prodList.ColumnHeaders.Clear
    Me.DELETE_prodList.ListItems.Clear
    
    'Set Column Headers
    With Me.DELETE_prodList.ColumnHeaders
        .Add Text:="Product ID", Width:=1100
        .Add Text:="Product Name", Width:=2500
        .Add Text:="Product Size", Width:=1100
        .Add Text:="Unit Price", Width:=1000
    End With
    
    RS.Open "SELECT * FROM tblProducts", DB, adOpenStatic, adLockOptimistic
        With RS
            For i = 1 To .RecordCount Step 1
                Set item_ = Me.DELETE_prodList.ListItems.Add(Text:=!prodID)
                item_.SubItems(1) = !prodName
                item_.SubItems(2) = !prodSize
                item_.SubItems(3) = "Php " & Format(!prodPrice, "##,##0.00")
                .MoveNext
            Next i
        End With
    RS.Close
End Function
Private Function retrieveProducts3()
    Dim item_ As ListItem
    Dim i As Long
    Dim stocks As InventoryStocks
    
    'Clear Headers and Lists
    Me.ADDSTOCKS_prodList.ColumnHeaders.Clear
    Me.ADDSTOCKS_prodList.ListItems.Clear
    
    'Set Column Headers
    With Me.ADDSTOCKS_prodList.ColumnHeaders
        .Add Text:="Product ID", Width:=1100
        .Add Text:="Product Name", Width:=2500
        .Add Text:="Product Size", Width:=1000
        .Add Text:="Current Stock(s)", Width:=3000
        .Add Text:="Unit Price", Width:=1000
        .Add Text:="Total Price", Width:=1000
    End With
    
    RS.Open "SELECT * FROM tblProducts", DB, adOpenStatic, adLockOptimistic
        With RS
            For i = 1 To .RecordCount Step 1
                stocks = retrieveStocks(!prodID)
                Set item_ = Me.ADDSTOCKS_prodList.ListItems.Add(Text:=!prodID)
                item_.SubItems(1) = !prodName
                item_.SubItems(2) = !prodSize
                item_.SubItems(3) = "Stocks: " & stocks.NonExpiredStocks & " / " & "Expired: " & stocks.ExpiredStocks
                item_.SubItems(4) = "Php " & Format(!prodPrice, "##,##0.00")
                item_.SubItems(5) = "Php " & Format(CDbl(CDbl(!prodPrice) * retrieveStocks(!prodID).AllStocks), "##,#0.00")
                .MoveNext
            Next i
        End With
    RS.Close
End Function
Private Function retrieveProducts4()
    Dim item_ As ListItem
    Dim i As Long
    
    'Clear Headers and Lists
    Me.VIEWSTOCKS_prodList.ColumnHeaders.Clear
    Me.VIEWSTOCKS_prodList.ListItems.Clear
    
    'Set Column Headers
    With Me.VIEWSTOCKS_prodList.ColumnHeaders
        .Add Text:="Product ID", Width:=1100
        .Add Text:="Product Name", Width:=2500
        .Add Text:="Product Size", Width:=1100
        .Add Text:="Unit Price", Width:=1000
    End With
    
    RS.Open "SELECT * FROM tblProducts ORDER BY prodID ASC", DB, adOpenStatic, adLockOptimistic
        With RS
            For i = 1 To .RecordCount Step 1
                Set item_ = Me.VIEWSTOCKS_prodList.ListItems.Add(Text:=!prodID)
                item_.SubItems(1) = !prodName
                item_.SubItems(2) = !prodSize
                item_.SubItems(3) = "Php " & Format(!prodPrice, "##,##0.00")
                .MoveNext
            Next i
        End With
    RS.Close
End Function



'Retrieve PRODUCT information (USED IN EDIT FRAME)
Private Function retrieveProdInfo(strProductID As String)
    RS.Open "SELECT * FROM tblProducts WHERE prodID = '" & strProductID & "'", DB, adOpenStatic, adLockOptimistic
        With RS
            Me.edit_prodID = !prodID
            Me.edit_prodDesc = !prodDesc
            Me.edit_prodName = !prodName
            Me.edit_prodNameHolder = !prodName
            Me.edit_prodSize = !prodSize
            Me.edit_prodUnitPrice = !prodPrice
        End With
    RS.Close
End Function

'Retrieve PRODUCT INFORMATION including its current stock (used in ADD STOCKS)
Private Function retrieveProdInfo2(strProdID As String)
    Dim stocks As InventoryStocks
    RS.Open "SELECT * FROM tblProducts WHERE prodID = '" & strProdID & "'", DB, adOpenStatic, adLockOptimistic
        With RS
            stocks = retrieveStocks(!prodID)
            Me.ADDSTOCKS_prodID = !prodID
            Me.ADDSTOCKS_prodName = !prodName
            Me.ADDSTOCKS_prodPrice = !prodPrice
            Me.ADDSTOCKS_currentStock = "Non-Expired Stocks: " & stocks.NonExpiredStocks & vbNewLine & "Expired: " & stocks.ExpiredStocks & vbNewLine & "Total Stocks: " & stocks.AllStocks
        End With
    RS.Close
End Function


'Add stock to inventory
Private Function addToInventory(strProductID As String, intQuantity As Integer, dteExpiration As Date, intSupplier As Integer, strSupplier As String)
    Dim flag1 As Integer
    
    RS.Open "SELECT * FROM tblInventory WHERE prodID = '" & strProductID & "' AND expiration = #" & dteExpiration & "#", DB, adOpenStatic, adLockOptimistic
        If RS.RecordCount = 0 Then flag1 = 0 Else: flag1 = 1
    RS.Close
    
    Select Case flag1
        Case 0
            RS.Open "SELECT * FROM tblInventory", DB, adOpenStatic, adLockOptimistic
                With RS
                    .AddNew
                        !prodID = strProductID
                        !qty = intQuantity
                        !expiration = dteExpiration
                        !supplier = intSupplier
                    .Update
                    recordTransaction strProductID, 0, intQuantity, Me.ADDSTOCKS_deliveryDate.Value, strSupplier, dteExpiration, CInt(!inventoryID)
                End With
                
            RS.Close
            
            
            
        Case 1
            RS.Open "SELECT * FROM tblInventory WHERE prodID = '" & strProductID & "' AND expiration = #" & dteExpiration & "#", DB, adOpenStatic, adLockOptimistic
                With RS
                    .Update
                        '!prodID = strProductID
                        !qty = CInt(!qty + intQuantity)
                        !expiration = dteExpiration
                        !supplier = intSupplier
                    .Update
                    recordTransaction strProductID, 0, intQuantity, Me.ADDSTOCKS_deliveryDate.Value, strSupplier, dteExpiration, CInt(!inventoryID)
                End With
                
            RS.Close
            
    End Select
End Function


'Retrieve Inventory
Private Function retrieveInventory(strProdID As String)
    Dim item_ As ListItem
    Dim i As Long
    Dim stocks As InventoryStocks
    
    'Clear Headers and Lists
    Me.VIEWSTOCKS_inventoryList.ColumnHeaders.Clear
    Me.VIEWSTOCKS_inventoryList.ListItems.Clear
    
    'Set Column Headers
    With Me.VIEWSTOCKS_inventoryList.ColumnHeaders
        .Add Text:="Inventory ID", Width:=1000
        .Add Text:="Product Name", Width:=2000
        .Add Text:="Current Stock(s)", Width:=1000
        .Add Text:="Total Price", Width:=1300
        .Add Text:="Expiration Status", Width:=2500
        .Add Text:="Supplier", Width:=3000
    End With
    
    Me.VIEWSTOCKS_allStocks = 0
    Me.VIEWSTOCKS_nonExpiredStocks = 0
    Me.VIEWSTOCKS_expiredStocks = 0
    
    RS.Open "SELECT * FROM tblInventory WHERE prodID = '" & strProdID & "' AND qty > 0 ORDER BY EXPIRATION ASC", DB, adOpenStatic, adLockOptimistic
        With RS
            For i = 1 To .RecordCount Step 1
                'Retrieve Summary
                stocks = retrieveStocks(!prodID)
                Me.VIEWSTOCKS_allStocks = Format(stocks.AllStocks, "##,##0")
                Me.VIEWSTOCKS_expiredStocks = Format(stocks.ExpiredStocks, "##,##0")
                Me.VIEWSTOCKS_nonExpiredStocks = Format(stocks.NonExpiredStocks, "##,##0")
                
                
                'Retrieve DETAILED stocks
                Set item_ = Me.VIEWSTOCKS_inventoryList.ListItems.Add(Text:=Format(!inventoryID, "00000"))
                item_.SubItems(1) = getProdName(!prodID)
                item_.SubItems(2) = !qty 'retrieveStocks(!prodID, !expiration)
                item_.SubItems(3) = "Php " & Format(CDbl(CDbl(getProdPrice(!prodID)) * !qty), "##,#0.00")
                'item_.SubItems(4) = !expiration
                item_.SubItems(5) = Suppliers(!supplier + 1)
                
                If CDate(!expiration) > CDate(Date) Then
                    item_.SubItems(4) = "Expiration is on " & !expiration
                End If

                If CLng(CDate(!expiration) - Date) <= DEFAULT_EXPIRATION_INTERVAL Then
                    item_.SubItems(4) = "Product will expire in " & CLng(CDate(!expiration) - Date) & " day(s)"
                End If
                
                If CDate(!expiration) <= Date Then
                    item_.SubItems(4) = "Product is expired"
                End If
                
                .MoveNext
            Next i
        End With
    RS.Close
End Function










'---------------------------------------
'  RESET FIELDS
'---------------------------------------
Private Function clearADD()
On Error Resume Next
    Me.add_prodID = Empty
    Me.add_prodName = Empty
    Me.add_prodSize = 1
    Me.add_prodUnitPrice = 0
    Me.add_prodDesc = Empty
End Function
Private Function clearEDIT()
On Error Resume Next
    Me.edit_confirmationPassword = Empty
    Me.edit_prodDesc = Empty
    Me.edit_prodID = Empty
    Me.edit_prodName = Empty
    Me.edit_prodNameHolder = Empty
    Me.edit_prodSize = 1
    Me.edit_prodUnitPrice = 0
    
    Me.edit_prodList.ListItems.Clear
    Me.edit_prodList.ColumnHeaders.Clear
    
    Me.EDIT_FRAME_confirmation.Visible = True
    Me.EDIT_FRAME_ProdInfo.Visible = False
    Me.EDIT_FRAME_prodList.Visible = False
    Me.EDIT_FRAME_confirmation.ZOrder vbBringToFront
    
    Me.EDIT_cmdPropoerties.Enabled = False
End Function
Private Function clearDELETE()
On Error Resume Next
    Me.DELETE_confirmationPassword = Empty
    Me.DELETE_prodList.ListItems.Clear
    Me.DELETE_prodList.ColumnHeaders.Clear
    
    Me.DELETE_FRAME_confirmation.Visible = True
    Me.DELETE_FRAME_prodLIST.Visible = False
    Me.DELETE_FRAME_confirmation.ZOrder vbBringToFront
    Me.DELETE_cmdDelete.Enabled = False
End Function
Private Function clearADDSTOCKS()
On Error Resume Next
    Me.ADDSTOCKS_prodID = Empty
    Me.ADDSTOCKS_prodName = Empty
    Me.ADDSTOCKS_prodPrice = 0
    Me.ADDSTOCKS_Quantity = 1
    Me.ADDSTOCKS_supplier.ListIndex = 0
    Me.ADDSTOCKS_currentStock = 0
    Me.ADDSTOCKS_expiration = Date
    Me.ADDSTOCKS_deliveryDate = Date
    

    Me.ADDSTOCKS_prodList.ListItems.Clear
    Me.ADDSTOCKS_prodList.ColumnHeaders.Clear
    
    Me.ADDSTOCKS_FRAME_addStock.Visible = False
    Me.ADDSTOCKS_FRAME_prodList.Visible = True
    Me.ADDSTOCKS_FRAME_prodList.ZOrder vbBringToFront
    
    Me.ADDSTOCKS_cmdUpdateStock.Enabled = False
End Function
Private Function clearVIEWSTOCKS()
On Error Resume Next
    Me.VIEWSTOCKS_inventoryList.ColumnHeaders.Clear
    Me.VIEWSTOCKS_inventoryList.ListItems.Clear
    Me.VIEWSTOCKS_prodDesc = Empty
    Me.VIEWSTOCKS_prodList.ColumnHeaders.Clear
    Me.VIEWSTOCKS_prodList.ListItems.Clear
    Me.VIEWSTOCKS_allStocks = Empty
    Me.VIEWSTOCKS_expiredStocks = Empty
    Me.VIEWSTOCKS_nonExpiredStocks = Empty
End Function



