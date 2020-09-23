VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form expirationAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Expiration Alert"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "expirationAlert.frx":0000
   ScaleHeight     =   3375
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Reminding"
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
      TabIndex        =   6
      ToolTipText     =   "Stop showing this alert."
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   5760
      Top             =   240
   End
   Begin VB.CommandButton cmdRemindLater 
      Caption         =   "Remind Me Later..."
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
      TabIndex        =   3
      ToolTipText     =   "This action wil let the alert shows again in the future."
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdPurgeExpired 
      Caption         =   "Remove All Expired Products"
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
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "This action will delete all expired products from the database."
      Top             =   2880
      Width           =   2415
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "expirationAlert.frx":E44F2
      ScaleHeight     =   705
      ScaleWidth      =   6105
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "expirationAlert.frx":E6898
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration Alert"
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
         TabIndex        =   4
         Top             =   120
         Width           =   2310
      End
   End
   Begin MSComctlLib.ListView expiredList 
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2355
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
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The following products will expire in few days:"
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
      TabIndex        =   5
      Top             =   960
      Width           =   3840
   End
End
Attribute VB_Name = "expirationAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPurgeExpired_Click()
    DB.Execute "DELETE FROM tblInventory WHERE expiration <= #" & Date & "#"
    MsgBox "All expired stocks have been successfully removed from the database.", vbInformation, appTitle
    Call Timer1_Timer
End Sub

Private Sub cmdRemindLater_Click()
    Me.Timer1.Enabled = True
    Me.Hide
End Sub

Private Sub cmdStop_Click()
    Me.Hide
    Me.Timer1.Enabled = False
End Sub

Public Sub Timer1_Timer()
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    Dim item_ As ListItem
    Dim i As Long
    Dim flag1 As Integer
    
    'Clear Headers and Lists
    Me.expiredList.ColumnHeaders.Clear
    Me.expiredList.ListItems.Clear
    
    'Set Column Headers
    With Me.expiredList.ColumnHeaders
        .Add Text:="Product ID", Width:=1100
        .Add Text:="Product Name", Width:=2500
        .Add Text:="Expiration", Width:=3000
    End With
    
    flag1 = 0
    RS2.Open "SELECT * FROM tblInventory WHERE expiration <= #" & CDate(Date + DEFAULT_EXPIRATION_INTERVAL) & "#", DB, adOpenStatic, adLockOptimistic
        If RS2.RecordCount > 0 Then flag1 = 1
    RS2.Close
    
    If flag1 = 1 Then
        RS2.Open "SELECT * FROM tblInventory WHERE expiration <= #" & CDate(Date + DEFAULT_EXPIRATION_INTERVAL) & "# ORDER BY expiration ASC", DB, adOpenStatic, adLockOptimistic
            With RS2
                For i = 1 To .RecordCount Step 1
                    Set item_ = Me.expiredList.ListItems.Add(Text:=!prodID)
                    item_.SubItems(1) = getProdName(!prodID)
                    
                    If CDate(!expiration) > CDate(Date) Then
                        item_.SubItems(2) = "Expiration is on " & !expiration
                    End If
    
                    If CLng(CDate(!expiration) - Date) <= DEFAULT_EXPIRATION_INTERVAL Then
                        item_.SubItems(2) = "Product will expire in " & CLng(CDate(!expiration) - Date) & " day(s)"
                    End If
                    
                    If CDate(!expiration) <= Date Then
                        item_.SubItems(2) = "Product is expired"
                    End If
                    .MoveNext
                Next i
            End With
        RS2.Close
        Me.Left = (Screen.Width - Me.Width) - 100
        Me.Top = (Screen.Height - Me.Height) - 450
        If user.level = 1 Then
            Me.cmdPurgeExpired.Enabled = True
        Else
            Me.cmdPurgeExpired.Enabled = False
        End If
        
        Me.Show
        Me.ZOrder vbBringToFront
        Me.Timer1.Enabled = False
        Me.expiredList.SetFocus
    End If
End Sub
