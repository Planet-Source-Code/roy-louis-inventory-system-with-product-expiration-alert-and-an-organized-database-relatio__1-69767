VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form reports 
   Caption         =   "Reports"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13575
   Icon            =   "reports.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "reports.frx":1272
   ScaleHeight     =   7935
   ScaleWidth      =   13575
   Begin VB.PictureBox frameVIEWSTOCKS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   120
      Picture         =   "reports.frx":9642
      ScaleHeight     =   6705
      ScaleWidth      =   13305
      TabIndex        =   2
      Top             =   1080
      Width           =   13335
      Begin VB.PictureBox VIEWSTOCKS_FRAME_inventoryList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   120
         Picture         =   "reports.frx":11A12
         ScaleHeight     =   5505
         ScaleWidth      =   12945
         TabIndex        =   5
         Top             =   600
         Width           =   12975
         Begin VB.CommandButton Command1 
            Caption         =   "Refresh Report Table"
            Height          =   375
            Left            =   7440
            TabIndex        =   16
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton transactionType 
            BackColor       =   &H00CDC5B8&
            Caption         =   "Product Income"
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
            Left            =   480
            TabIndex        =   8
            Top             =   960
            Width           =   2655
         End
         Begin VB.OptionButton transactionType 
            BackColor       =   &H00CDC5B8&
            Caption         =   "Product Release"
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
            Left            =   480
            TabIndex        =   7
            Top             =   600
            Value           =   -1  'True
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker transactionFrom 
            Height          =   375
            Left            =   4680
            TabIndex        =   6
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16187393
            CurrentDate     =   39433
         End
         Begin MSComctlLib.ListView transactionTable 
            Height          =   3375
            Left            =   240
            TabIndex        =   9
            Top             =   1920
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   5953
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
         Begin MSComCtl2.DTPicker transactionTo 
            Height          =   375
            Left            =   4680
            TabIndex        =   13
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16187393
            CurrentDate     =   39433
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To:"
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
            Left            =   4080
            TabIndex        =   15
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From:"
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
            Left            =   4080
            TabIndex        =   14
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Choose transaction type:"
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
            TabIndex        =   12
            Top             =   240
            Width           =   2115
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Report Table"
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
            TabIndex        =   11
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Set transaction time table:"
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
            Left            =   3840
            TabIndex        =   10
            Top             =   240
            Width           =   2265
         End
      End
      Begin VB.CommandButton VIEWSTOCKS_cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   11760
         TabIndex        =   3
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Image Image5 
         Height          =   360
         Left            =   120
         Picture         =   "reports.frx":19DE2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stocks Monitoring Report"
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
         TabIndex        =   4
         Top             =   120
         Width           =   3390
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "reports.frx":1A16C
      ScaleHeight     =   705
      ScaleWidth      =   13305
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Reports"
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
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "reports.frx":761EE
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function showTransactions(intTransactionType As Integer, dteFrom As Date, dteTo As Date)
    Dim item_ As ListItem
    Dim i As Long
    
    'Clear Headers and Lists
    Me.transactionTable.ColumnHeaders.Clear
    Me.transactionTable.ListItems.Clear
    
    Select Case intTransactionType
        'PRODUCT IN
        Case 0
        
            'Set Column Headers
            With Me.transactionTable.ColumnHeaders
                .Add Text:="Transaction ID", Width:=1000
                .Add Text:="Transaction Date", Width:=2000
                .Add Text:="Date Delivered", Width:=1500
                .Add Text:="Product ID", Width:=1000
                .Add Text:="Product Name", Width:=2800
                .Add Text:="Supplier", Width:=2800
                .Add Text:="Expiration Date", Width:=1500
                .Add Text:="Quantity Received", Width:=1300
                .Add Text:="Total Amount", Width:=1300
            End With
            
            RS.Open "SELECT * FROM tblTransactions " _
                    & "WHERE " _
                    & "transactionDate >= #" & dteFrom & "# AND " _
                    & "transactionDate <= #" & dteTo & "# AND " _
                    & "transactionType = 0 " _
                    & "ORDER BY transactionDate ASC", DB, adOpenStatic, adLockOptimistic
                    
                With RS
                    For i = 1 To .RecordCount Step 1
                        Set item_ = Me.transactionTable.ListItems.Add(Text:=!transactionID)
                        item_.SubItems(1) = !transactionDate
                        item_.SubItems(2) = !delivery
                        item_.SubItems(3) = !prodID
                        item_.SubItems(4) = getProdName(!prodID)
                        item_.SubItems(5) = !supplier


                        If CDate(!expiration) > CDate(Date) Then
                            item_.SubItems(6) = "Expiration is on " & !expiration
                        End If
        
                        If CLng(CDate(!expiration) - Date) <= DEFAULT_EXPIRATION_INTERVAL Then
                            item_.SubItems(6) = "Product will expire in " & CLng(CDate(!expiration) - Date) & " day(s)"
                        End If
                        
                        If CDate(!expiration) <= Date Then
                            item_.SubItems(6) = "Product is expired"
                        End If

                        item_.SubItems(7) = !qty
                        item_.SubItems(8) = "Php " & Format(!prodTotalPrice, "##,##0.00")
                        .MoveNext
                    Next i
                End With
            RS.Close
        
        'PRODUCT OUT
        Case 1
            'Set Column Headers
            With Me.transactionTable.ColumnHeaders
                .Add Text:="Transaction ID", Width:=1000
                .Add Text:="Transaction Date", Width:=2000
                .Add Text:="Date Sold", Width:=1500
                .Add Text:="Product ID", Width:=1000
                .Add Text:="Product Name", Width:=2800
                .Add Text:="Supplier", Width:=2800
                .Add Text:="Expiration Date", Width:=1500
                .Add Text:="Quantity Deducted", Width:=1300
                .Add Text:="Total Amount", Width:=1300
            End With
            
            RS.Open "SELECT * FROM tblTransactions " _
                    & "WHERE " _
                    & "transactionDate >= #" & dteFrom & "# AND " _
                    & "transactionDate <= #" & dteTo & "# AND " _
                    & "transactionType = 1 " _
                    & "ORDER BY transactionDate ASC", DB, adOpenStatic, adLockOptimistic
                    
                With RS
                    For i = 1 To .RecordCount Step 1
                        Set item_ = Me.transactionTable.ListItems.Add(Text:=!transactionID)
                        item_.SubItems(1) = !transactionDate
                        item_.SubItems(2) = !delivery
                        item_.SubItems(3) = !prodID
                        item_.SubItems(4) = getProdName(!prodID)
                        item_.SubItems(5) = !supplier


                        If CDate(!expiration) > CDate(Date) Then
                            item_.SubItems(6) = "Expiration is on " & !expiration
                        End If
        
                        If CLng(CDate(!expiration) - Date) <= DEFAULT_EXPIRATION_INTERVAL Then
                            item_.SubItems(6) = "Product will expire in " & CLng(CDate(!expiration) - Date) & " day(s)"
                        End If
                        
                        If CDate(!expiration) <= Date Then
                            item_.SubItems(6) = "Product is expired"
                        End If

                        item_.SubItems(7) = !qty
                        item_.SubItems(8) = "Php " & Format(!prodTotalPrice, "##,##0.00")
                        .MoveNext
                    Next i
                End With
            RS.Close
    End Select
End Function

Private Sub Command1_Click()
    showTransactions IIf(Me.transactionType(0).Value = True, 1, 0), Me.transactionFrom, Me.transactionTo
End Sub

Private Sub Form_Load()
    Me.Width = 13695
    Me.Height = 8340
End Sub

Private Sub VIEWSTOCKS_cmdClose_Click()
    Unload Me
End Sub
