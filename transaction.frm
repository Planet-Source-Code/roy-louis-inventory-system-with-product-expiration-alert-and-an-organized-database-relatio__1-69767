VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form transaction 
   Caption         =   "Transaction"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13590
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "transaction.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   13590
   Begin VB.PictureBox frameVIEWSTOCKS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      Picture         =   "transaction.frx":83D0
      ScaleHeight     =   6345
      ScaleWidth      =   13305
      TabIndex        =   2
      Top             =   960
      Width           =   13335
      Begin VB.PictureBox VIEWSTOCKS_FRAME_inventoryList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5055
         Left            =   120
         Picture         =   "transaction.frx":107A0
         ScaleHeight     =   5025
         ScaleWidth      =   12945
         TabIndex        =   4
         Top             =   600
         Width           =   12975
         Begin VB.ComboBox prodSupplier 
            Height          =   315
            Left            =   10080
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton cmdDischarge 
            Caption         =   "Discharge from the Inventory"
            Height          =   375
            Left            =   10320
            TabIndex        =   30
            Top             =   4440
            Width           =   2415
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   495
            Left            =   9000
            TabIndex        =   29
            Top             =   3600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.TextBox prodPriceMIRROR 
            Height          =   315
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   2760
            Width           =   2175
         End
         Begin VB.TextBox prodPrice 
            Height          =   315
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   3000
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox prodQty 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   6840
            MaxLength       =   4
            TabIndex        =   24
            Text            =   "1"
            Top             =   3600
            Width           =   2175
         End
         Begin VB.TextBox prodName 
            Height          =   315
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   2760
            Width           =   2175
         End
         Begin VB.TextBox prodID 
            Height          =   315
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   2400
            Width           =   2175
         End
         Begin VB.TextBox VIEWSTOCKS_prodDesc 
            Height          =   1095
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   3600
            Width           =   4815
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
            TabIndex        =   7
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
            TabIndex        =   6
            Top             =   1320
            Width           =   1695
         End
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
            TabIndex        =   5
            Top             =   1320
            Width           =   1695
         End
         Begin MSComctlLib.ListView VIEWSTOCKS_prodList 
            Height          =   2295
            Left            =   240
            TabIndex        =   9
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Price:"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   9120
            TabIndex        =   27
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier:"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   9120
            TabIndex        =   25
            Top             =   2400
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product ID No.:"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   5520
            TabIndex        =   21
            Top             =   2400
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity:"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   5520
            TabIndex        =   20
            Top             =   3600
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name:"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   5520
            TabIndex        =   19
            Top             =   2760
            Width           =   1065
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
            TabIndex        =   17
            Top             =   240
            Width           =   5880
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Information"
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
            TabIndex        =   16
            Top             =   2040
            Width           =   1725
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
            TabIndex        =   15
            Top             =   720
            Width           =   1095
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
            TabIndex        =   14
            Top             =   3360
            Width           =   1665
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inventory Stocks"
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
            TabIndex        =   13
            Top             =   720
            Width           =   1470
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expired Stocks:"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   9600
            TabIndex        =   12
            Top             =   1080
            Width           =   1110
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Non-Expired Stocks:"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   7560
            TabIndex        =   11
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Stocks:"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   5520
            TabIndex        =   10
            Top             =   1080
            Width           =   930
         End
      End
      Begin VB.CommandButton VIEWSTOCKS_cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   11760
         TabIndex        =   3
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Discharge"
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
         TabIndex        =   18
         Top             =   120
         Width           =   2670
      End
      Begin VB.Image Image5 
         Height          =   360
         Left            =   120
         Picture         =   "transaction.frx":18B70
         Stretch         =   -1  'True
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "transaction.frx":18EFA
      ScaleHeight     =   705
      ScaleWidth      =   13305
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction"
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
         Width           =   1710
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "transaction.frx":74F7C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Retrieve Inventory
Private Function retrieveInventory(strProdID As String)
    Dim item_ As ListItem
    Dim i As Long
    Dim stocks As InventoryStocks
    
    Me.VIEWSTOCKS_allStocks = 0
    Me.VIEWSTOCKS_expiredStocks = 0
    Me.VIEWSTOCKS_nonExpiredStocks = 0
    Me.cmdDischarge.Enabled = False
    Me.prodQty.Enabled = False
    
    Me.prodPrice = 0#
    Me.prodID = Empty
    Me.prodName = Empty
    Me.prodSupplier.ListIndex = 5
    Me.VIEWSTOCKS_prodDesc = Empty
    
    RS.Open "SELECT * FROM tblInventory WHERE prodID = '" & strProdID & "' AND qty > 0", DB, adOpenStatic, adLockOptimistic
        If RS.RecordCount > 0 Then
            With RS
                For i = 1 To .RecordCount Step 1
                    'Retrieve Summary
                    stocks = retrieveStocks(!prodID)
                    Me.VIEWSTOCKS_allStocks = Format(stocks.AllStocks, "##,##0")
                    Me.VIEWSTOCKS_expiredStocks = Format(stocks.ExpiredStocks, "##,##0")
                    Me.VIEWSTOCKS_nonExpiredStocks = Format(stocks.NonExpiredStocks, "##,##0")
                    
                    Me.prodPrice = getProdPrice(!prodID)
                    Me.prodID = !prodID
                    Me.prodName = getProdName(!prodID)
                    Me.prodSupplier.ListIndex = !supplier
                    Me.VIEWSTOCKS_prodDesc = getProdDesc(!prodID)
                    
                    .MoveNext
                Next i
            End With
            Me.cmdDischarge.Enabled = True
            Me.prodQty.Enabled = True
        End If
    RS.Close
End Function


Private Function retrieveProducts()
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

Private Sub cmdDischarge_Click()
    deductFromInventory Me.prodID, Me.prodQty
End Sub


Private Sub Form_Load()
    Me.Height = 7950
    Me.Width = 13710
    
    Me.Show
    
    
    loadSuppliers Me.prodSupplier
    Me.prodSupplier.AddItem ""
    retrieveProducts
    Me.VIEWSTOCKS_prodList.SetFocus
    Me.cmdDischarge.Enabled = False
    
    Me.prodPrice = 0#
    Me.prodID = Empty
    Me.prodName = Empty
    Me.prodSupplier.ListIndex = 5
    Me.VIEWSTOCKS_prodDesc = Empty
    Me.prodQty = 1
    Me.prodQty.Enabled = False
    
    Me.VIEWSTOCKS_allStocks = 0
    Me.VIEWSTOCKS_expiredStocks = 0
    Me.VIEWSTOCKS_nonExpiredStocks = 0
End Sub



Private Sub prodPrice_Change()
    Me.prodPriceMIRROR = "Php " & Format(Me.prodPrice, "##,##0.00")
End Sub









Private Sub prodQty_Change()
    If Me.prodQty = Empty Then Me.prodQty = 1
End Sub

Private Sub prodQty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 58       'Numerci [0-9]
        Case 8              'Backspace
        Case 13             'Enter Key
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub UpDown1_DownClick()
    If Me.prodQty > 1 Then
        Me.prodQty = Int(Me.prodQty) - 1
    Else
        Me.prodQty = 1
    End If
End Sub

Private Sub UpDown1_UpClick()
    Me.prodQty = Int(Me.prodQty) + 1
End Sub

Private Sub VIEWSTOCKS_cmdClose_Click()
    Unload Me
End Sub

Private Sub VIEWSTOCKS_prodList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    retrieveInventory (Me.VIEWSTOCKS_prodList.SelectedItem.Text)
End Sub











Private Function deductFromInventory(strProdID As String, lngQty As Long)
    If IsNumeric(Me.prodQty) = False Then MsgBox "Quantity is invalid. Provide numeric value.", vbExclamation, appTitle: Exit Function
    Dim RS2 As ADODB.Recordset
    Dim RS3 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    Set RS3 = New ADODB.Recordset
    Dim i As Long
    Dim stocks As InventoryStocks
    Dim deduction As Long
    Dim temp As Long
    Dim SQL As String
    
    stocks = retrieveStocks(strProdID)

    
    If stocks.NonExpiredStocks >= lngQty Then
    
        'Non Expired Products and Stocks > 0
        RS2.Open "SELECT * FROM tblInventory WHERE prodID = '" & strProdID & "' AND qty > 0 AND expiration > #" & Date & "# ORDER BY expiration ASC", DB, adOpenStatic, adLockOptimistic
            With RS2
                deduction = lngQty

                Do Until (deduction = 0)
                    If !qty >= deduction Then
                        SQL = "UPDATE tblInventory SET qty = " & CLng(CLng(!qty) - deduction) & " WHERE inventoryID = " & !inventoryID & ""
                        recordTransaction Me.prodID, 1, CInt(deduction), Date, Me.prodSupplier.List(Me.prodSupplier.ListIndex), !expiration, CInt(!inventoryID)
                        deduction = 0
                        DB.Execute SQL
                    Else
                        SQL = "UPDATE tblInventory SET qty = 0 WHERE inventoryID = " & !inventoryID & ""
                        recordTransaction Me.prodID, 1, CInt(!qty), Date, Me.prodSupplier.List(Me.prodSupplier.ListIndex), !expiration, CInt(!inventoryID)
                        deduction = deduction - !qty
                        DB.Execute SQL
                    End If
                    .MoveNext
                Loop
            End With
        RS2.Close
        
        MsgBox "Transaction Successful!", vbInformation, appTitle
        
        Call Form_Load
    Else
        MsgBox "Insufficient stocks!", vbInformation, appTitle
    End If
End Function
