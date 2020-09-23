VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   735
   ClientWidth     =   11325
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu menu0 
      Caption         =   "&Start"
      Begin VB.Menu menu1 
         Caption         =   "&User Accounts"
         Shortcut        =   {F1}
      End
      Begin VB.Menu menuY 
         Caption         =   "-"
      End
      Begin VB.Menu menu2 
         Caption         =   "&Product and Inventory Management"
         Shortcut        =   {F2}
      End
      Begin VB.Menu menu3 
         Caption         =   "&Transactions"
         Shortcut        =   {F3}
      End
      Begin VB.Menu menu4 
         Caption         =   "&Reports"
         Shortcut        =   {F4}
      End
      Begin VB.Menu menuX 
         Caption         =   "-"
      End
      Begin VB.Menu menu5 
         Caption         =   "&Log Off"
         Shortcut        =   ^L
      End
      Begin VB.Menu menu6 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Me.Caption = appTitle
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub menu1_Click()
    Unload prodInfo
    Unload reports
    Unload transaction
    Load userAcct
    userAcct.Left = 0
    userAcct.Top = 0
    userAcct.Show
End Sub

Private Sub menu2_Click()
    Load prodInfo
    Unload reports
    Unload transaction
    Unload userAcct
    prodInfo.Left = 0
    prodInfo.Top = 0
    prodInfo.Show
End Sub

Private Sub menu3_Click()
    Unload prodInfo
    Unload reports
    Load transaction
    Unload userAcct
    transaction.Left = 0
    transaction.Top = 0
    transaction.Show
End Sub

Private Sub menu4_Click()
    Unload prodInfo
    Load reports
    Unload transaction
    Unload userAcct
    reports.Left = 0
    reports.Top = 0
    reports.Show
End Sub

Private Sub menu5_Click()
    Unload Me
    Load login
End Sub

Private Sub menu6_Click()
    End
End Sub
