Attribute VB_Name = "Module1"
'===========================================================
'
'                    GLOBAL DECLARATIONS
'
'===========================================================

Option Explicit
Option Base 1

'Types -----------------------------------------------------
Public Type UserInfo
    lastname As String         'Last Name
    firstname As String        'First Name
    username As String         'User Name
    password As String         'Password (for faster access)
    level As Integer           '[0]-Limited, [1]-Admin
    secretQ As Integer         'ListIndex of Secret Question
    secretA As String          'Secret Answer
End Type

Public Type InventoryStocks
    AllStocks As Long          'ALL stocks in the inventory, including expired prod.
    ExpiredStocks As Long      'ALL expired products
    NonExpiredStocks As Long   'ALL non-expired products (haha, tma ba?)
    ReservedVariable As Long   'Holds special values
End Type
'-----------------------------------------------------------

'API
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'Database Connection and Recordset
Public DB As ADODB.Connection         'Var. for database connection
Public RS As ADODB.Recordset          'Var. for recordsets (accessing tables)

'Software Name
Public appTitle As String             'Var. that contains the app. name

'Secret Questions (Hard-Coded)
Public Questions(5) As String         'An array containing all Secret Questions

'Suppliers (Hard-Coded)
Public Suppliers(5) As String         'An array containing all Suppliers

'Current User Logged On
Public user As UserInfo               'Holds data of the currently logged on user

'Default Account
Public Const DEFAULT_USERNAME As String = "Admin"
Public Const DEFAULT_PASSWORD As String = "1234"
Public Const DEFAULT_FIRSTNAME As String = "Roy"
Public Const DEFAULT_LASTNAME As String = "Louis"
Public Const DEFAULT_LEVEL As String = 1

'Expiration Alert (Number of days before showing an expiration alert)
Public Const DEFAULT_EXPIRATION_INTERVAL As Long = 7



'===========================================================
'                      MAIN FUNCTION
'                    (loads at startup)
'===========================================================
Sub Main()
    Load splash
    splash.Show
    splash.ZOrder vbBringToFront
    '-------------------------------------------------------
    ' Connect and Open Database File
    '-------------------------------------------------------
    
        'Place the filename of DB here --|
        '                                |
        '                                v
        Const DB_FileName As String = "mainDB.mdb"
        Const DB_Password As String = "1234"
        
        Set DB = CreateObject("ADODB.Connection")                             'Create ActiveX Object
        Set RS = New ADODB.Recordset                                          'Set New RS
        DB.Provider = "Microsoft Jet 4.0 OLE DB Provider"                     'Provider
        DB.ConnectionString = "Data Source=" & App.Path & "\" & DB_FileName   'Data Source (Database Filename)
        DB.Properties("Jet OLEDB:Database Password") = DB_Password            'Password
        DB.Open                                                               'Open Database File
    
    
    
    '-------------------------------------------------------
    ' Set Application Title
    '-------------------------------------------------------
    
        'Provide the Software Name you want --|
        '                                     |
        '                     |---------------|
        '                     |
        '                     v
        Const softwareName = "•RBN Trading•"
        
        appTitle = softwareName _
                 & " Version " & App.Major & "." & App.Minor
                 
                 
    '-------------------------------------------------------
    ' Set Secret Questions
    '-------------------------------------------------------
    
        Questions(1) = "Who is your spouse?"
        Questions(2) = "What is your favorite pet?"
        Questions(3) = "What is your pet's name?"
        Questions(4) = "What is your favorite color?"
        Questions(5) = "Who is your first crush?"
        
        
        
    '-------------------------------------------------------
    ' Set Suppliers
    '-------------------------------------------------------
    
        Suppliers(1) = "Coca Cola"
        Suppliers(2) = "Uniliver"
        Suppliers(3) = "Procter and Gamble"
        Suppliers(4) = "Nissin"
        Suppliers(5) = "Colgate"
        
        
    Sleep 1500
    Unload splash
    
    '-------------------------------------------------------
    ' Load Splash Screen
    '-------------------------------------------------------
    
        'Load splash
        'splash.Show
        
        'Load userAcct
        'userAcct.Show
        
        Load login
        login.Show
        

End Sub



'===========================================================
'                 GLOBALLY USED FUNCTIONS
'                  (used on most modules)
'===========================================================

'Retrieve Current Stock
Public Function retrieveStocks(strProductID As String, Optional dteExpiration As Date) As InventoryStocks
    Dim RS2 As New ADODB.Recordset
    If dteExpiration <> Empty Then
        RS2.Open "SELECT qty FROM tblInventory WHERE prodID ='" & strProductID & "' AND expiration = #" & dteExpiration & "#", DB, adOpenKeyset, adLockOptimistic
    Else
        RS2.Open "SELECT qty FROM tblInventory WHERE prodID ='" & strProductID & "'", DB, adOpenKeyset, adLockOptimistic
    End If
        With RS2
            Dim i As Long
            Dim totalStocks As Long
            
            For i = 1 To .RecordCount Step 1
                totalStocks = totalStocks + !qty
                .MoveNext
            Next i
            
            retrieveStocks.ReservedVariable = totalStocks
        End With
    RS2.Close
    
    RS2.Open "SELECT qty FROM tblInventory WHERE prodID = '" & strProductID & "' AND expiration <= #" & Date & "#", DB, adOpenStatic, adLockOptimistic
        With RS2
            Dim i2 As Long
            Dim totalStocks2 As Long
            
            For i2 = 1 To .RecordCount Step 1
                totalStocks2 = totalStocks2 + !qty
                .MoveNext
            Next i2
            retrieveStocks.ExpiredStocks = totalStocks2
        End With
    RS2.Close
    
    RS2.Open "SELECT qty FROM tblInventory WHERE prodID = '" & strProductID & "' AND expiration > #" & Date & "#", DB, adOpenStatic, adLockOptimistic
        With RS2
            Dim i3 As Long
            Dim totalStocks3 As Long
            
            For i3 = 1 To .RecordCount Step 1
                totalStocks3 = totalStocks3 + !qty
                .MoveNext
            Next i3
            retrieveStocks.NonExpiredStocks = totalStocks3
        End With
    RS2.Close
    
    RS2.Open "SELECT qty FROM tblInventory WHERE prodID = '" & strProductID & "'", DB, adOpenStatic, adLockOptimistic
        With RS2
            Dim i4 As Long
            Dim totalStocks4 As Long
            
            For i4 = 1 To .RecordCount Step 1
                totalStocks4 = totalStocks4 + !qty
                .MoveNext
            Next i4
            retrieveStocks.AllStocks = totalStocks4
        End With
    RS2.Close
End Function


'Update Transaction Table
Public Function recordTransaction(strProdID As String, intTransactionType As Integer, intQuantity As Integer, dteDelivery As Date, strSupplier As String, dteExpiration As Date, intInventoryID As Integer)
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    RS2.Open "SELECT * FROM tblTransactions", DB, adOpenStatic, adLockOptimistic
        With RS2
            .AddNew
                !prodID = strProdID
                !inventoryID = intInventoryID
                !transactionDate = DateTime.Now
                !transactionType = intTransactionType
                !qty = intQuantity
                !prodTotalPrice = CDbl(getProdPrice(strProdID) * intQuantity)
                !delivery = dteDelivery
                !supplier = strSupplier
                !expiration = dteExpiration
            .Update
        End With
    RS2.Close
End Function

'Get Item Price
Public Function getProdPrice(strProdID As String) As Double
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    RS2.Open "SELECT prodPrice FROM tblProducts WHERE prodID = '" & strProdID & "'", DB, adOpenStatic, adLockOptimistic
        With RS2
        If .RecordCount > 0 Then getProdPrice = !prodPrice
        End With
    RS2.Close
End Function

'Get Item Name
Public Function getProdName(strProdID As String) As String
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    RS2.Open "SELECT prodName FROM tblProducts WHERE prodID = '" & strProdID & "'", DB, adOpenStatic, adLockOptimistic
        With RS2
        If .RecordCount > 0 Then getProdName = !prodName
        End With
    RS2.Close
End Function

'Get Item Description
Public Function getProdDesc(strProdID As String) As String
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    RS2.Open "SELECT prodDesc FROM tblProducts WHERE prodID = '" & strProdID & "'", DB, adOpenStatic, adLockOptimistic
        With RS2
        If .RecordCount > 0 Then getProdDesc = !prodDesc
        End With
    RS2.Close
End Function

'Retrieve User Account Information
Public Function updateUserInfoVariable(strUsername As String) 'As UserInfo
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    RS2.Open "SELECT * FROM users WHERE username = '" & strUsername & "'", DB, adOpenStatic, adLockOptimistic
        With RS2
            user.firstname = !firstname
            user.lastname = !lastname
            user.level = !accountType
            user.password = !password
            user.secretA = !secretA
            user.secretQ = !secretQ
            user.username = !username
        End With
    RS2.Close
End Function


'Load Suppliers
Public Function loadSuppliers(onComboBox As Object)
On Error Resume Next
    Dim container As ComboBox
    Dim i As Integer
    
    Set container = onComboBox
    
    container.Clear
    For i = 1 To 5 Step 1
        container.AddItem Suppliers(i)
    Next i
    container.ListIndex = 0
End Function
