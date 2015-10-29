Public Const versionNumber = 0.4
Public Const versionType = "Custom"
Public Const versionSite = "Dronfield"

Const SettingsSheetName = "Control Buttons"

Public Conn As ADODB.Connection
Public boolConnectionOpen As Boolean
Dim boolRecordsetOpen As Boolean
Dim databaseLocation As String

'Temporarily changes the selected database for useage.
Public Sub changeTargetDatabase(newDatabaseLocation As String)
    databaseLocation = newDatabaseLocation
End Sub

'Returns the currently loaded database location.
Public Function returnLoadedDatabaseLocation() As String
    returnLoadedDatabaseLocation = databaseLocation
End Function

'Loads the default database settings from the worksheet.
Public Function Load_Settings() As Long
    Dim wbCurrent As Workbook
    Dim wsSettings As Worksheet
    
    Set wbCurrent = ThisWorkbook
    Set wsSettings = wbCurrent.Sheets(SettingsSheetName)
    databaseLocation = wsSettings.Range("H2").Value
    Set wsSettings = Nothing
    Set wbCurrent = Nothing
End Function


Private Sub saveNewDatabaseLocation(databaseLocation As String)
    Dim wbCurrent As Workbook
    Dim wsSettings As Worksheet
    
    Set wbCurrent = ThisWorkbook
    Set wsSettings = wbCurrent.Sheets(SettingsSheetName)
    wsSettings.Range("H2").Value = databaseLocation
    Set wsSettings = Nothing
    Set wbCurrent = Nothing
End Sub

'function to execute raw SQL commands
Public Sub rawSQLExecute(stSQL)
    Conn.Execute stSQL
End Sub

'Function to return a record count from a raw SQL string
Public Function getRawRecordCount(stSQL) As Long
    Dim lRecordCount As Long
        rsTemporary As New ADODB.Recordset
        rsTemporary.CursorType = adOpenKeyset
        rsTemporary.Open stSQL, Conn
        lRecordCount = rsTemporary.RecordCount
        rsTemporary.Close
        getRawRecordCount = lRecordCount
        Set rsTemporary = Nothing
End Function

Private Function getNewDatabaseLocation() As String
    getNewDatabaseLocation = Application.GetOpenFilename("*.mdb,*.mdb", , "Select Database Location")
    Call saveNewDatabaseLocation(getNewDatabaseLocation)
End Function

Private Function checkDatabaseLocation(databaseLocation As String) As String
    Dim boolTest As Boolean
    
    If databaseLocation = "" Or IsNull(databaseLocation) Then
        databaseLocation = getNewDatabaseLocation()
    End If
    
    boolTest = Checks.FileFolderExists(databaseLocation)
    
    If boolTest = False Then
        databaseLocation = getNewDatabaseLocation()
    End If
    
    checkDatabaseLocation = databaseLocation
End Function

'Gather the information to connect to the appropriate database.
Public Function Open_Connection(Optional customDatabaseLocation As String = vbNullString) As Long
    Set Conn = New ADODB.Connection
    
    If Not customDatabaseLocation = vbNullString Then databaseLocation = customDatabaseLocation 'Use the custom location if it is supplied
    If databaseLocation = "" Or IsNull(databaseLocation) Then Load_Settings 'Use the default if not supplied.
    
    databaseLocation = checkDatabaseLocation(databaseLocation) 'Check to see if the database exists.
    
    Conn.Provider = "Microsoft.Jet.OLEDB.4.0"    'Main Connection method used
    Conn.ConnectionString = "Data Source=" & databaseLocation 'DB Location to open
    
    Conn.Open   'Open the Connection to the database.

    Debug.Print Err.Number
    Open_Connection = Err.Number
    If Err.Number = 0 Then
        boolConnectionOpen = True
    Else
        boolConnectionOpen = False
    End If
End Function

Public Function UpdateProductList(singlesBarcode As String, productDescription As String, packSize As Long, productSupplier As String, lastOrdered As Date) As Long
    Dim rsProductList As New ADODB.Recordset
    Dim rsCount As Long
    Dim stSQL As String
    Dim sqlProductSupplier As String
    rsProductList.CursorType = adOpenKeyset
    
    stSQL = "SELECT [Singles Barcode] " & _
            "FROM [Products] " & _
            "WHERE [Singles Barcode] = '" & singlesBarcode & "';"
            
    rsProductList.Open stSQL, Conn
    boolRecordsetOpen = True
    rsCount = rsProductList.RecordCount
    
    productDescription = Replace(productDescription, "'", "''")
    sqlProductSupplier = Replace(productSupplier, "'", "''")
    
    If rsCount > 0 Then

    ElseIf rsCount = 0 Then
        stSQL = "INSERT INTO [Products] ([Singles Barcode], [Product Description], [Pk Sz], [Supplier to Exel], [Last Ordered]) " & _
                "VALUES ('" & singlesBarcode & "', '" & productDescription & "', " & packSize & ", '" & sqlProductSupplier & "', #" & Format(lastOrdered, "mm/dd/yyyy") & "#);"
    End If
    
    rsProductList.Close
    boolRecordsetOpen = False
    Conn.Execute stSQL
    
    If boolRecordsetOpen = True Then
        rsProductList.Close
        boolRecordsetOpen = False
    End If
    
    UpdateProductList = 1
End Function

Public Function fetchPickRate(productSinglesCode As String) As Long
    Dim rs As New ADODB.Recordset
    Dim stSQL As String
    Dim rsCount As Long
    Dim resultant As Variant
    
    rs.CursorType = adOpenKeyset
    
    stSQL = "SELECT [Singles Barcode], [Pick Rate] FROM [Products] WHERE [Singles Barcode] = '" & productSinglesCode & "';"
    
    rs.Open stSQL, Conn
    rsCount = rs.RecordCount
    
    If rsCount > 0 Then
        If IsEmpty(rs.Fields(1)) Then
            resultant = 0
        Else
            resultant = rs.Fields(1).Value
        End If
    End If
    
    rs.Close
    fetchPickRate = resultant
End Function

Public Function FetchString(tableName As String, searchField As String, resultField As String, searchString As String) As String
    Dim rs As New ADODB.Recordset
    Dim resultString As String
    Dim rsCount As Long
    
    rs.CursorType = adOpenKeyset
    
    stSQL = "SELECT [" & searchField & "], [" & resultField & "] FROM [" & tableName & "] " & _
            "WHERE [" & searchField & "] = '" & searchString & "';"

    rs.Open stSQL, Conn
    rsCount = rs.RecordCount
        
    If rsCount > 0 Then
        If Not IsNull(rs.Fields(1).Value) Or rs.Fields(1).Value = "" Then
            resultString = rs.Fields(1).Value
        Else
            resultString = ""
        End If
    End If
    
    rs.Close
    FetchString = resultString
    Set rs = Nothing
End Function

Public Function Close_Connection() As Long
    Conn.Close
    boolConnectionOpen = False
    Set Conn = Nothing
End Function
