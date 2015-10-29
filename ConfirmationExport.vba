Private Const helpFile = "U:\Boots_Contract\Chill\Dronfield\Dude's Folder\Information\Thermo BAC55.pdf"
Private Const stDronfieldSystem = "\\prg-dc.dhl.com\uk\dsc\sites\Boots\Boots_Contract\Chill\Dronfield\Dude's Folder\Maintenance\Dronfield-System.mdb"
Private wsConfirmation As Worksheet
Private arSheetNumber() As String

Sub loadHelpFile()
    Dim WSHShell        As Object
    Dim sAcrobatPath    As String
    
    Set WSHShell = CreateObject("Wscript.Shell")
    sAcrobatPath = WSHShell.regread("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe\")
    
    MsgBox sAcrobatPath
    
    'ActiveWorkbook.FollowHyperlink helpFile
    Documents.Open "U:\Boots_Contract\Chill\Dronfield\Dude's Folder\Information\Help Files\Shift Plan.pdf"
End Sub

Private Sub displayError(strError As String)
    Debug.Print strError
    End
End Sub

'Main sub for the entire export process
Public Sub main()
    'Load values to be used
    initialiseValues
    clearPreviousData
    addToDatabase collectNewData   'Enter the data to the database.
    cleanUp                        'Clear up the values
End Sub

'Set up the main values to be used
Private Function initialiseValues() As Long
    Set wsConfirmation = ThisWorkbook.Sheets("Pick Confirmation")
    Erase arSheetNumber
    initialiseValues = 0
End Function

'Add to the preArray for database useage.
Private Sub addToDatabaseArray(stEntry As String)
    If IsVarArrayEmpty(arSheetNumber) Then 'Determine if the array is empty
        ReDim arSheetNumber(0)
        arSheetNumber(0) = stEntry
        Debug.Print stEntry & " entered as the first value in the array"
    Else
        If DoesValueExistInArray(arSheetNumber, stEntry) = False Then
            ReDim Preserve arSheetNumber(UBound(arSheetNumber) + 1) 'Increase the size of the array
            arSheetNumber(UBound(arSheetNumber)) = stEntry 'Add the new value to the array
            Debug.Print stEntry & " added to the array"
        Else
            Debug.Print stEntry & " already exists in the array"
        End If
    End If
    'Check that the current string is not already entered into the array
End Sub

'Check to see if there is a matching value in the array
Private Function DoesValueExistInArray(anArray As Variant, searchValue As Variant) As Boolean
    Dim lPos As Long
    
    If Not IsVarArrayEmpty(anArray) Then
        For lPos = 0 To UBound(anArray)
        If anArray(lPos) = searchValue Then
            DoesValueExistInArray = True
        End If
        Next lPos
    End If
End Function

'Check to see if an array is empty
Private Function IsVarArrayEmpty(anArray As Variant)
    Dim i As Integer

    On Error Resume Next
    i = UBound(anArray, 1)
    If Err.Number = 0 Then
        IsVarArrayEmpty = False
    Else
        IsVarArrayEmpty = True
    End If
End Function

'Call the database to remove the values stored
Private Sub clearDatabaseValues(anArray As Variant)
    Dim lPos As Long
    Dim stSQL As String
    
    If Not IsVarArrayEmpty(arSheetNumber) Then
        Dim tempString As String
        
        Call ADO_Conn.Open_Connection(stDronfieldSystem) 'Open the connection to a different database.
        For lPos = 0 To UBound(arSheetNumber)
            stSQL = "DELETE FROM [@Pickings] WHERE [sheetNumber] = '" & arSheetNumber(lPos) & "';"
            ADO_Conn.Conn.Execute stSQL
            Debug.Print "database execution string: " & stSQL
        Next lPos
        Call ADO_Conn.Close_Connection
        tempString = ADO_Conn.returnLoadedDatabaseLocation
        Debug.Print tempString
    End If
End Sub

'Clear entries from any exisiting date that matches the same date
Private Function clearPreviousData() As Long
    Dim lookR As Range, rCell As Range
    Dim finalRow As Long
    
    Set lookR = wsConfirmation.Range("B6")
    
    If lookR.Value = "Pick sheet number" Then
        'Correct Column found
        finalRow = wsConfirmation.Range("B10000").End(xlUp).Row
        If finalRow > 6 Then
            Set lookR = wsConfirmation.Range("B7:B" & finalRow)
            For Each rCell In lookR
                If rCell.Value <> "" Then
                    Call addToDatabaseArray(rCell.Value) 'Add the value to the array for database processing.
                End If
            Next rCell
            Call clearDatabaseValues(arSheetNumber) 'Call the delete SQL to remove the value from the database that are in the array.
            Set lookR = Nothing
            clearPreviousData = 0
        Else
            Debug.Print "No values entered into the picking confirmation sheet."
            End
            clearPreviousData = -1
        End If
    Else
        'Incorrect column found, return error and stop the program.
    End If
End Function

'collect the data to be entered into the database
Private Function collectNewData() As clsPickingFigures
    Dim lookR As Range, rCell As Range
    Dim finalRow As Long, currentRow As Long
    Dim pickingValues As clsPickingFigure
    Dim pickCollection As New clsPickingFigures
        
    'Set the range to read.
    finalRow = wsConfirmation.Range("B10000").End(xlUp).Row
    Set lookR = wsConfirmation.Range("B7:B" & finalRow)
    For Each rCell In lookR
        currentRow = rCell.Row
        If Not wsConfirmation.Range("E" & currentRow).Value = "" Then
            Set pickingValues = New clsPickingFigure
            With pickingValues
                .sheetNumber = rCell.Value
                .pickDate = wsConfirmation.Range("A" & currentRow).Value
                .operatorID = wsConfirmation.Range("U" & currentRow).Value
                .casesQty = wsConfirmation.Range("C" & currentRow).Value
                .singlesQty = wsConfirmation.Range("D" & currentRow).Value
                .reasonCode = wsConfirmation.Range("F" & currentRow).Value
                .productCode = wsConfirmation.Range("Q" & currentRow).Value
            End With
            pickCollection.Add pickingValues
            Set pickingValues = Nothing
        End If
    Next rCell
    
    Set collectNewData = pickCollection
    
    Set lookR = Nothing
    'Add the value to the Picking Object Array.
End Function

'Add data to the database
Private Sub addToDatabase(pickCollection As clsPickingFigures)
    Dim stSQL As String
    Dim pickEntry As clsPickingFigure
        
    'Open the database connection (custom connection required).
    ADO_Conn.Open_Connection stDronfieldSystem
    
    For Each pickEntry In pickCollection.Items
        'Create the string for each of the entries for the database.
        stSQL = "INSERT INTO [@Pickings] ([sheetNumber], [pickDate], [employeeID], [productCode], [singlePicks], [casePicks]) " & _
                "VALUES ('" & pickEntry.sheetNumber & "', #" & Format(pickEntry.pickDate, "mm/dd/yyyy") & "#, " & pickEntry.operatorID & ", '" & pickEntry.productCode & "', " & pickEntry.singlesQty & ", " & pickEntry.casesQty & ");"
        Debug.Print "stSQL : " & stSQL
        
        'Execute the SQL code for the database entries.
        ADO_Conn.Conn.Execute stSQL
    Next pickEntry
    
    'close the connection to the database.
    ADO_Conn.Close_Connection
    
    Set pickCollection = Nothing
End Sub

'Removes objects from memory
Private Sub cleanUp()
    Erase arSheetNumber
    Set wsConfirmation = Nothing
End Sub
