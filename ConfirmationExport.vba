Private Const helpFile = "U:\Boots_Contract\Chill\Dronfield\Dude's Folder\Information\Thermo BAC55.pdf"
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

'Main sub for the entire export process
Public Sub main()
    Call initialiseValues               'Load values to be used
    Call clearPreviousData              'Clear the entries matching picking sheet number from the database.
    Call collectNewData                 'Collect the new data from the sheet
    Call addToDatabase                  'Enter the data to the database.
    Call cleanUp                        'Clear up the values
End Sub

'Set up the main values to be used
Private Sub initialiseValues()
    Set wsConfirmation = ThisWorkbook.Sheets("Pick Confirmation")
    Erase arSheetNumber
End Sub

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

'Clear entries from any exisiting date that matches the same date
Private Sub clearPreviousData()
    Dim lookR As Range, rCell As Range
    Dim finalRow As Long
    
    Set lookR = wsConfirmation.Range("B6")
    
    If lookR.Value = "Pick sheet number" Then
        'Correct Column found
        finalRow = wsConfirmation.Range("B10000").End(xlUp).Row
        Set lookR = wsConfirmation.Range("B7:B" & finalRow)
        For Each rCell In lookR
            If rCell.Value <> "" Then
                Call addToDatabaseArray(rCell.Value) 'Add the value to the array for database processing.
            End If
        Next rCell
        'Call the delete SQL to remove the value from the database that are in the array.
    Else
        'Incorrect column found, return error and stop the program.
    End If
    
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    'Collect the date of the sheet
    'Run an SQL command to clear the data from the database.
    
ErrorHandler:
    Debug.Print ErrorCode
    Resume
End Sub

'collect the data to be entered into the database
Private Sub collectNewData()
    
End Sub

'Add data to the database
Private Sub addToDatabase()

End Sub

'Removes objects from memory
Private Sub cleanUp()
    Erase arSheetNumber
    Set wsConfirmation = Nothing
End Sub
