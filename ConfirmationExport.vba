Private Const helpFile = "U:\Boots_Contract\Chill\Dronfield\Dude's Folder\Information\Thermo BAC55.pdf"
Private wsConfirmation As Worksheet

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
    
End Sub

'Clear entries from any exisiting date that matches the same date
Private Sub clearPreviousData()
    Dim lookR As Range, c As Range
    Set lookR = wsConfirmation.Range("TestRange")
    
    On Error GoTo ErrorHandler
    
    For Each rCells In lookR
            'Let's see if it's a combined account
            Debug.Print rCells.Value
    Next rCells
    
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

    Set wsConfirmation = Nothing
End Sub
