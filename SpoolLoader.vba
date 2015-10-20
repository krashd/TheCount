Public Const versionNumber = 0.42
Public Const versionType = "Custom"
Public Const versionSite = "Dronfield"

Dim arrOutput() As Variant
Dim arrPickSheetData As Variant

Const selArrayMethod = 1      'Array method selector, purely for testing, 0 = Multiple, 1 = Single Large

Dim outputSheetPosition As Long
Dim outputSheet As Worksheet
Dim currentWB As Workbook

Private Sub setupVariables()
    DebugText.PrintText "Loading Values Started"
    
    If selArrayMethod = 0 Then
        DebugText.PrintText "Multi Array Method Selected"
    ElseIf selArrayMethod = 1 Then
        DebugText.PrintText "Single Large Array Method Selected"
    End If
    
    readingHeader = True
End Sub
    
Public Sub Load_Download_Spools()
    Dim arrProcessedString() As Variant
    Dim arrSplitString() As String

    Dim lProcessedStringPosition As Long
    Dim lSplitStringArrayPosition As Long
    Dim lArrayLimit As Long
    
    Dim inputFileName As String
    
    Dim fileNum As Long
    Dim dataLine As Variant
    Dim fileName As String
    Dim x As Long
        
    Call setupVariables
    
    inputFileName = Application.GetOpenFilename("Spool File (SPOOL.*), SPOOL.*", , "Select Spool File")
    If inputFileName = "False" Then
        MsgBox "No Spool File Selected, Aborting Import"
        Exit Sub
    End If
    
    fileNum = FreeFile()
    
    Open inputFileName For Input As #fileNum    'Open input file for reading

    lProcessedStringPosition = 0
    
    ReDim arrProcessedString(0)
    
    Load frmProgress
    frmProgress.Show
    frmProgress.bar.Width = 0
    frmProgress.Repaint
    previousPercentage = 0
    updateBreaker = 0
    
    While Not EOF(fileNum)
        Line Input #fileNum, dataLine
        arrSplitString = Split(dataLine, vbLf)              'Split the incoming data line into seperate lines
        
        frmProgress.frame.Text = "Reading Spool File"
        frmProgress.Repaint
        
        For lSplitStringArrayPosition = 0 To UBound(arrSplitString)
            ReDim Preserve arrProcessedString(lProcessedStringPosition)
            arrProcessedString(lProcessedStringPosition) = Replace(arrSplitString(lSplitStringArrayPosition), vbLf, vbCrLf)
            arrProcessedString(lProcessedStringPosition) = Replace(arrProcessedString(lProcessedStringPosition), Chr(18), "")
            arrProcessedString(lProcessedStringPosition) = Replace(arrProcessedString(lProcessedStringPosition), Chr(20), "")
            lProcessedStringPosition = lProcessedStringPosition + 1
        Next lSplitStringArrayPosition
        
    Wend
    Close fileNum
    
    frmProgress.frame.Text = "Processing Spool"
    frmProgress.Repaint
    
    arrPickSheetData = DR_PickingSpoolLoader.processSpoolStream(arrProcessedString)
    
    DebugText.PrintText "Values Calculation Finished"
    
    Call loadOutputSheet
    'Call CreateTestOutput
    Call outputToWorksheet
    DebugText.PrintText "Data Entry Completed"
    Call cleanUp
    frmProgress.frame.Text = "Refreshing Data"
    ThisWorkbook.RefreshAll
    Unload frmProgress
End Sub

Private Sub updateProductsList()
    Dim lTemp As Long
    lTemp = ADO_Conn.UpdateProductList(outerBarcode, singlesBarcode, productDescription, productPackSize, supplier, grnDate)
End Sub

Public Sub Clean_Formulas_Sheets()
    ThisWorkbook.Sheets("Formulas").Range("A2:S59999").ClearContents
    MsgBox "Formulas Cleared"
End Sub

Private Sub enterOutput(Optional arrayPosition As Long)
'Locations (Columns) for each data entry
    Dim wsTest As Worksheet
    Dim lFreeRow As Long
    Dim lUsedRow As Long
    Dim lLastArrayRow As Long
    
    Set wsTest = ThisWorkbook.Sheets("Formulas")
    lUsedRow = wsTest.Range("A59999").End(xlUp).Row
    
    ' Find the last Empty Row
    If lUsedRow = 1 Then
        If wsTest.Range("A1").Value = "" Or IsNull(wsTest.Range("A1")) Then
            lFreeRow = 1
        Else
            lFreeRow = lUsedRow + 1
        End If
    Else
        lFreeRow = lUsedRow + 1
    End If

    arrPickSheetData = Application.Transpose(arrPickSheetData)
    lLastArrayRow = UBound(arrPickSheetData)
    wsTest.Range("A" & lFreeRow & ":S" & lFreeRow + lLastArrayRow - 1).Value = arrPickSheetData
End Sub

Private Sub outputToWorksheet()
    Application.ScreenUpdating = False
    Call enterOutput
    Application.ScreenUpdating = True
    Call updateRangeName
End Sub

Public Sub updateRangeName()
    Call updateFormulasRange
    Call updateOperatorRange
End Sub

Private Sub updateFormulasRange()
    Dim wsOutput As Worksheet
    Dim lLastRow As Long

    Set wsOutput = ThisWorkbook.Sheets("Formulas")
    
    lLastRow = wsOutput.Range("A59999").End(xlUp).Row
    wsOutput.Range("A1:S" & lLastRow).Name = "Data"
    
    Set wsOutput = Nothing
End Sub

Private Sub updateOperatorRange()
    Dim wsLists As Worksheet
    Dim lLastRow As Long

    Set wsLists = ThisWorkbook.Sheets("Lists")
    
    lLastRow = wsLists.Range("E59999").End(xlUp).Row
    'sort ranges by operator codes (ascending)
    wsLists.Range("E1:H" & lLastRow).Sort key1:=wsLists.Range("E1"), Order1:=xlAscending, Header:=xlYes, Orientation:=xlSortColumns, dataoption1:=xlSortNormal
    
    lLastRow = wsLists.Range("E59999").End(xlUp).Row
    'define range names
    wsLists.Range("E1:H" & lLastRow).Name = "Operators"
    wsLists.Range("E1:E" & lLastRow).Name = "Operator_codes"
    
    Set wsLists = Nothing
End Sub

Private Function nextEmptyRow(targetWorksheet As Worksheet) As Long
'Finds the last empty row of a worksheet (using the C coloumn)
    Dim returnedRow As Long
    
    returnedRow = targetWorksheet.Range("C59999").End(xlUp).Row
    If Cells(returnedRow, 1).Value <> "" Then
        nextEmptyRow = returnedRow + 1
    Else
        nextEmptyRow = returnedRow
    End If
End Function

Private Sub loadOutputSheet()
'Loads the output worksheet and finds the last unused row for data entry
    Set currentWB = ThisWorkbook
    Set outputSheet = currentWB.Sheets("Formulas")
End Sub

Private Sub cleanUp()
    Dim nullLong As Long
'Clean up all objects, close any open files prior to exiting
    Set outputSheet = Nothing
    Set currentWB = Nothing
    Erase arrPickSheetData
    Call DR_PickingSpoolLoader.eraseArray
End Sub
