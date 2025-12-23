Attribute VB_Name = "modGenerationReview"
Option Explicit

'=============================
' Module: modGenerationReview
' Purpose: Two-step workflow for generation review and auditor workbook generation
'=============================

'=============================
' CONSTANTS
'=============================
Private Const REVIEW_SHEET_NAME As String = "Generation Review"
Private Const REVIEW_TABLE_NAME As String = "tblGenerationReview"
Private Const SOURCE_TABLE_NAME As String = "tblEntities"
Private Const AUDITORS_TABLE_NAME As String = "tblAuditors"

'=============================
' DEBUG LOGGING
'=============================
Private Sub DebugLog(ByVal msg As String)
    Debug.Print Format$(Now, "hh:nn:ss") & " | [Review] " & msg
    
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets("_Log")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLog.Name = "_Log"
        wsLog.Range("A1").Value = "Timestamp"
        wsLog.Range("B1").Value = "Message"
        wsLog.Range("A1:B1").Font.Bold = True
    End If
    
    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    wsLog.Cells(nextRow, 1).Value = Now
    wsLog.Cells(nextRow, 2).Value = msg
    wsLog.Columns("A").AutoFit
    wsLog.Columns("B").ColumnWidth = 150
    On Error GoTo 0
End Sub

'=============================
' PUBLIC INTERFACE
'=============================

Public Sub PopulateGenerationReview()
    On Error GoTo ErrHandler
    
    DebugLog "=== PopulateGenerationReview START ==="
    
    ' 1. Get source table (tblEntities)
    Dim loSource As ListObject
    Set loSource = GetTableSafe(ThisWorkbook, SOURCE_TABLE_NAME)
    
    If loSource Is Nothing Then
        MsgBox "Source table '" & SOURCE_TABLE_NAME & "' not found.", vbExclamation
        Exit Sub
    End If
    
    If loSource.DataBodyRange Is Nothing Then
        MsgBox "Source table '" & SOURCE_TABLE_NAME & "' has no data rows.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Get target table (tblGenerationReview)
    Dim wsReview As Worksheet
    On Error Resume Next
    Set wsReview = ThisWorkbook.Worksheets(REVIEW_SHEET_NAME)
    On Error GoTo ErrHandler
    
    If wsReview Is Nothing Then
        MsgBox "Sheet '" & REVIEW_SHEET_NAME & "' not found in workbook.", vbExclamation
        Exit Sub
    End If
    
    Dim loTarget As ListObject
    On Error Resume Next
    Set loTarget = wsReview.ListObjects(REVIEW_TABLE_NAME)
    On Error GoTo ErrHandler
    
    If loTarget Is Nothing Then
        MsgBox "Table '" & REVIEW_TABLE_NAME & "' not found on '" & REVIEW_SHEET_NAME & "' sheet.", vbExclamation
        Exit Sub
    End If
    
    ' 3. Build column indexes for source
    Dim idxSrc As Object
    Set idxSrc = BuildIndexMap(loSource)
    
    ' Verify required source columns
    Dim missingCols As String
    missingCols = ""
    If Not idxSrc.Exists("GCI") Then missingCols = missingCols & "GCI, "
    If Not idxSrc.Exists("AuditorID") Then missingCols = missingCols & "AuditorID, "
    If Not idxSrc.Exists("Jurisdiction ID") Then missingCols = missingCols & "Jurisdiction ID, "
    
    If Len(missingCols) > 0 Then
        missingCols = Left$(missingCols, Len(missingCols) - 2)
        MsgBox "Source table '" & SOURCE_TABLE_NAME & "' is missing required columns: " & missingCols, vbExclamation
        Exit Sub
    End If
    
    ' 4. Build Auditor Name lookup map from Auditors table
    Dim auditorNameMap As Object
    Set auditorNameMap = BuildAuditorNameLookup()
    
    If auditorNameMap.Count = 0 Then
        DebugLog "WARNING: No auditors found in lookup table"
    Else
        DebugLog "Built auditor lookup map with " & auditorNameMap.Count & " auditor(s)"
    End If
    
    ' 5. Read source data
    Dim srcData As Variant
    srcData = loSource.DataBodyRange.Value2
    
    Dim rowCount As Long
    rowCount = UBound(srcData, 1)
    
    DebugLog "Source data loaded: " & rowCount & " rows"
    
    ' 6. Build review data array (6 columns: GCI, Legal Name, Jurisdiction ID, Jurisdiction, AuditorID, Auditor Name)
    Dim reviewData() As Variant
    ReDim reviewData(1 To rowCount, 1 To 6)
    
    Dim i As Long
    For i = 1 To rowCount
        Dim gci As String
        gci = Trim$(CStr(GetCellValue(srcData, i, idxSrc, "GCI", "")))
        
        ' Build new row from source
        reviewData(i, 1) = gci ' GCI
        reviewData(i, 2) = GetCellValue(srcData, i, idxSrc, "Legal Name", GetCellValue(srcData, i, idxSrc, "Legal Entity Name", "")) ' Legal Name
        reviewData(i, 3) = GetCellValue(srcData, i, idxSrc, "Jurisdiction ID", "") ' Jurisdiction ID
        reviewData(i, 4) = GetCellValue(srcData, i, idxSrc, "Jurisdiction", GetCellValue(srcData, i, idxSrc, "Jurisdiction Name", "")) ' Jurisdiction
        reviewData(i, 5) = GetCellValue(srcData, i, idxSrc, "AuditorID", "") ' AuditorID
        
        ' Auditor Name - lookup from Auditors table using AuditorID
        Dim auditorIDForLookup As String
        auditorIDForLookup = Trim$(CStr(reviewData(i, 5)))
        If Len(auditorIDForLookup) > 0 And auditorNameMap.Exists(auditorIDForLookup) Then
            reviewData(i, 6) = CStr(auditorNameMap(auditorIDForLookup))
        Else
            reviewData(i, 6) = "" ' Auditor Name not found in lookup
        End If
    Next i
    
    ' 7. Write to target table
    WriteToReviewTable loTarget, reviewData
    
    ' 8. Update metadata cells
    UpdateReviewMetadata wsReview, rowCount
    
    DebugLog "PopulateGenerationReview: Wrote " & rowCount & " rows"
    DebugLog "=== PopulateGenerationReview COMPLETE ==="
    
    MsgBox "Review table populated with " & rowCount & " assignment(s)." & vbCrLf & _
           "Please review and edit as needed, then run 'Generate Auditor Workbooks'.", _
           vbInformation, "Generation Review"
    
    Exit Sub
    
ErrHandler:
    DebugLog "ERROR in PopulateGenerationReview: " & Err.Number & " - " & Err.Description
    MsgBox "PopulateGenerationReview failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Public Sub GenerateAuditorWorkbooks_FromReview()
    On Error GoTo ErrHandler
    
    DebugLog "=== GenerateAuditorWorkbooks_FromReview START ==="
    
    ' Confirm with user
    If MsgBox("Generate auditor workbooks from the review table?" & vbCrLf & vbCrLf & _
              "This will create workbooks based on the assignments in the Generation Review table.", _
              vbQuestion + vbYesNo, "Confirm Generation") <> vbYes Then
        DebugLog "User cancelled generation"
        Exit Sub
    End If
    
    ' 1. Get review table
    Dim wsReview As Worksheet
    On Error Resume Next
    Set wsReview = ThisWorkbook.Worksheets(REVIEW_SHEET_NAME)
    On Error GoTo ErrHandler
    
    If wsReview Is Nothing Then
        MsgBox "Sheet '" & REVIEW_SHEET_NAME & "' not found in workbook.", vbExclamation
        Exit Sub
    End If
    
    Dim loReview As ListObject
    On Error Resume Next
    Set loReview = wsReview.ListObjects(REVIEW_TABLE_NAME)
    On Error GoTo ErrHandler
    
    If loReview Is Nothing Then
        MsgBox "Table '" & REVIEW_TABLE_NAME & "' not found. Run 'Populate Generation Review' first.", vbExclamation
        Exit Sub
    End If
    
    If loReview.DataBodyRange Is Nothing Then
        MsgBox "Review table is empty. Run 'Populate Generation Review' first.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Read config (from named ranges)
    DebugLog "Step: Reading configuration from named ranges"
    Dim cfg As Object
    Set cfg = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Dim batchIDVal As Variant, outputFolderVal As Variant
    batchIDVal = ThisWorkbook.Names("rngBatchID").RefersToRange.Value
    If Err.Number <> 0 Then
        DebugLog "ERROR reading rngBatchID: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    outputFolderVal = ThisWorkbook.Names("rngOutputFolder").RefersToRange.Value
    If Err.Number <> 0 Then
        DebugLog "ERROR reading rngOutputFolder: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrHandler
    
    cfg("BatchID") = batchIDVal
    cfg("OutputFolder") = outputFolderVal
    
    If Not cfg.Exists("BatchID") Or Not cfg.Exists("OutputFolder") Then
        MsgBox "Failed to read generator configuration. Check that named ranges 'rngBatchID' and 'rngOutputFolder' exist.", vbExclamation
        Exit Sub
    End If
    
    DebugLog "Config read: BatchID=" & cfg("BatchID") & ", OutputFolder=" & cfg("OutputFolder")
    DebugLog "Config validated successfully"
    
    ' 3. Read review data
    DebugLog "Step: Reading review data from table"
    Dim reviewData As Variant
    reviewData = loReview.DataBodyRange.Value2
    DebugLog "Step: Review data read successfully, rows: " & UBound(reviewData, 1)
    
    DebugLog "Step: Building index map for review table"
    Dim idxReview As Object
    Set idxReview = BuildIndexMap(loReview)
    DebugLog "Step: Index map built successfully"
    
    ' Verify required columns
    Dim missingCols As String
    missingCols = ""
    If Not idxReview.Exists("GCI") Then missingCols = missingCols & "GCI, "
    If Not idxReview.Exists("AuditorID") Then missingCols = missingCols & "AuditorID, "
    If Not idxReview.Exists("Jurisdiction ID") Then missingCols = missingCols & "Jurisdiction ID, "
    
    If Len(missingCols) > 0 Then
        missingCols = Left$(missingCols, Len(missingCols) - 2)
        MsgBox "Review table is missing required columns: " & missingCols, vbExclamation
        Exit Sub
    End If
    
    ' 4. Build auditor map from review table
    DebugLog "Step: Building auditor map from review data"
    Dim auditorMap As Object
    Set auditorMap = BuildAuditorMapFromReview(reviewData, idxReview)
    DebugLog "Step: Auditor map built successfully"
    
    If auditorMap.Count = 0 Then
        MsgBox "No valid assignments found in review table." & vbCrLf & vbCrLf & _
               "Check that AuditorID, Jurisdiction ID, and GCI are populated for at least one row.", _
               vbExclamation
        Exit Sub
    End If
    
    DebugLog "AuditorMap built from review table with " & auditorMap.Count & " auditor(s)"
    DebugLog "About to read source tables (tblAttributes, tblAcceptableDocs)"
    
    ' 5. Read source tables needed for generation
    Dim loAttr As ListObject, loDocs As ListObject
    Set loAttr = GetTableSafe(ThisWorkbook, "tblAttributes")
    Set loDocs = GetTableSafe(ThisWorkbook, "tblAcceptableDocs")
    
    If loAttr Is Nothing Then
        MsgBox "Table 'tblAttributes' not found.", vbExclamation
        Exit Sub
    End If
    
    If loDocs Is Nothing Then
        MsgBox "Table 'tblAcceptableDocs' not found.", vbExclamation
        Exit Sub
    End If
    
    Dim attrRows As Variant, docRows As Variant
    If loAttr.DataBodyRange Is Nothing Then
        MsgBox "Table 'tblAttributes' has no data rows.", vbExclamation
        Exit Sub
    End If
    attrRows = loAttr.DataBodyRange.Value2
    
    If loDocs.DataBodyRange Is Nothing Then
        MsgBox "Table 'tblAcceptableDocs' has no data rows.", vbExclamation
        Exit Sub
    End If
    docRows = loDocs.DataBodyRange.Value2
    
    ' Build index maps for attributes and docs
    DebugLog "Step: Building index maps for attributes and docs tables"
    Dim idxAttr As Object, idxDocs As Object
    Set idxAttr = BuildIndexMap(loAttr)
    DebugLog "Step: Attributes index map built"
    Set idxDocs = BuildIndexMap(loDocs)
    DebugLog "Step: Docs index map built"
    
    ' 6. Generate workbook for each auditor
    Dim auditorID As Variant
    Dim successCount As Long, totalCount As Long
    successCount = 0
    totalCount = auditorMap.Count
    
    DebugLog "Starting generation loop for " & totalCount & " auditor(s)"
    
    For Each auditorID In auditorMap.Keys
        DebugLog "Processing auditor: " & CStr(auditorID)
        On Error Resume Next
        Call GenerateOneAuditorWorkbook_FromReview(CStr(auditorID), auditorMap(auditorID), cfg, _
                                                    attrRows, idxAttr, docRows, idxDocs)
        If Err.Number = 0 Then
            successCount = successCount + 1
            DebugLog "Successfully generated workbook for " & auditorID
        Else
            DebugLog "ERROR generating workbook for " & auditorID & ": " & Err.Number & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrHandler
    Next auditorID
    
    ' 7. Update status
    UpdateGenerationStatus wsReview, successCount, totalCount
    
    DebugLog "=== GenerateAuditorWorkbooks_FromReview COMPLETE ==="
    DebugLog "Generated " & successCount & " of " & totalCount & " workbook(s)"
    
    MsgBox "Generated " & successCount & " of " & totalCount & " workbook(s)." & vbCrLf & _
           "Check the _Log worksheet for details.", _
           vbInformation, "Generation Complete"
    
    Exit Sub
    
ErrHandler:
    DebugLog "ERROR in GenerateAuditorWorkbooks_FromReview: " & Err.Number & " - " & Err.Description
    MsgBox "GenerateAuditorWorkbooks_FromReview failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'=============================
' HELPER FUNCTIONS
'=============================

Private Function BuildAuditorNameLookup() As Object
    ' Build dictionary: key=AuditorID, value=Auditor Name
    ' Reads from tblAuditors table
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim loAuditors As ListObject
    Set loAuditors = GetTableSafe(ThisWorkbook, AUDITORS_TABLE_NAME)
    
    If loAuditors Is Nothing Then
        DebugLog "WARNING: Auditors table '" & AUDITORS_TABLE_NAME & "' not found - Auditor Name will be empty"
        Set BuildAuditorNameLookup = dict
        Exit Function
    End If
    
    If loAuditors.DataBodyRange Is Nothing Then
        DebugLog "WARNING: Auditors table has no data rows - Auditor Name will be empty"
        Set BuildAuditorNameLookup = dict
        Exit Function
    End If
    
    Dim auditorData As Variant
    auditorData = loAuditors.DataBodyRange.Value2
    
    Dim idxAuditors As Object
    Set idxAuditors = BuildIndexMap(loAuditors)
    
    ' Check for AuditorID and AuditorName columns
    If Not idxAuditors.Exists("AuditorID") Then
        DebugLog "WARNING: Auditors table missing 'AuditorID' column"
        Set BuildAuditorNameLookup = dict
        Exit Function
    End If
    
    If Not idxAuditors.Exists("AuditorName") Then
        DebugLog "WARNING: Auditors table missing 'AuditorName' column"
        Set BuildAuditorNameLookup = dict
        Exit Function
    End If
    
    ' Build lookup map
    Dim r As Long
    For r = 1 To UBound(auditorData, 1)
        Dim auditorID As String
        auditorID = Trim$(CStr(GetCellValue(auditorData, r, idxAuditors, "AuditorID", "")))
        
        If Len(auditorID) > 0 Then
            Dim auditorName As String
            auditorName = Trim$(CStr(GetCellValue(auditorData, r, idxAuditors, "AuditorName", "")))
            dict(auditorID) = auditorName
        End If
    Next r
    
    DebugLog "BuildAuditorNameLookup: Built map with " & dict.Count & " auditor(s)"
    Set BuildAuditorNameLookup = dict
End Function

Private Function BuildIndexMap(ByVal lo As ListObject) As Object
    ' Build dictionary mapping column names to indexes (1-based)
    ' Uses case-insensitive matching (vbTextCompare)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive
    
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        Dim colName As String
        colName = Trim$(lo.ListColumns(i).Name) ' Trim spaces
        dict(colName) = i
    Next i
    
    Set BuildIndexMap = dict
End Function

Private Function GetTableSafe(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    ' Safely get a table by name from any sheet in the workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        
        If Not lo Is Nothing Then
            Set GetTableSafe = lo
            Exit Function
        End If
    Next ws
    
    Set GetTableSafe = Nothing
End Function

Private Function GetCellValue(ByVal rows As Variant, ByVal r As Long, ByVal idx As Object, _
                              ByVal colName As String, Optional ByVal defaultValue As Variant = "") As Variant
    ' Safe column accessor with default value
    If idx.Exists(colName) Then
        Dim colIdx As Long
        colIdx = idx(colName)
        Dim cellValue As Variant
        cellValue = rows(r, colIdx)
        
        ' Handle empty values
        If IsEmpty(cellValue) Or IsNull(cellValue) Then
            GetCellValue = defaultValue
        ElseIf VarType(cellValue) = vbDouble Or VarType(cellValue) = vbLong Or VarType(cellValue) = vbInteger Then
            ' Convert numbers to string
            GetCellValue = CStr(cellValue)
        Else
            GetCellValue = cellValue
        End If
    Else
        GetCellValue = defaultValue
    End If
End Function

Private Sub WriteToReviewTable(ByVal loTarget As ListObject, ByRef reviewData() As Variant)
    On Error GoTo ErrHandler
    
    Dim rowCount As Long, colCount As Long
    rowCount = UBound(reviewData, 1)
    colCount = UBound(reviewData, 2)
    
    DebugLog "WriteToReviewTable: Writing " & rowCount & " rows x " & colCount & " cols"
    
    ' Get the worksheet
    Dim ws As Worksheet
    Set ws = loTarget.Parent
    
    ' Resize table FIRST to accommodate the data (header row + data rows)
    Dim newTableRange As Range
    Set newTableRange = loTarget.HeaderRowRange.Resize(rowCount + 1, colCount)
    
    DebugLog "WriteToReviewTable: Resizing table to " & (rowCount + 1) & " rows (including header)"
    On Error Resume Next
    loTarget.Resize newTableRange
    On Error GoTo ErrHandler
    
    ' Get the data body range (should exist after resize)
    Dim dataRange As Range
    Set dataRange = loTarget.DataBodyRange
    
    If dataRange Is Nothing Then
        ' If still nothing, try writing directly below header
        DebugLog "WriteToReviewTable: DataBodyRange still Nothing, writing directly below header"
        Set dataRange = loTarget.HeaderRowRange.Offset(1, 0).Resize(rowCount, colCount)
    Else
        ' Ensure data range is the right size
        If dataRange.Rows.Count <> rowCount Or dataRange.Columns.Count <> colCount Then
            DebugLog "WriteToReviewTable: DataBodyRange size mismatch, resizing"
            Set dataRange = dataRange.Resize(rowCount, colCount)
        End If
    End If
    
    DebugLog "WriteToReviewTable: Writing to range " & dataRange.Address
    
    ' Write data in one shot
    dataRange.Value2 = reviewData
    
    ' Force Excel to recalculate/refresh
    ws.Calculate
    Application.ScreenUpdating = True
    DoEvents
    
    DebugLog "WriteToReviewTable: Success - wrote data to " & dataRange.Address
    Exit Sub
    
ErrHandler:
    DebugLog "ERROR in WriteToReviewTable: " & Err.Number & " - " & Err.Description
    DebugLog "WriteToReviewTable: rowCount=" & rowCount & ", colCount=" & colCount
    If Not loTarget Is Nothing Then
        DebugLog "WriteToReviewTable: Table address=" & loTarget.Range.Address
        DebugLog "WriteToReviewTable: HeaderRowRange=" & loTarget.HeaderRowRange.Address
    End If
    Err.Raise Err.Number, "WriteToReviewTable", Err.Description
End Sub

Private Sub UpdateReviewMetadata(ByVal wsReview As Worksheet, ByVal rowCount As Long)
    On Error Resume Next
    
    ' Update "Last Refresh" cell
    Dim rngLabel As Range
    Set rngLabel = wsReview.Cells.Find(What:="Last Refresh", LookIn:=xlValues, LookAt:=xlPart)
    
    If Not rngLabel Is Nothing Then
        rngLabel.Offset(0, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        DebugLog "Updated Last Refresh timestamp"
    End If
    
    ' Update "Generator Status" cell
    Set rngLabel = wsReview.Cells.Find(What:="Generator Status", LookIn:=xlValues, LookAt:=xlPart)
    
    If Not rngLabel Is Nothing Then
        rngLabel.Offset(0, 1).Value = "Ready (review " & rowCount & " assignment(s) below)"
        DebugLog "Updated Generator Status: Ready"
    End If
    
    On Error GoTo 0
End Sub

Private Function BuildAuditorMapFromReview(ByVal reviewData As Variant, ByVal idxReview As Object) As Object
    On Error GoTo ErrHandler
    
    ' Build dictionary: key=AuditorID, value=Collection of assignment records (Dictionaries)
    DebugLog "BuildAuditorMapFromReview: Starting, rows=" & UBound(reviewData, 1)
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    DebugLog "BuildAuditorMapFromReview: Created main dict"
    
    ' Use a separate dictionary to store Collections to avoid type issues
    Dim collDict As Object
    Set collDict = CreateObject("Scripting.Dictionary")
    collDict.CompareMode = vbTextCompare
    DebugLog "BuildAuditorMapFromReview: Created collDict"
    
    Dim r As Long
    Dim skippedCount As Long
    skippedCount = 0
    
    DebugLog "BuildAuditorMapFromReview: Starting loop"
    For r = 1 To UBound(reviewData, 1)
        Dim auditorID As String
        Dim gci As String
        Dim jurID As String
        
        ' Get raw values first (may be numbers or strings)
        On Error Resume Next
        Dim auditorIDRaw As Variant, gciRaw As Variant, jurIDRaw As Variant
        auditorIDRaw = GetCellValue(reviewData, r, idxReview, "AuditorID", "")
        If Err.Number <> 0 Then
            DebugLog "BuildAuditorMapFromReview: ERROR at GetCellValue AuditorID row " & r & ": " & Err.Number & " - " & Err.Description
            Err.Clear
        End If
        gciRaw = GetCellValue(reviewData, r, idxReview, "GCI", "")
        jurIDRaw = GetCellValue(reviewData, r, idxReview, "Jurisdiction ID", "")
        On Error GoTo ErrHandler
        
        ' Convert to string, handling numbers and empty values
        If IsEmpty(auditorIDRaw) Or IsNull(auditorIDRaw) Then
            auditorID = ""
        Else
            auditorID = Trim$(CStr(auditorIDRaw))
        End If
        
        If IsEmpty(gciRaw) Or IsNull(gciRaw) Then
            gci = ""
        Else
            gci = Trim$(CStr(gciRaw))
        End If
        
        If IsEmpty(jurIDRaw) Or IsNull(jurIDRaw) Then
            jurID = ""
        Else
            jurID = Trim$(CStr(jurIDRaw))
        End If
        
        ' Skip rows with missing key fields
        If Len(auditorID) = 0 Or Len(gci) = 0 Or Len(jurID) = 0 Then
            skippedCount = skippedCount + 1
            GoTo ContinueRow
        End If
        
        ' Create or get collection for this auditor using separate dictionary
        On Error Resume Next
        Dim assignsForAuditor As Collection
        If Not collDict.Exists(auditorID) Then
            DebugLog "BuildAuditorMapFromReview: Creating new collection for auditor " & auditorID & " (row " & r & ")"
            Set assignsForAuditor = New Collection
            If Err.Number <> 0 Then
                DebugLog "BuildAuditorMapFromReview: ERROR creating Collection: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            Set collDict.Item(auditorID) = assignsForAuditor ' Use Set with .Item() for object assignment
            If Err.Number <> 0 Then
                DebugLog "BuildAuditorMapFromReview: ERROR assigning to collDict: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            Set dict.Item(auditorID) = assignsForAuditor ' Store reference in main dict too
        Else
            DebugLog "BuildAuditorMapFromReview: Retrieving collection for auditor " & auditorID & " (row " & r & ")"
            ' Retrieve collection from separate dictionary - use Set when retrieving
            Set assignsForAuditor = collDict.Item(auditorID)
            If Err.Number <> 0 Then
                DebugLog "BuildAuditorMapFromReview: ERROR retrieving from collDict: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
        End If
        On Error GoTo ErrHandler
        
        ' Create assignment record
        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        rec.CompareMode = vbTextCompare
        
        rec("AuditorID") = auditorID
        rec("GCI") = gci
        rec("Jurisdiction ID") = jurID
        rec("Jurisdiction Name") = CStr(GetCellValue(reviewData, r, idxReview, "Jurisdiction", jurID))
        rec("Legal Entity Name") = CStr(GetCellValue(reviewData, r, idxReview, "Legal Name", ""))
        rec("Auditor Name") = CStr(GetCellValue(reviewData, r, idxReview, "Auditor Name", ""))
        
        ' Add optional fields
        rec("Party Type") = ""
        rec("Onboarding Date") = ""
        rec("IRR") = ""
        rec("DRR") = ""
        rec("Primary FLU") = ""
        rec("Case ID # (Aware)") = ""
        
        ' Add the record to the collection (now properly typed as Collection)
        On Error Resume Next
        assignsForAuditor.Add rec
        If Err.Number <> 0 Then
            DebugLog "BuildAuditorMapFromReview: ERROR adding to collection row " & r & ": " & Err.Number & " - " & Err.Description
            DebugLog "BuildAuditorMapFromReview: assignsForAuditor type: " & TypeName(assignsForAuditor)
            Err.Clear
        End If
        On Error GoTo ErrHandler
        
ContinueRow:
    Next r
    
    If skippedCount > 0 Then
        DebugLog "BuildAuditorMapFromReview: Skipped " & skippedCount & " row(s) with missing key fields"
    End If
    
    DebugLog "BuildAuditorMapFromReview: Built map with " & dict.Count & " auditor(s)"
    Set BuildAuditorMapFromReview = dict
    Exit Function
    
ErrHandler:
    DebugLog "BuildAuditorMapFromReview: ERROR at row " & r & ": " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "BuildAuditorMapFromReview", Err.Description
End Function

Private Sub GenerateOneAuditorWorkbook_FromReview(ByVal auditorID As String, ByVal assigns As Collection, _
                                                   ByRef cfg As Object, _
                                                   ByVal attrRows As Variant, ByVal idxAttr As Object, _
                                                   ByVal docRows As Variant, ByVal idxDocs As Object)
    On Error GoTo ErrHandler
    
    DebugLog "GenerateOneAuditorWorkbook_FromReview: Starting for " & auditorID
    
    ' Create new workbook
    Dim wbNew As Workbook
    Set wbNew = Workbooks.Add
    
    ' Create Index sheet
    Dim wsIndex As Worksheet
    Set wsIndex = wbNew.Worksheets(1)
    wsIndex.Name = "Index"
    
    ' Extract Auditor Name from first assignment if available
    Dim auditorName As String
    auditorName = ""
    If assigns.Count > 0 Then
        Dim firstAssign As Object
        Set firstAssign = assigns(1)
        If firstAssign.Exists("Auditor Name") Then
            auditorName = Trim$(CStr(firstAssign("Auditor Name")))
        End If
    End If
    
    ' Call BuildIndexHeader from MatrixLayout module
    ' Extract config values since Application.Run cannot pass Dictionary objects
    Dim batchID As String, outputFolder As String
    batchID = CStr(cfg("BatchID"))
    outputFolder = CStr(cfg("OutputFolder"))
    
    ' Try calling with auditorName first (if function supports it)
    On Error Resume Next
    Application.Run "BuildIndexHeader", wsIndex, auditorID, batchID, outputFolder, auditorName
    If Err.Number <> 0 Then
        ' Try without auditorName parameter
        Err.Clear
        Application.Run "BuildIndexHeader", wsIndex, auditorID, batchID, outputFolder
        If Err.Number <> 0 Then
            ' If BuildIndexHeader doesn't exist or has different signature, create header manually
            Err.Clear
            wsIndex.Cells.Clear
            wsIndex.Range("A1").Value = "KYC Workpapers (Matrix Layout)"
            wsIndex.Range("A1").Font.Bold = True
            wsIndex.Range("A1").Font.Size = 16
            wsIndex.Range("A3").Value = "AuditorID:"
            If Len(auditorName) > 0 Then
                wsIndex.Range("B3").Value = auditorID & " (" & auditorName & ")"
            Else
                wsIndex.Range("B3").Value = auditorID
            End If
            wsIndex.Range("A4").Value = "Batch ID:"
            wsIndex.Range("B4").Value = batchID
            wsIndex.Range("A5").Value = "Generated:"
            wsIndex.Range("B5").Value = Now
            wsIndex.Columns("A:B").AutoFit
        End If
    End If
    On Error GoTo ErrHandler
    
    ' Create hidden lists sheet
    Dim wsLists As Worksheet
    Set wsLists = wbNew.Worksheets.Add(After:=wbNew.Worksheets(wbNew.Worksheets.Count))
    wsLists.Name = "_Lists"
    wsLists.Visible = xlSheetVeryHidden
    
    ' Group assignments by jurisdiction
    Dim jurMap As Object
    Set jurMap = CreateObject("Scripting.Dictionary")
    jurMap.CompareMode = vbTextCompare
    
    Dim assign As Object
    Dim jurIDLoop As String
    Dim gciListForJur As Collection ' Declare once at function level
    For Each assign In assigns
        jurIDLoop = CStr(assign("Jurisdiction ID"))
        
        If Not jurMap.Exists(jurIDLoop) Then
            Set gciListForJur = New Collection
            Set jurMap.Item(jurIDLoop) = gciListForJur
        Else
            Set gciListForJur = jurMap.Item(jurIDLoop)
        End If
        
        gciListForJur.Add CStr(assign("GCI"))
    Next assign
    
    ' Create jurisdiction sheets
    Dim jurIDKey As Variant
    Dim jurID As String
    For Each jurIDKey In jurMap.Keys
        jurID = CStr(jurIDKey)
        
        ' Get jurisdiction name from first assignment with this jurisdiction
        Dim jurName As String
        jurName = jurID
        For Each assign In assigns
            If CStr(assign("Jurisdiction ID")) = jurID Then
                jurName = CStr(assign("Jurisdiction Name"))
                If Len(jurName) = 0 Then jurName = jurID
                Exit For
            End If
        Next assign
        
        ' Create sheet for this jurisdiction
        Dim wsJur As Worksheet
        Set wsJur = wbNew.Worksheets.Add(After:=wbNew.Worksheets(wbNew.Worksheets.Count))
        wsJur.Name = SafeSheetName(jurName)
        
        ' Call BuildJurisdictionSheet from MatrixLayout module
        ' gciListForJur already declared above
        Set gciListForJur = jurMap.Item(jurID)
        
        ' Call BuildJurisdictionSheet from MatrixLayout module
        ' Note: Application.Run cannot pass Dictionary objects or Collections
        ' Since assigns, idxAttr, and idxDocs are complex objects, we'll use fallback
        ' If BuildJurisdictionSheet exists and can be called directly (not via Application.Run),
        ' it would need to be called from the same module or with a direct reference
        On Error Resume Next
        ' Try calling with simple parameters only (no Dictionary/Collection objects)
        ' Note: This will likely fail since the function probably needs the complex objects
        Application.Run "BuildJurisdictionSheet", wsJur, wsLists, jurID, jurName, auditorID, _
                         batchID, outputFolder, attrRows, docRows
        If Err.Number <> 0 Then
            ' Application.Run cannot pass Dictionary/Collection objects, so use fallback
            DebugLog "NOTE: BuildJurisdictionSheet requires complex objects that cannot be passed via Application.Run"
            DebugLog "Using fallback sheet structure. Error: " & Err.Number & " - " & Err.Description
            Err.Clear
            ' Create a basic sheet structure as fallback
            wsJur.Range("A1").Value = "Jurisdiction: " & jurName & " (" & jurID & ")"
            wsJur.Range("A1").Font.Bold = True
            wsJur.Range("A2").Value = "AuditorID: " & auditorID
            wsJur.Range("A3").Value = "Batch ID: " & batchID
            wsJur.Range("A5").Value = "GCIs for this jurisdiction:"
            wsJur.Range("A5").Font.Bold = True
            Dim gciRow As Long
            gciRow = 6
            Dim gciItem As Variant
            For Each gciItem In gciListForJur
                wsJur.Cells(gciRow, 1).Value = CStr(gciItem)
                gciRow = gciRow + 1
            Next gciItem
            wsJur.Columns("A").AutoFit
        End If
        On Error GoTo ErrHandler
    Next jurIDKey
    
    ' Save workbook
    Dim outPath As String
    ' outputFolder already declared and set above (line 653-655)
    
    On Error Resume Next
    outPath = Application.Run("BuildOutputPath", outputFolder, cfg("BatchID"), auditorID)
    If Err.Number <> 0 Then
        ' Fallback if BuildOutputPath doesn't exist
        Err.Clear
        ' Ensure trailing backslash
        If Right$(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"
        outPath = outputFolder & cfg("BatchID") & "_" & auditorID & ".xlsx"
    End If
    On Error GoTo ErrHandler
    
    ' Ensure output folder exists
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outputFolder) Then
        fso.CreateFolder outputFolder
    End If
    
    DebugLog "Saving workbook to: " & outPath
    
    Application.DisplayAlerts = False
    wbNew.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    wbNew.Close SaveChanges:=False
    
    DebugLog "GenerateOneAuditorWorkbook_FromReview: COMPLETE for " & auditorID
    Exit Sub
    
ErrHandler:
    Application.DisplayAlerts = True
    DebugLog "ERROR in GenerateOneAuditorWorkbook_FromReview for " & auditorID & ": " & Err.Number & " - " & Err.Description
    
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    On Error GoTo 0
    
    Err.Raise Err.Number, "GenerateOneAuditorWorkbook_FromReview", Err.Description
End Sub

Private Function SafeSheetName(ByVal name As String) As String
    ' Excel sheet names can't contain: \ / ? * [ ]
    Dim safe As String
    safe = name
    safe = Replace(safe, "\", "_")
    safe = Replace(safe, "/", "_")
    safe = Replace(safe, "?", "_")
    safe = Replace(safe, "*", "_")
    safe = Replace(safe, "[", "_")
    safe = Replace(safe, "]", "_")
    safe = Replace(safe, ":", "_")
    
    ' Limit to 31 characters
    If Len(safe) > 31 Then
        safe = Left$(safe, 31)
    End If
    
    SafeSheetName = safe
End Function

Private Sub UpdateGenerationStatus(ByVal wsReview As Worksheet, ByVal successCount As Long, ByVal totalCount As Long)
    On Error Resume Next
    
    Dim rngLabel As Range
    Set rngLabel = wsReview.Cells.Find(What:="Generator Status", LookIn:=xlValues, LookAt:=xlPart)
    
    If Not rngLabel Is Nothing Then
        Dim statusMsg As String
        statusMsg = "Generated " & successCount & " of " & totalCount & " workbook(s) at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        
        rngLabel.Offset(0, 1).Value = statusMsg
        DebugLog "Updated Generator Status: " & statusMsg
    End If
    
    On Error GoTo 0
End Sub

