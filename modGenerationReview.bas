Attribute VB_Name = "modGenerationReview"
Option Explicit

'=============================
' Module: modGenerationReview
' Purpose: Two-step workflow for generation review and auditor workbook generation
' Dependencies: Requires patched VBA Script_MatrixLayout module with Public helper functions
'=============================

'=============================
' CONSTANTS
'=============================
Private Const REVIEW_SHEET_NAME As String = "Generation Review"
Private Const REVIEW_TABLE_NAME As String = "tblGenerationReview"
Private Const SOURCE_TABLE_NAME As String = "tblEntities"

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
    On Error GoTo 0
End Sub

'=============================
' PUBLIC INTERFACE
'=============================

Public Sub PopulateGenerationReview(Optional ByVal ForceRefresh As Boolean = False)
    On Error GoTo ErrHandler

    DebugLog "=== PopulateGenerationReview START ==="
    DebugLog "ForceRefresh: " & ForceRefresh

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
        missingCols = Left$(missingCols, Len(missingCols) - 2) ' trim trailing comma
        MsgBox "Source table '" & SOURCE_TABLE_NAME & "' is missing required columns: " & missingCols, vbExclamation
        Exit Sub
    End If

    ' 4. Read source data
    Dim srcData As Variant
    srcData = loSource.DataBodyRange.Value

    DebugLog "Source data loaded: " & UBound(srcData, 1) & " rows"

    ' 5. Build existing data map (for preserving user edits)
    Dim existingMap As Object
    Set existingMap = BuildExistingReviewMap(loTarget, ForceRefresh)

    ' 6. Build new review data
    Dim reviewData() As Variant
    Dim rowCount As Long

    rowCount = UBound(srcData, 1)
    ReDim reviewData(1 To rowCount, 1 To 6) ' 6 columns: GCI, Legal Name, Jurisdiction ID, Jurisdiction, AuditorID, Auditor Name

    Dim i As Long, j As Long
    For i = 1 To rowCount
        Dim gci As String
        gci = CStr(GetCellValue(srcData, i, idxSrc, "GCI"))

        ' If preserving edits and this GCI exists, use existing row
        If Not ForceRefresh And existingMap.Exists(gci) Then
            Dim existingRow As Variant
            existingRow = existingMap(gci)
            For j = 1 To 6
                reviewData(i, j) = existingRow(j - 1)
            Next j
        Else
            ' Build new row from source
            reviewData(i, 1) = gci ' GCI
            reviewData(i, 2) = GetCellValue(srcData, i, idxSrc, "Legal Entity Name", GetCellValue(srcData, i, idxSrc, "Legal Name", "")) ' Legal Name
            reviewData(i, 3) = GetCellValue(srcData, i, idxSrc, "Jurisdiction ID") ' Jurisdiction ID
            reviewData(i, 4) = GetCellValue(srcData, i, idxSrc, "Jurisdiction Name", GetCellValue(srcData, i, idxSrc, "Jurisdiction", "")) ' Jurisdiction
            reviewData(i, 5) = GetCellValue(srcData, i, idxSrc, "AuditorID") ' AuditorID
            reviewData(i, 6) = GetCellValue(srcData, i, idxSrc, "Auditor Name", "") ' Auditor Name (optional)
        End If
    Next i

    ' 7. Write to target table
    WriteToReviewTable loTarget, reviewData

    ' 8. Update metadata cells
    UpdateReviewMetadata wsReview, rowCount, ForceRefresh

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
    Dim cfg As GenConfig
    cfg = ReadGeneratorConfig()
    Call ValidateConfig(cfg)

    DebugLog "Config validated: BatchID=" & cfg.BatchID

    ' 3. Read review data
    Dim reviewData As Variant
    reviewData = loReview.DataBodyRange.Value

    Dim idxReview As Object
    Set idxReview = BuildIndexMap(loReview)

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

    ' 4. Build auditor map from review table (same structure as original generator)
    Dim auditorMap As Object
    Set auditorMap = BuildAuditorMapFromReview(reviewData, idxReview)

    If auditorMap.Count = 0 Then
        MsgBox "No valid assignments found in review table." & vbCrLf & vbCrLf & _
               "Check that AuditorID, Jurisdiction ID, and GCI are populated for at least one row.", _
               vbExclamation
        Exit Sub
    End If

    DebugLog "AuditorMap built from review table with " & auditorMap.Count & " auditor(s)"

    ' 5. Load attributes and docs tables (needed for matrix and DV lists)
    Dim loAttr As ListObject, loDocs As ListObject
    Set loAttr = GetTableSafe(ThisWorkbook, "tblAttributes")
    Set loDocs = GetTableSafe(ThisWorkbook, "tblAcceptableDocs")

    If loAttr Is Nothing Or loDocs Is Nothing Then
        MsgBox "Required tables 'tblAttributes' and/or 'tblAcceptableDocs' not found.", vbExclamation
        Exit Sub
    End If

    Dim attrRows As Variant, docRows As Variant
    attrRows = loAttr.DataBodyRange.Value
    docRows = loDocs.DataBodyRange.Value

    Dim idxAttr As Object, idxDocs As Object
    Set idxAttr = BuildIndexMap(loAttr)
    Set idxDocs = BuildIndexMap(loDocs)

    DebugLog "Attributes and AcceptableDocs tables loaded"

    ' 6. Create dummy entRows array (not used in review workflow but required by signature)
    Dim entRows As Variant
    ReDim entRows(1 To 1, 1 To 1)

    Dim idxEnt As Object
    Set idxEnt = CreateObject("Scripting.Dictionary")

    ' 7. Generate workbooks for each auditor
    Dim auditorID As Variant
    Dim successCount As Long
    successCount = 0

    For Each auditorID In auditorMap.Keys
        DebugLog "Generating workbook for AuditorID: " & auditorID

        On Error Resume Next ' Allow individual auditor failures
        GenerateOneAuditorWorkbook_FromReview CStr(auditorID), cfg, auditorMap, _
                                              entRows, idxEnt, attrRows, idxAttr, docRows, idxDocs

        If Err.Number = 0 Then
            successCount = successCount + 1
        Else
            DebugLog "ERROR generating for " & auditorID & ": " & Err.Description
            MsgBox "Warning: Failed to generate workbook for AuditorID " & auditorID & vbCrLf & _
                   "Error: " & Err.Description, vbExclamation
        End If
        Err.Clear
        On Error GoTo ErrHandler
    Next auditorID

    ' 8. Update status
    UpdateGenerationStatus wsReview, successCount, auditorMap.Count

    DebugLog "=== GenerateAuditorWorkbooks_FromReview COMPLETE ==="
    MsgBox "Successfully generated " & successCount & " of " & auditorMap.Count & " workbook(s).", vbInformation

    Exit Sub

ErrHandler:
    DebugLog "ERROR in GenerateAuditorWorkbooks_FromReview: " & Err.Number & " - " & Err.Description
    MsgBox "GenerateAuditorWorkbooks_FromReview failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'=============================
' WORKBOOK GENERATION (delegates to MatrixLayout)
'=============================

Private Sub GenerateOneAuditorWorkbook_FromReview(ByVal auditorID As String, _
                                                   ByRef cfg As GenConfig, _
                                                   ByVal auditorMap As Object, _
                                                   ByVal entRows As Variant, ByVal idxEnt As Object, _
                                                   ByVal attrRows As Variant, ByVal idxAttr As Object, _
                                                   ByVal docRows As Variant, ByVal idxDocs As Object)
    On Error GoTo ErrHandler

    DebugLog "GenerateOneAuditorWorkbook_FromReview: START for " & auditorID

    Dim wbNew As Workbook
    Set wbNew = Application.Workbooks.Add(xlWBATWorksheet)

    ' Create Index sheet
    Dim wsIndex As Worksheet
    Set wsIndex = wbNew.Worksheets(1)
    wsIndex.Name = "Index"

    ' Call public function from MatrixLayout module
    BuildIndexHeader wsIndex, auditorID, cfg

    ' Create hidden lists sheet for DV
    Dim wsLists As Worksheet
    Set wsLists = wbNew.Worksheets.Add(After:=wbNew.Worksheets(wbNew.Worksheets.Count))
    wsLists.Name = "_Lists"
    wsLists.Visible = xlSheetVeryHidden

    ' Get assignments for this auditor
    Dim assigns As Collection
    Set assigns = auditorMap(auditorID)

    DebugLog "Auditor " & auditorID & " has " & assigns.Count & " assignment(s)"

    ' Create ENT sheet
    Dim wsENT As Worksheet
    Set wsENT = wbNew.Worksheets.Add(After:=wbNew.Worksheets(wbNew.Worksheets.Count))
    wsENT.Name = "ENT"

    ' Call public function from MatrixLayout module
    BuildJurisdictionSheet wsENT, wsLists, "ENT", "ENT", auditorID, cfg, assigns, _
                          entRows, idxEnt, attrRows, idxAttr, docRows, idxDocs

    ' Create jurisdiction sheets (one per unique jurisdiction in assignments)
    Dim jurMap As Object
    Set jurMap = BuildJurisdictionMapFromAssignments(assigns)

    DebugLog "Found " & jurMap.Count & " non-ENT jurisdiction(s)"

    Dim jurID As Variant
    For Each jurID In jurMap.Keys
        Dim jurName As String
        jurName = CStr(jurMap(jurID))

        Dim ws As Worksheet
        Set ws = wbNew.Worksheets.Add(After:=wbNew.Worksheets(wbNew.Worksheets.Count))
        ws.Name = SafeSheetName(jurName)

        ' Call public function from MatrixLayout module
        BuildJurisdictionSheet ws, wsLists, CStr(jurID), jurName, auditorID, cfg, assigns, _
                              entRows, idxEnt, attrRows, idxAttr, docRows, idxDocs
    Next jurID

    ' Save workbook
    Dim outPath As String
    outPath = BuildOutputPath(cfg.OutputFolder, cfg.BatchID, auditorID)

    DebugLog "Saving workbook to: " & outPath

    Application.DisplayAlerts = False
    wbNew.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook ' 51 = .xlsx format
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

'=============================
' REVIEW TABLE HELPERS
'=============================

Private Function BuildExistingReviewMap(ByVal loTarget As ListObject, ByVal forceRefresh As Boolean) As Object
    ' Returns a Dictionary: key=GCI, value=Array(col1, col2, ..., col6)
    ' Used to preserve user edits when ForceRefresh=False

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If forceRefresh Then
        DebugLog "ForceRefresh=True: Not preserving existing data"
        Set BuildExistingReviewMap = dict
        Exit Function
    End If

    If loTarget.DataBodyRange Is Nothing Then
        DebugLog "Review table is empty: Nothing to preserve"
        Set BuildExistingReviewMap = dict
        Exit Function
    End If

    Dim existingData As Variant
    existingData = loTarget.DataBodyRange.Value

    Dim idxTarget As Object
    Set idxTarget = BuildIndexMap(loTarget)

    If Not idxTarget.Exists("GCI") Then
        DebugLog "WARNING: Review table has no GCI column - cannot preserve data"
        Set BuildExistingReviewMap = dict
        Exit Function
    End If

    Dim i As Long
    For i = 1 To UBound(existingData, 1)
        Dim gci As String
        gci = Trim$(CStr(GetCellValue(existingData, i, idxTarget, "GCI")))

        If Len(gci) > 0 Then
            Dim rowValues(0 To 5) As Variant
            rowValues(0) = gci
            rowValues(1) = GetCellValue(existingData, i, idxTarget, "Legal Name", "")
            rowValues(2) = GetCellValue(existingData, i, idxTarget, "Jurisdiction ID", "")
            rowValues(3) = GetCellValue(existingData, i, idxTarget, "Jurisdiction", "")
            rowValues(4) = GetCellValue(existingData, i, idxTarget, "AuditorID", "")
            rowValues(5) = GetCellValue(existingData, i, idxTarget, "Auditor Name", "")

            dict(gci) = rowValues
        End If
    Next i

    DebugLog "BuildExistingReviewMap: Preserved " & dict.Count & " existing row(s)"
    Set BuildExistingReviewMap = dict
End Function

Private Sub WriteToReviewTable(ByVal loTarget As ListObject, ByRef reviewData() As Variant)
    On Error GoTo ErrHandler

    ' Clear existing data rows (preserve header)
    On Error Resume Next
    If Not loTarget.DataBodyRange Is Nothing Then
        loTarget.DataBodyRange.Delete
    End If
    On Error GoTo ErrHandler

    Dim rowCount As Long, colCount As Long
    rowCount = UBound(reviewData, 1)
    colCount = UBound(reviewData, 2)

    DebugLog "WriteToReviewTable: Writing " & rowCount & " rows x " & colCount & " cols"

    ' Get the range just below the header
    Dim ws As Worksheet
    Set ws = loTarget.Parent

    Dim targetRange As Range
    Set targetRange = loTarget.HeaderRowRange.Offset(1, 0).Resize(rowCount, colCount)

    ' Write data in one shot
    targetRange.Value = reviewData

    ' Resize table to include new data
    On Error Resume Next
    loTarget.Resize loTarget.Range.Resize(rowCount + 1, colCount)
    On Error GoTo 0

    DebugLog "WriteToReviewTable: Success"
    Exit Sub

ErrHandler:
    DebugLog "ERROR in WriteToReviewTable: " & Err.Description
    Err.Raise Err.Number, "WriteToReviewTable", Err.Description
End Sub

Private Sub UpdateReviewMetadata(ByVal wsReview As Worksheet, ByVal rowCount As Long, ByVal forceRefresh As Boolean)
    On Error Resume Next

    ' Update "Last Refresh" cell
    Dim rngLabel As Range
    Set rngLabel = wsReview.Cells.Find(What:="Last Refresh", LookIn:=xlValues, LookAt:=xlPart)

    If Not rngLabel Is Nothing Then
        Dim refreshMsg As String
        refreshMsg = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        If forceRefresh Then refreshMsg = refreshMsg & " (Force)"

        rngLabel.Offset(0, 1).Value = refreshMsg
        DebugLog "Updated Last Refresh: " & refreshMsg
    Else
        DebugLog "WARNING: 'Last Refresh' label not found on review sheet"
    End If

    ' Update "Generator Status" cell
    Set rngLabel = wsReview.Cells.Find(What:="Generator Status", LookIn:=xlValues, LookAt:=xlPart)

    If Not rngLabel Is Nothing Then
        rngLabel.Offset(0, 1).Value = "Ready (review " & rowCount & " assignment(s) below)"
        DebugLog "Updated Generator Status: Ready"
    Else
        DebugLog "WARNING: 'Generator Status' label not found on review sheet"
    End If

    On Error GoTo 0
End Sub

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

'=============================
' AUDITOR MAP BUILDING (from review table)
'=============================

Private Function BuildAuditorMapFromReview(ByVal reviewData As Variant, ByVal idxReview As Object) As Object
    ' Build same structure as BuildAuditorMap but from review table
    ' Returns: Dictionary where key=AuditorID, value=Collection of assignment records (Dictionaries)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long
    Dim skippedCount As Long
    skippedCount = 0

    For r = 1 To UBound(reviewData, 1)
        Dim auditorID As String
        auditorID = Trim$(CStr(GetCellValue(reviewData, r, idxReview, "AuditorID")))

        Dim gci As String
        gci = Trim$(CStr(GetCellValue(reviewData, r, idxReview, "GCI")))

        Dim jurID As String
        jurID = Trim$(CStr(GetCellValue(reviewData, r, idxReview, "Jurisdiction ID")))

        ' Skip rows with missing key fields
        If Len(auditorID) = 0 Or Len(gci) = 0 Or Len(jurID) = 0 Then
            skippedCount = skippedCount + 1
            GoTo ContinueRow
        End If

        ' Add auditor to map if not exists
        If Not dict.Exists(auditorID) Then
            dict(auditorID) = New Collection
        End If

        ' Build assignment record (Dictionary matching original structure)
        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        rec.CompareMode = vbTextCompare

        rec("AuditorID") = auditorID
        rec("GCI") = gci
        rec("Jurisdiction ID") = jurID
        rec("Jurisdiction Name") = CStr(GetCellValue(reviewData, r, idxReview, "Jurisdiction", jurID))
        rec("Legal Entity Name") = CStr(GetCellValue(reviewData, r, idxReview, "Legal Name", ""))
        rec("Auditor Name") = CStr(GetCellValue(reviewData, r, idxReview, "Auditor Name", ""))

        ' Add optional fields (code expects these, but review table may not have them)
        rec("Party Type") = ""
        rec("Onboarding Date") = ""
        rec("IRR") = ""
        rec("DRR") = ""
        rec("Primary FLU") = ""
        rec("Case ID # (Aware)") = ""

        dict(auditorID).Add rec

ContinueRow:
    Next r

    If skippedCount > 0 Then
        DebugLog "BuildAuditorMapFromReview: Skipped " & skippedCount & " row(s) with missing key fields"
    End If

    DebugLog "BuildAuditorMapFromReview: Built map with " & dict.Count & " auditor(s)"
    Set BuildAuditorMapFromReview = dict
End Function

'=============================
' HELPER FUNCTIONS (local and delegated)
'=============================

Private Function BuildIndexMap(ByVal lo As ListObject) As Object
    ' Build dictionary mapping column names to indexes (1-based)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        dict(lo.ListColumns(i).Name) = i
    Next i

    Set BuildIndexMap = dict
End Function

Private Function GetTableSafe(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    ' Find table across all worksheets, return Nothing if not found
    On Error Resume Next

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim lo As ListObject
        Set lo = ws.ListObjects(tableName)
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
        GetCellValue = rows(r, idx(colName))
    Else
        GetCellValue = defaultValue
    End If
End Function

Private Function GetNamedRangeValue(ByVal wb As Workbook, ByVal rangeName As String) As Variant
    ' Read named range value, handle multi-cell ranges
    On Error GoTo ErrHandler

    Dim rng As Range
    Set rng = wb.Names(rangeName).RefersToRange

    If rng.Cells.Count = 1 Then
        GetNamedRangeValue = rng.Value
    Else
        DebugLog "WARNING: Named range '" & rangeName & "' contains multiple cells. Using first cell only."
        GetNamedRangeValue = rng.Cells(1, 1).Value
    End If
    Exit Function

ErrHandler:
    DebugLog "ERROR: GetNamedRangeValue failed for '" & rangeName & "' - " & Err.Description
    Err.Raise vbObjectError + 2101, "GetNamedRangeValue", "Named range not found or invalid: " & rangeName
End Function

Private Function BuildOutputPath(ByVal outFolder As String, ByVal batchID As String, ByVal auditorID As String) As String
    BuildOutputPath = EnsureTrailingSlash(outFolder) & "KYC_Workpapers_" & CleanFilePart(batchID) & "_" & CleanFilePart(auditorID) & ".xlsx"
End Function

Private Function EnsureTrailingSlash(ByVal path As String) As String
    If Len(path) = 0 Then
        EnsureTrailingSlash = path
    ElseIf Right$(path, 1) = "\" Or Right$(path, 1) = "/" Then
        EnsureTrailingSlash = path
    Else
        EnsureTrailingSlash = path & "\"
    End If
End Function

Private Function CleanFilePart(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "_")
    Next i
    CleanFilePart = Trim$(s)
End Function

Private Function SafeSheetName(ByVal s As String) As String
    Dim bad As Variant
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), " ")
    Next i
    s = Application.WorksheetFunction.Trim(s)
    If Len(s) = 0 Then s = "Sheet"
    If Len(s) > 31 Then s = Left$(s, 31)
    SafeSheetName = s
End Function

Private Function BuildJurisdictionMapFromAssignments(ByVal assigns As Collection) As Object
    ' Extract unique jurisdictions from assignment records (excluding ENT)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To assigns.Count
        Dim rec As Object
        Set rec = assigns(i)

        Dim jurID As String
        jurID = CStr(rec("Jurisdiction ID"))

        Dim jurName As String
        jurName = CStr(rec("Jurisdiction Name"))

        If UCase$(Trim$(jurID)) <> "ENT" Then
            If Not dict.Exists(jurID) Then
                dict(jurID) = jurName
            End If
        End If
    Next i

    Set BuildJurisdictionMapFromAssignments = dict
End Function

'=============================
' CONFIG READING (uses same signature as MatrixLayout)
'=============================

Private Function ReadGeneratorConfig() As GenConfig
    Dim cfg As GenConfig

    cfg.BatchID = CStr(GetNamedRangeValue(ThisWorkbook, "rngBatchID"))
    cfg.OutputFolder = CStr(GetNamedRangeValue(ThisWorkbook, "rngOutputFolder"))

    DebugLog "Config read: BatchID=" & cfg.BatchID & ", OutputFolder=" & cfg.OutputFolder

    ReadGeneratorConfig = cfg
End Function

Private Sub ValidateConfig(ByRef cfg As GenConfig)
    If Len(Trim$(cfg.BatchID)) = 0 Then
        DebugLog "ERROR: BatchID is empty"
        Err.Raise vbObjectError + 2000, "ValidateConfig", "BatchID is required."
    End If

    If Len(Trim$(cfg.OutputFolder)) = 0 Then
        DebugLog "ERROR: OutputFolder is empty"
        Err.Raise vbObjectError + 2001, "ValidateConfig", "OutputFolder is required."
    End If

    DebugLog "Config validated successfully"
End Sub
