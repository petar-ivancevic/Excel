Option Explicit

'=============================
' Generator configuration
'=============================
Public Type GenConfig
    BatchID As String
    OutputFolder As String
End Type

'=============================
' DEBUG LOGGING
'=============================
Private Sub DebugLog(ByVal msg As String)
    ' Logs to Immediate Window
    Debug.Print Format$(Now, "hh:nn:ss") & " | " & msg

    ' Optional: Write to _Log sheet
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
' Entry point
'=============================
Public Sub GenerateWorkpapers()
    On Error GoTo ErrHandler

    DebugLog "=== GenerateWorkpapers START ==="

    Dim cfg As GenConfig
    cfg = ReadGeneratorConfig()
    ValidateConfig cfg

    Dim wbSrc As Workbook
    Set wbSrc = ThisWorkbook

    ' Source tables
    Dim loEnt As ListObject, loAttr As ListObject, loDocs As ListObject
    Set loEnt = GetTable(wbSrc, "tblEntities")
    Set loAttr = GetTable(wbSrc, "tblAttributes")
    Set loDocs = GetTable(wbSrc, "tblAcceptableDocs")

    DebugLog "Tables loaded: tblEntities, tblAttributes, tblAcceptableDocs"

    ' Pull data into arrays
    Dim entRows As Variant, attrRows As Variant, docRows As Variant
    entRows = loEnt.DataBodyRange.Value
    attrRows = loAttr.DataBodyRange.Value
    docRows = loDocs.DataBodyRange.Value

    DebugLog "Data arrays populated"

    ' Index maps
    Dim idxEnt As Object, idxAttr As Object, idxDocs As Object
    Set idxEnt = BuildIndexMap(loEnt)
    Set idxAttr = BuildIndexMap(loAttr)
    Set idxDocs = BuildIndexMap(loDocs)

    ' auditor -> collection of GCI assignments
    Dim auditorMap As Object
    Set auditorMap = BuildAuditorMap(entRows, idxEnt)

    DebugLog "AuditorMap built with " & auditorMap.Count & " auditors"

    Dim auditorID As Variant
    For Each auditorID In auditorMap.Keys
        DebugLog "Generating workbook for AuditorID: " & auditorID
        GenerateOneAuditorWorkbook CStr(auditorID), cfg, auditorMap, entRows, idxEnt, attrRows, idxAttr, docRows, idxDocs
    Next auditorID

    DebugLog "=== GenerateWorkpapers COMPLETE ==="
    MsgBox "Generation complete.", vbInformation
    Exit Sub

ErrHandler:
    DebugLog "ERROR in GenerateWorkpapers: " & Err.Number & " - " & Err.Description
    MsgBox "GenerateWorkpapers failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

'=============================
' Config
'=============================
Private Function ReadGeneratorConfig() As GenConfig
    Dim cfg As GenConfig

    cfg.BatchID = CStr(GetNamedRangeValue(ThisWorkbook, "rngBatchID"))
    cfg.OutputFolder = CStr(GetNamedRangeValue(ThisWorkbook, "rngOutputFolder"))

    DebugLog "Config read: BatchID=" & cfg.BatchID & ", OutputFolder=" & cfg.OutputFolder

    ReadGeneratorConfig = cfg
End Function

' FIX: UDT cannot be passed ByVal
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

'=============================
' Core generation
'=============================
' FIX: UDT cannot be passed ByVal
Private Sub GenerateOneAuditorWorkbook(ByVal auditorID As String, _
                                      ByRef cfg As GenConfig, _
                                      ByVal auditorMap As Object, _
                                      ByVal entRows As Variant, ByVal idxEnt As Object, _
                                      ByVal attrRows As Variant, ByVal idxAttr As Object, _
                                      ByVal docRows As Variant, ByVal idxDocs As Object)

    On Error GoTo ErrHandler

    DebugLog "GenerateOneAuditorWorkbook: START for " & auditorID

    Dim wbNew As Workbook
    Set wbNew = Application.Workbooks.Add(xlWBATWorksheet)

    ' Create Index first
    Dim wsIndex As Worksheet
    Set wsIndex = wbNew.Worksheets(1)
    wsIndex.Name = "Index"
    BuildIndexHeader wsIndex, auditorID, cfg

    ' Create hidden lists sheet (for DV lists)
    Dim wsLists As Worksheet
    Set wsLists = wbNew.Worksheets.Add(After:=wbNew.Worksheets(wbNew.Worksheets.Count))
    wsLists.Name = "_Lists"
    wsLists.Visible = xlSheetVeryHidden

    ' Assignments for this auditor
    Dim assigns As Collection
    Set assigns = auditorMap(auditorID)

    DebugLog "Auditor " & auditorID & " has " & assigns.Count & " assignments"

    ' ENT sheet
    Dim wsENT As Worksheet
    Set wsENT = wbNew.Worksheets.Add(After:=wbNew.Worksheets(wbNew.Worksheets.Count))
    wsENT.Name = "ENT"
    BuildJurisdictionSheet wsENT, wsLists, "ENT", "ENT", auditorID, cfg, assigns, entRows, idxEnt, attrRows, idxAttr, docRows, idxDocs

    ' One sheet per jurisdiction (excluding ENT)
    Dim jurMap As Object
    Set jurMap = BuildJurisdictionMapFromAssignments(assigns)

    DebugLog "Found " & jurMap.Count & " non-ENT jurisdictions"

    Dim jurID As Variant
    For Each jurID In jurMap.Keys
        Dim jurName As String
        jurName = CStr(jurMap(jurID))

        Dim ws As Worksheet
        Set ws = wbNew.Worksheets.Add(After:=wbNew.Worksheets(wbNew.Worksheets.Count))
        ws.Name = SafeSheetName(jurName)

        BuildJurisdictionSheet ws, wsLists, CStr(jurID), jurName, auditorID, cfg, assigns, entRows, idxEnt, attrRows, idxAttr, docRows, idxDocs
    Next jurID

    ' Save workbook
    Dim outPath As String
    outPath = BuildOutputPath(cfg.OutputFolder, cfg.BatchID, auditorID)

    DebugLog "Saving workbook to: " & outPath

    Application.DisplayAlerts = False
    ' FIX: Changed to xlOpenXMLWorkbook (51) to match .xlsx extension
    wbNew.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    wbNew.Close SaveChanges:=False

    DebugLog "GenerateOneAuditorWorkbook: COMPLETE for " & auditorID
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    DebugLog "ERROR in GenerateOneAuditorWorkbook for " & auditorID & ": " & Err.Number & " - " & Err.Description
    MsgBox "GenerateOneAuditorWorkbook failed for AuditorID=" & auditorID & ": " & Err.Number & " - " & Err.Description, vbExclamation
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
End Sub

' FIX: UDT cannot be passed ByVal
Private Sub BuildJurisdictionSheet(ByVal ws As Worksheet, ByVal wsLists As Worksheet, _
                                  ByVal jurID As String, ByVal jurName As String, _
                                  ByVal auditorID As String, ByRef cfg As GenConfig, _
                                  ByVal assigns As Collection, _
                                  ByVal entRows As Variant, ByVal idxEnt As Object, _
                                  ByVal attrRows As Variant, ByVal idxAttr As Object, _
                                  ByVal docRows As Variant, ByVal idxDocs As Object)

    On Error GoTo ErrHandler

    DebugLog "BuildJurisdictionSheet: " & jurID & " (" & jurName & ")"

    ws.Cells.Clear

    ' Build list of GCIs for this auditor AND jurisdiction
    Dim gciList As Collection
    Set gciList = New Collection

    Dim i As Long
    For i = 1 To assigns.Count
        Dim a As Object
        Set a = assigns(i)
        If UCase$(Trim$(CStr(a("Jurisdiction ID")))) = UCase$(Trim$(jurID)) Then
            gciList.Add CStr(a("GCI"))
        End If
    Next i

    ' If ENT: include all GCIs assigned to auditor
    If UCase$(Trim$(jurID)) = "ENT" Then
        Set gciList = New Collection
        For i = 1 To assigns.Count
            Set a = assigns(i)
            On Error Resume Next
            gciList.Add CStr(a("GCI"))
            On Error GoTo 0
        Next i
    End If

    DebugLog "  GCIs for this sheet: " & gciList.Count

    ' Filter attributes for this jurisdiction
    Dim attrIdxs As Collection
    Set attrIdxs = FilterAttributeRowIndexes(attrRows, idxAttr, jurID)

    DebugLog "  Attributes for this jurisdiction: " & attrIdxs.Count

    ' Build DV lists for each Attribute ID (acceptable docs)
    Dim dvMap As Object
    Set dvMap = BuildDvLists(wsLists, attrRows, idxAttr, docRows, idxDocs)

    ' Render header rows
    RenderTopHeader ws, auditorID, jurID, jurName, gciList, assigns, cfg, entRows, idxEnt

    ' Render attribute/question rows & matrix columns
    RenderMatrix ws, jurID, gciList, attrRows, idxAttr, attrIdxs, dvMap

    ' Finishing touches
    ApplySheetLayout ws

    Exit Sub

ErrHandler:
    DebugLog "ERROR in BuildJurisdictionSheet: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "BuildJurisdictionSheet", Err.Description
End Sub

' FIX: UDT cannot be passed ByVal
Private Sub BuildIndexHeader(ByVal ws As Worksheet, ByVal auditorID As String, ByRef cfg As GenConfig)
    ws.Cells.Clear

    ws.Range("A1").Value = "KYC Workpapers (Matrix Layout)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A3").Value = "AuditorID:"
    ws.Range("B3").Value = auditorID

    ws.Range("A4").Value = "Batch ID:"
    ws.Range("B4").Value = cfg.BatchID

    ws.Range("A5").Value = "Generated:"
    ws.Range("B5").Value = Now

    ws.Columns("A:B").AutoFit
End Sub

'=============================
' Rendering helpers
'=============================
Private Sub RenderTopHeader(ByVal ws As Worksheet, _
    ByVal auditorID As String, _
    ByVal jurID As String, _
    ByVal jurName As String, _
    ByVal gciList As Collection, _
    ByVal assigns As Collection, _
    ByRef cfg As GenConfig, _
    ByVal entRows As Variant, ByVal idxEnt As Object)

    ws.Range("A1").Value = "Auditor"
    ws.Range("B1").Value = auditorID

    ws.Range("A2").Value = "Jurisdiction"
    ws.Range("B2").Value = jurName & " (" & jurID & ")"

    ws.Range("A3").Value = "Batch"
    ws.Range("B3").Value = cfg.BatchID

    Dim startCol As Long: startCol = 8
    Dim c As Long
    For c = 1 To gciList.Count
        ws.Cells(5, startCol + (c - 1)).Value = gciList(c)
        ws.Cells(5, startCol + (c - 1)).Font.Bold = True
    Next c
End Sub

Private Sub RenderMatrix(ByVal ws As Worksheet, _
    ByVal jurID As String, _
    ByVal gciList As Collection, _
    ByVal attrRows As Variant, ByVal idxAttr As Object, _
    ByVal attrIdxs As Collection, _
    ByVal dvMap As Object)

    Dim rStart As Long: rStart = 7
    Dim startCol As Long: startCol = 8

    ws.Cells(rStart, 1).Value = "Source File"
    ws.Cells(rStart, 2).Value = "Attribute ID"
    ws.Cells(rStart, 3).Value = "Attribute Name"
    ws.Cells(rStart, 4).Value = "Category"
    ws.Cells(rStart, 5).Value = "Source"
    ws.Cells(rStart, 6).Value = "Source Page"
    ws.Cells(rStart, 7).Value = "Question Text"

    Dim i As Long
    For i = 1 To attrIdxs.Count
        Dim rr As Long: rr = rStart + i
        Dim aRow As Long: aRow = attrIdxs(i)

        ws.Cells(rr, 1).Value = GetIfColumn(attrRows, aRow, idxAttr, "Source File")
        ws.Cells(rr, 2).Value = GetIfColumn(attrRows, aRow, idxAttr, "Attribute ID")
        ws.Cells(rr, 3).Value = GetIfColumn(attrRows, aRow, idxAttr, "Attribute Name")
        ws.Cells(rr, 4).Value = GetIfColumn(attrRows, aRow, idxAttr, "Category")
        ws.Cells(rr, 5).Value = GetIfColumn(attrRows, aRow, idxAttr, "Source")
        ws.Cells(rr, 6).Value = GetIfColumn(attrRows, aRow, idxAttr, "Source Page")
        ws.Cells(rr, 7).Value = GetIfColumn(attrRows, aRow, idxAttr, "Question Text")

        Dim attrID As String
        attrID = CStr(GetIfColumn(attrRows, aRow, idxAttr, "Attribute ID"))

        Dim c As Long
        For c = 1 To gciList.Count
            Dim cell As Range
            Set cell = ws.Cells(rr, startCol + (c - 1))
            ApplyDvToCell cell, dvMap, attrID
        Next c
    Next i
End Sub

Private Sub ApplySheetLayout(ByVal ws As Worksheet)
    ws.Rows(1).Font.Bold = True
    ws.Rows(2).Font.Bold = True
    ws.Rows(3).Font.Bold = True
    ws.Rows(7).Font.Bold = True

    ws.Columns("A:F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 60
End Sub

'=============================
' DV list building
'=============================
Private Function BuildDvLists(ByVal wsLists As Worksheet, _
    ByVal attrRows As Variant, ByVal idxAttr As Object, _
    ByVal docRows As Variant, ByVal idxDocs As Object) As Object

    Dim dvMap As Object
    Set dvMap = CreateObject("Scripting.Dictionary")
    dvMap.CompareMode = vbTextCompare

    wsLists.Cells.Clear

    Dim nextRow As Long: nextRow = 1

    Dim r As Long
    For r = 1 To UBound(attrRows, 1)
        Dim attrID As String
        attrID = CStr(GetIfColumn(attrRows, r, idxAttr, "Attribute ID"))
        If Len(Trim$(attrID)) = 0 Then GoTo ContinueAttr

        If Not dvMap.Exists(attrID) Then
            Dim startRow As Long: startRow = nextRow

            Dim d As Long
            For d = 1 To UBound(docRows, 1)
                If CStr(GetIfColumn(docRows, d, idxDocs, "Attribute ID")) = attrID Then
                    wsLists.Cells(nextRow, 1).Value = CStr(GetIfColumn(docRows, d, idxDocs, "Evidence Source/Document"))
                    nextRow = nextRow + 1
                End If
            Next d

            wsLists.Cells(nextRow, 1).Value = "Fail 1": nextRow = nextRow + 1
            wsLists.Cells(nextRow, 1).Value = "Fail 2": nextRow = nextRow + 1

            Dim addr As String
            addr = wsLists.Range(wsLists.Cells(startRow, 1), wsLists.Cells(nextRow - 1, 1)).Address(True, True, xlA1, True)
            dvMap(attrID) = addr
        End If

ContinueAttr:
    Next r

    Set BuildDvLists = dvMap
End Function

Private Sub ApplyDvToCell(ByVal cell As Range, ByVal dvMap As Object, ByVal attrID As String)
    On Error Resume Next
    cell.Validation.Delete
    On Error GoTo ErrHandler

    If dvMap.Exists(attrID) Then
        ' FIX: Removed Operator parameter for xlValidateList (Error 450)
        cell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="=" & dvMap(attrID)
    Else
        ' FIX: Removed Operator parameter for xlValidateList (Error 450)
        cell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="Fail 1,Fail 2"
    End If

    Exit Sub

ErrHandler:
    ' Log DV error but don't crash
    DebugLog "WARNING: Failed to apply validation to " & cell.Address & " for AttrID=" & attrID & " - " & Err.Description
    On Error GoTo 0
End Sub

'=============================
' Attribute filtering
'=============================
Private Function FilterAttributeRowIndexes(ByVal attrRows As Variant, ByVal idxAttr As Object, ByVal jurID As String) As Collection
    Dim col As New Collection

    Dim r As Long
    For r = 1 To UBound(attrRows, 1)
        Dim thisJur As String
        thisJur = CStr(GetIfColumn(attrRows, r, idxAttr, "Jurisdiction ID"))

        If UCase$(Trim$(thisJur)) = UCase$(Trim$(jurID)) Then
            If IsRequiredAttribute(attrRows, r, idxAttr) Then
                col.Add r
            End If
        End If
    Next r

    Set FilterAttributeRowIndexes = col
End Function

Private Function IsRequiredAttribute(ByVal attrRows As Variant, ByVal r As Long, ByVal idxAttr As Object) As Boolean
    If idxAttr.Exists("IsRequired") Then
        Dim v As String
        v = UCase$(Trim$(CStr(attrRows(r, idxAttr("IsRequired")))))
        If v = "" Or v = "Y" Or v = "YES" Then
            IsRequiredAttribute = True
        Else
            IsRequiredAttribute = False
        End If
    Else
        IsRequiredAttribute = True
    End If
End Function

'=============================
' Auditor / jurisdiction maps
'=============================
Private Function BuildAuditorMap(ByVal entRows As Variant, ByVal idxEnt As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long
    For r = 1 To UBound(entRows, 1)
        Dim auditorID As String
        auditorID = CStr(entRows(r, idxEnt("AuditorID")))
        If Len(Trim$(auditorID)) = 0 Then GoTo ContinueRow

        If Not dict.Exists(auditorID) Then
            dict(auditorID) = New Collection
        End If

        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        rec.CompareMode = vbTextCompare

        rec("AuditorID") = auditorID
        rec("GCI") = CStr(entRows(r, idxEnt("GCI")))
        rec("Jurisdiction ID") = CStr(entRows(r, idxEnt("Jurisdiction ID")))
        rec("Jurisdiction Name") = CStr(entRows(r, idxEnt("Jurisdiction Name")))

        AddIfExists rec, "Legal Entity Name", entRows, r, idxEnt
        AddIfExists rec, "Party Type", entRows, r, idxEnt
        AddIfExists rec, "Onboarding Date", entRows, r, idxEnt
        AddIfExists rec, "IRR", entRows, r, idxEnt
        AddIfExists rec, "DRR", entRows, r, idxEnt
        AddIfExists rec, "Primary FLU", entRows, r, idxEnt
        AddIfExists rec, "Case ID # (Aware)", entRows, r, idxEnt

        dict(auditorID).Add rec

ContinueRow:
    Next r

    Set BuildAuditorMap = dict
End Function

Private Function BuildJurisdictionMapFromAssignments(ByVal assigns As Collection) As Object
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
            If Not dict.Exists(jurID) Then dict(jurID) = jurName
        End If
    Next i

    Set BuildJurisdictionMapFromAssignments = dict
End Function

Private Sub AddIfExists(ByVal rec As Object, ByVal colName As String, ByVal rows As Variant, ByVal r As Long, ByVal idx As Object)
    If idx.Exists(colName) Then
        rec(colName) = rows(r, idx(colName))
    Else
        rec(colName) = ""
    End If
End Sub

'=============================
' Table / named range helpers
'=============================
Private Function BuildIndexMap(ByVal lo As ListObject) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        dict(lo.ListColumns(i).Name) = i
    Next i

    Set BuildIndexMap = dict
End Function

Private Function GetTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        On Error Resume Next
        Dim lo As ListObject
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set GetTable = lo
            Exit Function
        End If
    Next ws

    Err.Raise vbObjectError + 2100, "GetTable", "Table not found: " & tableName
End Function

Private Function GetNamedRangeValue(ByVal wb As Workbook, ByVal rangeName As String) As Variant
    On Error GoTo ErrHandler

    Dim rng As Range
    Set rng = wb.Names(rangeName).RefersToRange

    ' FIX: Check if range is single cell to avoid Error 13 (array vs single value)
    If rng.Cells.Count = 1 Then
        GetNamedRangeValue = rng.Value
    Else
        ' If multi-cell range, return top-left cell
        DebugLog "WARNING: Named range '" & rangeName & "' contains multiple cells. Using first cell only."
        GetNamedRangeValue = rng.Cells(1, 1).Value
    End If
    Exit Function

ErrHandler:
    DebugLog "ERROR: GetNamedRangeValue failed for '" & rangeName & "' - " & Err.Description
    Err.Raise vbObjectError + 2101, "GetNamedRangeValue", "Named range not found or invalid: " & rangeName
End Function

Private Function GetIfColumn(ByVal rows As Variant, ByVal r As Long, ByVal idx As Object, ByVal colName As String) As Variant
    If idx.Exists(colName) Then
        GetIfColumn = rows(r, idx(colName))
    Else
        GetIfColumn = ""
    End If
End Function

'=============================
' Output path / sanitizers
'=============================
Private Function BuildOutputPath(ByVal outFolder As String, ByVal batchID As String, ByVal auditorID As String) As String
    ' FIX: Changed extension to .xlsx to match FileFormat
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
