Option Explicit

'===========================
' Types
'===========================

Public Type GenConfig
    AuditorID As String
    BatchID As String
    OutputFolder As String
End Type

'===========================
' PUBLIC TEST HARNESS
'===========================

Public Sub Test_ConfigAndEntities()
    On Error GoTo ErrHandler
    
    Dim cfg As GenConfig
    Dim entities As Collection
    Dim i As Long
    
    cfg = ReadGeneratorConfig()
    Set entities = GetSelectedEntityIDs()
    
    Call ValidateGeneratorInputs(cfg, entities)
    
    Debug.Print "=== Config ==="
    Debug.Print "AuditorID:    "; cfg.AuditorID
    Debug.Print "BatchID:      "; cfg.BatchID
    Debug.Print "OutputFolder: "; cfg.OutputFolder
    
    Debug.Print "=== Selected Entities ==="
    Debug.Print "Count: "; entities.Count
    For i = 1 To entities.Count
        Debug.Print "  Entity[" & i & "]: "; entities(i)
    Next i
    
    Exit Sub
    
ErrHandler:
    MsgBox "Test_ConfigAndEntities failed: " & Err.Description, vbExclamation
End Sub

Public Sub Test_BuildResponsesArray()
    On Error GoTo ErrHandler
    
    Dim cfg As GenConfig
    Dim selectedEntities As Collection
    Dim responses As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim headers As Variant
    
    cfg = ReadGeneratorConfig()
    Set selectedEntities = GetSelectedEntityIDs()
    Call ValidateGeneratorInputs(cfg, selectedEntities)
    
    responses = BuildResponsesArray(cfg, selectedEntities)
    
    ' Remove old TestOutput if present
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("TestOutput").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "TestOutput"
    
    headers = Array( _
        "EntityID", _
        "EntityName", _
        "JurisdictionID", _
        "AttributeID", _
        "AttributeName", _
        "QuestionText", _
        "DocumentationAgeRule", _
        "ResponseStatus", _
        "DocType", _
        "DocName", _
        "DocDate", _
        "Comments", _
        "AuditorID", _
        "BatchID")
    
    For j = LBound(headers) To UBound(headers)
        ws.Cells(1, j + 1).Value = headers(j)
    Next j
    
    If Not IsEmpty(responses) Then
        For i = LBound(responses, 1) To UBound(responses, 1)
            For j = LBound(responses, 2) To UBound(responses, 2)
                ws.Cells(i + 1, j).Value = responses(i, j)
            Next j
        Next i
    End If
    
    ws.Columns.AutoFit
    MsgBox "Test_BuildResponsesArray completed. Check the 'TestOutput' sheet.", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Test_BuildResponsesArray failed: " & Err.Description, vbExclamation
End Sub

'===========================
' CONFIG + SELECTION
'===========================

Public Function ReadGeneratorConfig() As GenConfig
    On Error GoTo ErrHandler
    
    Dim cfg As GenConfig
    
    cfg.AuditorID = CStr(ThisWorkbook.Names("rngAuditorID").RefersToRange.Value)
    cfg.BatchID = CStr(ThisWorkbook.Names("rngBatchID").RefersToRange.Value)
    cfg.OutputFolder = CStr(ThisWorkbook.Names("rngOutputFolder").RefersToRange.Value)
    
    ReadGeneratorConfig = cfg
    Exit Function
    
ErrHandler:
    MsgBox "Error reading generator configuration: " & Err.Description, _
           vbCritical, "Configuration Error"
    Dim emptyCfg As GenConfig
    ReadGeneratorConfig = emptyCfg
End Function

Public Function GetSelectedEntityIDs() As Collection
    On Error GoTo ErrHandler
    
    Dim rng As Range
    Dim cell As Range
    Dim col As New Collection
    
    Set rng = ThisWorkbook.Names("rngSelectedEntities").RefersToRange
    
    For Each cell In rng.Cells
        If Len(Trim$(cell.Value)) > 0 Then
            col.Add CStr(cell.Value)
        End If
    Next cell
    
    Set GetSelectedEntityIDs = col
    Exit Function
    
ErrHandler:
    MsgBox "Error reading selected entities: " & Err.Description, _
           vbCritical, "Selection Error"
    Set GetSelectedEntityIDs = Nothing
End Function

Public Sub ValidateGeneratorInputs(cfg As GenConfig, entities As Collection)
    ' Raises if inputs are invalid
    
    If Len(Trim$(cfg.AuditorID)) = 0 Then
        Err.Raise vbObjectError + 1000, "ValidateGeneratorInputs", _
                  "Auditor ID is required."
    End If
    
    If Len(Trim$(cfg.OutputFolder)) = 0 Then
        Err.Raise vbObjectError + 1001, "ValidateGeneratorInputs", _
                  "Output folder is required."
    End If
    
    If Dir(cfg.OutputFolder, vbDirectory) = vbNullString Then
        Err.Raise vbObjectError + 1002, "ValidateGeneratorInputs", _
                  "Output folder does not exist: " & cfg.OutputFolder
    End If
    
    If entities Is Nothing Then
        Err.Raise vbObjectError + 1003, "ValidateGeneratorInputs", _
                  "Entities collection is Nothing."
    End If
    
    If entities.Count = 0 Then
        Err.Raise vbObjectError + 1004, "ValidateGeneratorInputs", _
                  "At least one EntityID must be selected."
    End If
End Sub

'===========================
' TABLE HELPERS
'===========================

Private Function GetTableColumnIndex(lo As ListObject, headerName As String) As Long
    Dim hdr As Range
    For Each hdr In lo.HeaderRowRange.Cells
        If StrComp(Trim$(hdr.Value), headerName, vbTextCompare) = 0 Then
            GetTableColumnIndex = hdr.Column - lo.HeaderRowRange.Columns(1).Column + 1
            Exit Function
        End If
    Next hdr
    Err.Raise vbObjectError + 1100, "GetTableColumnIndex", _
              "Column not found in table '" & lo.Name & "': " & headerName
End Function

Private Function LoadEntitiesData(ByRef dataOut As Variant, _
                                  ByRef idxGCI As Long, _
                                  ByRef idxLegalName As Long, _
                                  ByRef idxJurID As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim lo As ListObject
    
    Set ws = ThisWorkbook.Worksheets("Entities")
    Set lo = ws.ListObjects("tblEntities")
    
    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1200, "LoadEntitiesData", "tblEntities has no data."
    End If
    
    dataOut = lo.DataBodyRange.Value
    idxGCI = GetTableColumnIndex(lo, "GCI")
    idxLegalName = GetTableColumnIndex(lo, "Legal Name")
    idxJurID = GetTableColumnIndex(lo, "Jurisdiction ID")
    
    LoadEntitiesData = True
    Exit Function
    
ErrHandler:
    MsgBox "LoadEntitiesData failed: " & Err.Description, vbCritical
    LoadEntitiesData = False
End Function

Private Function LoadAttributesData(ByRef dataOut As Variant, _
                                    ByRef idxAttrID As Long, _
                                    ByRef idxAttrName As Long, _
                                    ByRef idxQuestionText As Long, _
                                    ByRef idxJurID As Long, _
                                    ByRef idxRiskScope As Long, _
                                    ByRef idxIsRequired As Long, _
                                    ByRef idxDocAge As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim lo As ListObject
    
    Set ws = ThisWorkbook.Worksheets("Attributes")
    Set lo = ws.ListObjects("tblAttributes")
    
    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1300, "LoadAttributesData", "tblAttributes has no data."
    End If
    
    dataOut = lo.DataBodyRange.Value
    
    idxAttrID = GetTableColumnIndex(lo, "Attribute ID")
    idxAttrName = GetTableColumnIndex(lo, "Attribute Name")
    idxQuestionText = GetTableColumnIndex(lo, "Question Text")
    idxJurID = GetTableColumnIndex(lo, "Jurisdiction ID")
    idxRiskScope = GetTableColumnIndex(lo, "RiskScope")
    idxIsRequired = GetTableColumnIndex(lo, "IsRequired")
    idxDocAge = GetTableColumnIndex(lo, "DocumentationAgeRule")
    
    LoadAttributesData = True
    Exit Function
    
ErrHandler:
    MsgBox "LoadAttributesData failed: " & Err.Description, vbCritical
    LoadAttributesData = False
End Function

Private Function LoadAcceptableDocsData(ByRef dataOut As Variant, _
                                        ByRef idxAttrID As Long, _
                                        ByRef idxJurID As Long, _
                                        ByRef idxDocName As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim lo As ListObject
    
    Set ws = ThisWorkbook.Worksheets("AcceptableDocs")
    Set lo = ws.ListObjects("tblAcceptableDocs")
    
    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1400, "LoadAcceptableDocsData", "tblAcceptableDocs has no data."
    End If
    
    dataOut = lo.DataBodyRange.Value
    
    idxAttrID = GetTableColumnIndex(lo, "Attribute ID")
    idxJurID = GetTableColumnIndex(lo, "Jurisdiction ID")
    idxDocName = GetTableColumnIndex(lo, "Evidence Source/Document")
    
    LoadAcceptableDocsData = True
    Exit Function
    
ErrHandler:
    MsgBox "LoadAcceptableDocsData failed: " & Err.Description, vbCritical
    LoadAcceptableDocsData = False
End Function

Private Function BuildDocTypeList(attrID As String, jurID As String, _
                                  accData As Variant, _
                                  idxAttrID As Long, _
                                  idxJurID As Long, _
                                  idxDocName As Long) As String
    Dim i As Long
    Dim aID As String
    Dim jID As String
    Dim docName As String
    Dim result As String
    
    ' Build a comma-separated list of doc types for this (AttributeID, JurisdictionID)
    For i = 1 To UBound(accData, 1)
        aID = CStr(accData(i, idxAttrID))
        jID = CStr(accData(i, idxJurID))
        
        If StrComp(aID, attrID, vbTextCompare) = 0 Then
            If (jID = jurID) Or (jID = "All") Or (Len(jID) = 0) Then
                docName = Trim$(CStr(accData(i, idxDocName)))
                If Len(docName) > 0 Then
                    If Len(result) = 0 Then
                        result = docName
                    Else
                        ' avoid simple duplicates
                        If InStr(1, "," & result & ",", "," & docName & ",", vbTextCompare) = 0 Then
                            result = result & "," & docName
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    BuildDocTypeList = result
End Function

Private Function FindEntityRow(entData As Variant, idxGCI As Long, entityID As String) As Long
    Dim i As Long
    For i = 1 To UBound(entData, 1)
        If StrComp(CStr(entData(i, idxGCI)), entityID, vbTextCompare) = 0 Then
            FindEntityRow = i
            Exit Function
        End If
    Next i
    FindEntityRow = 0
End Function

Private Function AllocateEntitiesToAuditors(auditors As Collection, _
                                            allEntities As Collection) As Object
    ' Returns a Scripting.Dictionary: key = AuditorID, value = Collection of EntityIDs
    On Error GoTo ErrHandler
    
    Dim dict As Object
    Dim i As Long
    Dim audIdx As Long
    Dim nAud As Long
    Dim audID As String
    Dim entID As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Initialize dictionary with empty collections per auditor (preserves order)
    For i = 1 To auditors.Count
        audID = auditors(i)
        If Not dict.Exists(audID) Then
            dict.Add audID, New Collection
        End If
    Next i
    
    nAud = auditors.Count
    If nAud = 0 Or allEntities Is Nothing Or allEntities.Count = 0 Then
        Set AllocateEntitiesToAuditors = dict
        Exit Function
    End If
    
    ' Round-robin allocation: remainder goes to auditors in listed order
    audIdx = 1
    For i = 1 To allEntities.Count
        entID = allEntities(i)
        audID = auditors(audIdx)
        dict(audID).Add entID
        
        audIdx = audIdx + 1
        If audIdx > nAud Then audIdx = 1
    Next i
    
    Set AllocateEntitiesToAuditors = dict
    Exit Function
    
ErrHandler:
    MsgBox "AllocateEntitiesToAuditors failed: " & Err.Description, vbCritical
    Set AllocateEntitiesToAuditors = Nothing
End Function

Private Function ParseAuditorIDs(auditorIDText As String) As Collection
    On Error GoTo ErrHandler
    
    Dim col As New Collection
    Dim cleaned As String
    Dim parts() As String
    Dim i As Long
    Dim token As String
    
    cleaned = Replace(auditorIDText, ";", ",")
    parts = Split(cleaned, ",")
    
    For i = LBound(parts) To UBound(parts)
        token = Trim$(parts(i))
        If Len(token) > 0 Then
            col.Add token
        End If
    Next i
    
    If col.Count = 0 Then
        Set ParseAuditorIDs = Nothing
    Else
        Set ParseAuditorIDs = col
    End If
    Exit Function
    
ErrHandler:
    MsgBox "ParseAuditorIDs failed: " & Err.Description, vbExclamation
    Set ParseAuditorIDs = Nothing
End Function

'===========================
' BUILD RESPONSES ARRAY
'===========================

Public Function BuildResponsesArray(cfg As GenConfig, _
                                    selectedEntities As Collection) As Variant
    On Error GoTo ErrHandler
    
    Dim entData As Variant
    Dim attrData As Variant
    Dim idxEGCI As Long, idxELegalName As Long, idxEJurID As Long
    Dim idxAttrID As Long, idxAttrName As Long, idxQuestionText As Long
    Dim idxAJurID As Long, idxRiskScope As Long, idxIsRequired As Long, idxDocAge As Long
    Dim totalRows As Long
    Dim i As Long, j As Long
    Dim entityRow As Long
    Dim entityID As String
    Dim entityName As String
    Dim entityJurID As String
    Dim attrJur As String
    Dim riskScope As String
    Dim isReq As String
    Dim arr() As Variant
    Dim rowIndex As Long
    
    ' Load data
    If Not LoadEntitiesData(entData, idxEGCI, idxELegalName, idxEJurID) Then Exit Function
    If Not LoadAttributesData(attrData, idxAttrID, idxAttrName, idxQuestionText, _
                              idxAJurID, idxRiskScope, idxIsRequired, idxDocAge) Then Exit Function
    
    ' First pass: count rows
    For i = 1 To selectedEntities.Count
        entityID = CStr(selectedEntities(i))
        entityRow = FindEntityRow(entData, idxEGCI, entityID)
        
        If entityRow > 0 Then
            entityName = CStr(entData(entityRow, idxELegalName))
            entityJurID = CStr(entData(entityRow, idxEJurID))
            
            For j = 1 To UBound(attrData, 1)
                attrJur = CStr(attrData(j, idxAJurID))
                riskScope = UCase$(Trim$(CStr(attrData(j, idxRiskScope))))
                isReq = UCase$(Trim$(CStr(attrData(j, idxIsRequired))))
                
                If (attrJur = entityJurID Or attrJur = "All") Then
                    ' For now we ignore RiskScope and just take required ones
                    If (isReq = "" Or isReq = "Y" Or isReq = "YES") Then
                        totalRows = totalRows + 1
                    End If
                End If
            Next j
        Else
            Debug.Print "Warning: EntityID not found in tblEntities: "; entityID
        End If
    Next i
    
    If totalRows = 0 Then
        BuildResponsesArray = Empty
        Exit Function
    End If
    
    ReDim arr(1 To totalRows, 1 To 14)
    rowIndex = 0
    
    ' Second pass: populate
    For i = 1 To selectedEntities.Count
        entityID = CStr(selectedEntities(i))
        entityRow = FindEntityRow(entData, idxEGCI, entityID)
        
        If entityRow > 0 Then
            entityName = CStr(entData(entityRow, idxELegalName))
            entityJurID = CStr(entData(entityRow, idxEJurID))
            
            For j = 1 To UBound(attrData, 1)
                attrJur = CStr(attrData(j, idxAJurID))
                riskScope = UCase$(Trim$(CStr(attrData(j, idxRiskScope))))
                isReq = UCase$(Trim$(CStr(attrData(j, idxIsRequired))))
                
                If (attrJur = entityJurID Or attrJur = "All") Then
                    If (isReq = "" Or isReq = "Y" Or isReq = "YES") Then
                        rowIndex = rowIndex + 1
                        
                        ' 1: EntityID
                        arr(rowIndex, 1) = entityID
                        ' 2: EntityName
                        arr(rowIndex, 2) = entityName
                        ' 3: JurisdictionID
                        arr(rowIndex, 3) = entityJurID
                        ' 4: AttributeID
                        arr(rowIndex, 4) = attrData(j, idxAttrID)
                        ' 5: AttributeName
                        arr(rowIndex, 5) = attrData(j, idxAttrName)
                        ' 6: QuestionText
                        arr(rowIndex, 6) = attrData(j, idxQuestionText)
                        ' 7: DocumentationAgeRule
                        arr(rowIndex, 7) = attrData(j, idxDocAge)
                        ' 8: ResponseStatus
                        arr(rowIndex, 8) = ""
                        ' 9: DocType
                        arr(rowIndex, 9) = ""
                        ' 10: DocName
                        arr(rowIndex, 10) = ""
                        ' 11: DocDate
                        arr(rowIndex, 11) = ""
                        ' 12: Comments
                        arr(rowIndex, 12) = ""
                        ' 13: AuditorID
                        arr(rowIndex, 13) = cfg.AuditorID
                        ' 14: BatchID
                        arr(rowIndex, 14) = cfg.BatchID
                    End If
                End If
            Next j
        End If
    Next i
    
    BuildResponsesArray = arr
    Exit Function
    
ErrHandler:
    MsgBox "BuildResponsesArray failed: " & Err.Description, vbCritical
    BuildResponsesArray = Empty
End Function

'===========================
' GENERATE AUDITOR WORKBOOK
'===========================

Public Sub GenerateAuditorWorkbook()

If MsgBox("Generate auditor workbooks now?", _
          vbQuestion + vbYesNo, _
          "Confirm Generation") <> vbYes Then
    Exit Sub
End If

    On Error GoTo ErrHandler
    
    Dim cfg As GenConfig
    Dim selectedEntities As Collection
    Dim responses As Variant
    Dim wbNew As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long
    Dim savePath As String
    Dim folder As String
    Dim fileName As String
    Dim accData As Variant
    Dim idxAccAttrID As Long
    Dim idxAccJurID As Long
    Dim idxAccDocName As Long
    Dim dvList As String
    Dim attrID As String
    Dim jurID As String
    
    Dim auditorList As Collection
    Dim allocMap As Object
    Dim k As Long
    Dim curAud As String
    Dim entForThis As Collection
    
    '-----------------------------
    ' 1) Read config & entities
    '-----------------------------
    cfg = ReadGeneratorConfig()
    Set selectedEntities = GetSelectedEntityIDs()
    Call ValidateGeneratorInputs(cfg, selectedEntities)
    
    ' Parse one or more auditors from cfg.AuditorID
    Set auditorList = ParseAuditorIDs(cfg.AuditorID)
    If auditorList Is Nothing Or auditorList.Count = 0 Then
        MsgBox "No valid Auditor IDs found. Please enter one or more IDs " & _
               "in the Auditor ID cell (comma- or semicolon-separated).", vbExclamation
        Exit Sub
    End If
    
    ' Allocate selected entities across auditors (round-robin)
    Set allocMap = AllocateEntitiesToAuditors(auditorList, selectedEntities)
    If allocMap Is Nothing Then
        MsgBox "Failed to allocate entities to auditors.", vbCritical
        Exit Sub
    End If
    
    ' Load acceptable documents once (for DocType drop-downs)
    If Not LoadAcceptableDocsData(accData, idxAccAttrID, idxAccJurID, idxAccDocName) Then
        MsgBox "Warning: AcceptableDocs could not be loaded. DocType validation will be blank.", _
               vbExclamation
    End If
    
    folder = cfg.OutputFolder
    If Right$(folder, 1) <> "\" And Right$(folder, 1) <> "/" Then
        folder = folder & "\"
    End If
    
    '-----------------------------
    ' Loop over auditors in user order
    '-----------------------------
    For k = 1 To auditorList.Count
        curAud = auditorList(k)
        If Not allocMap.Exists(curAud) Then GoTo NextAuditor
        
        Set entForThis = allocMap(curAud)
        If entForThis Is Nothing Or entForThis.Count = 0 Then
            ' This auditor got no entities (e.g. more auditors than GCIs) – skip
            GoTo NextAuditor
        End If
        
        ' Override cfg.AuditorID for this auditor
        cfg.AuditorID = curAud
        
        '-----------------------------
        ' 2) Build responses array for this auditor's entities
        '-----------------------------
        responses = BuildResponsesArray(cfg, entForThis)
        
        If IsEmpty(responses) Then
            MsgBox "No rows were generated for Auditor " & curAud & _
                   ". Check your Attributes and Entities tables.", vbExclamation
            GoTo NextAuditor
        End If
        
        '-----------------------------
        ' 3) Create new workbook
        '-----------------------------
        Set wbNew = Application.Workbooks.Add(xlWBATWorksheet)
        Set ws = wbNew.Worksheets(1)
        ws.Name = "Responses"
        
        '-----------------------------
        ' 4) Write headers + data
        '-----------------------------
        headers = Array( _
            "EntityID", _
            "EntityName", _
            "JurisdictionID", _
            "AttributeID", _
            "AttributeName", _
            "QuestionText", _
            "DocumentationAgeRule", _
            "ResponseStatus", _
            "DocType", _
            "DocName", _
            "DocDate", _
            "Comments", _
            "AuditorID", _
            "BatchID")
        
        ' Headers
        For j = LBound(headers) To UBound(headers)
            ws.Cells(1, j + 1).Value = headers(j)
        Next j
        
        ' Data
        For i = LBound(responses, 1) To UBound(responses, 1)
            For j = LBound(responses, 2) To UBound(responses, 2)
                ws.Cells(i + 1, j).Value = responses(i, j)
            Next j
        Next i
        
        lastRow = UBound(responses, 1) + 1
        lastCol = UBound(responses, 2)
        
        '-----------------------------
        ' 5) Convert to table
        '-----------------------------
        Dim rngTable As Range
        Set rngTable = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        
        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                    Source:=rngTable, _
                                    XlListObjectHasHeaders:=xlYes)
        lo.Name = "tblResponses"
        
        ws.ListObjects("tblResponses").Range.Columns.AutoFit
        ws.Range("A2").Select
        ActiveWindow.FreezePanes = True
        
        '-----------------------------
        ' 6) Data validation
        '-----------------------------
        ' ResponseStatus: Pass / Fail / N/A
        With ws.Range("H2:H" & lastRow).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="Pass,Fail,N/A"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        
        ' DocType: per-row list based on AcceptableDocs
        If Not IsEmpty(accData) Then
            For i = 2 To lastRow
                attrID = CStr(ws.Cells(i, 4).Value) ' AttributeID column
                jurID = CStr(ws.Cells(i, 3).Value)  ' JurisdictionID column
                
                dvList = BuildDocTypeList(attrID, jurID, accData, _
                                          idxAccAttrID, idxAccJurID, idxAccDocName)
                
                With ws.Cells(i, 9).Validation  ' Column I = DocType
                    .Delete
                    If Len(dvList) > 0 Then
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                             Operator:=xlBetween, Formula1:=dvList
                        .IgnoreBlank = True
                        .InCellDropdown = True
                    End If
                End With
            Next i
        End If
        
        '-----------------------------
        ' 7) Save workbook as .xlsx
        '-----------------------------
        fileName = "Onboarding_" & cfg.BatchID & "_" & curAud & "_" & _
                   Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
        savePath = folder & fileName
        
        Application.DisplayAlerts = False
        wbNe   w.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        
        MsgBox "Auditor workbook generated for " & curAud & " with " & _
               entForThis.Count & " entity(ies), saved to:" & vbCrLf & savePath, _
               vbInformation, "Generate Auditor Workbook"
        
NextAuditor:
        On Error Resume Next
        If Not wbNew Is Nothing Then
            wbNew.Close SaveChanges:=False
            Set wbNew = Nothing
        End If
        On Error GoTo ErrHandler
    Next k
    
    Exit Sub
    
ErrHandler:
    MsgBox "GenerateAuditorWorkbook failed: " & Err.Description, vbCritical
    On Error Resume Next
    If Not wbNew Is Nothing Then
        wbNew.Close SaveChanges:=False
    End If
End Sub


