# Generation Review Workflow - Runbook

## Overview

This two-step workflow allows you to:
1. **Review & Edit** assignments before generation
2. **Generate** workbooks from the reviewed assignments

### Benefits
- Preview all assignments in one table before generating
- Edit AuditorID, Jurisdiction, or entity assignments
- Add/remove entities from generation batch
- Preserve manual edits when refreshing from source

---

## Prerequisites

### Required Tables
- ✅ **tblEntities** (source data) - must have columns:
  - `GCI`
  - `AuditorID`
  - `Jurisdiction ID`
  - `Legal Entity Name` or `Legal Name`
  - `Jurisdiction Name` or `Jurisdiction`
  - `Auditor Name` (optional)

- ✅ **tblGenerationReview** (on "Generation Review" sheet) - must have columns:
  - `GCI`
  - `Legal Name`
  - `Jurisdiction ID`
  - `Jurisdiction`
  - `AuditorID`
  - `Auditor Name`

- ✅ **tblAttributes** (attribute definitions)
- ✅ **tblAcceptableDocs** (document validation lists)

### Required Named Ranges
- ✅ **rngBatchID** - single cell containing batch identifier
- ✅ **rngOutputFolder** - single cell containing output folder path

### Required Sheets
- ✅ **Generation Review** - sheet containing tblGenerationReview table
  - Should have labels "Last Refresh" and "Generator Status" somewhere on the sheet (optional but recommended)

---

## Installation

### Step 1: Apply MatrixLayout Patch

1. Open `Question Library.xlsm`
2. Press **Alt+F11** to open VBA Editor
3. Find the **VBA Script_MatrixLayout_PATCHED** module in Project Explorer
4. Open **MatrixLayout_PATCH_FOR_REVIEW.txt** in a text editor
5. Make the two changes specified:
   - Line ~227: Change `Private Sub BuildJurisdictionSheet` → `Public Sub BuildJurisdictionSheet`
   - Line ~280: Change `Private Sub BuildIndexHeader` → `Public Sub BuildIndexHeader`
6. **Debug → Compile VBAProject** to verify no errors

### Step 2: Import modGenerationReview Module

1. In VBA Editor: **File → Import File** (or Ctrl+M)
2. Select **modGenerationReview.bas**
3. Click **Open**
4. Verify "modGenerationReview" appears in Project Explorer
5. **Debug → Compile VBAProject** to verify no errors

### Step 3: Set Up Review Table (If Not Already Exists)

If you don't have a "Generation Review" sheet with tblGenerationReview:

1. Create new sheet named **"Generation Review"**
2. In cell A1, add label: **Generator Status**
3. In cell A2, add label: **Last Refresh**
4. In cell A4, create headers:
   | GCI | Legal Name | Jurisdiction ID | Jurisdiction | AuditorID | Auditor Name |
5. Select range A4:F4 (headers + at least one data row)
6. **Insert → Table** (Ctrl+T)
7. Name the table: **tblGenerationReview** (Table Tools → Design → Table Name)

### Step 4: Assign Buttons/Macros

#### Option A: Assign to Existing Buttons

If you have buttons on the Generation Review sheet:

1. Right-click the button
2. Select **Assign Macro**
3. For "Populate" button: Select `modGenerationReview.PopulateGenerationReview`
4. For "Generate" button: Select `modGenerationReview.GenerateAuditorWorkbooks_FromReview`
5. Click OK

#### Option B: Create New Buttons

1. Go to **Generation Review** sheet
2. **Insert → Shapes → Rounded Rectangle** (or any shape)
3. Draw the shape, type label: **"1. Populate Review"**
4. Right-click shape → **Assign Macro** → Select `PopulateGenerationReview`
5. Repeat for second button: **"2. Generate Workbooks"**
6. Assign macro: `GenerateAuditorWorkbooks_FromReview`

#### Option C: Run from VBA Editor (for testing)

1. Press **Alt+F8** (Macros dialog)
2. Select macro name and click **Run**

---

## Usage Workflow

### Step 1: Populate Generation Review

**Purpose**: Load assignments from tblEntities into the review table for preview/editing.

**How to Run**:
- Click **"1. Populate Review"** button, OR
- Press **Alt+F8** → `PopulateGenerationReview` → Run

**What Happens**:
1. Reads all rows from `tblEntities`
2. Populates `tblGenerationReview` with:
   - GCI, Legal Name, Jurisdiction ID, Jurisdiction, AuditorID, Auditor Name
3. Updates "Last Refresh" timestamp
4. Sets "Generator Status" to "Ready (review N assignment(s) below)"
5. Shows confirmation message

**Parameters**:
- **Default**: Preserves existing values if GCI already exists in review table
- **Force Refresh**: To overwrite all existing data, edit the button macro to call:
  ```vb
  Call PopulateGenerationReview(ForceRefresh:=True)
  ```

**Example Confirmation**:
```
Review table populated with 150 assignment(s).
Please review and edit as needed, then run 'Generate Auditor Workbooks'.
```

### Step 2: Review & Edit (Manual)

**What to Check**:
- ✅ **AuditorID** - Is each GCI assigned to the correct auditor?
- ✅ **Jurisdiction ID** - Is the jurisdiction correct?
- ✅ **GCI** - Are all entities included? Remove rows for entities you don't want to generate.

**What You Can Edit**:
- ✅ Change `AuditorID` to reassign entities
- ✅ Change `Jurisdiction ID` or `Jurisdiction` name
- ✅ Change `Legal Name` (for display only, doesn't affect generation)
- ✅ Delete rows to exclude from generation
- ✅ Add rows manually (ensure GCI, AuditorID, Jurisdiction ID are filled)

**What NOT to Edit**:
- ⚠️ Don't remove required columns
- ⚠️ Don't leave GCI, AuditorID, or Jurisdiction ID blank (row will be skipped)

### Step 3: Generate Auditor Workbooks

**Purpose**: Generate one workbook per auditor based on the review table.

**How to Run**:
- Click **"2. Generate Workbooks"** button, OR
- Press **Alt+F8** → `GenerateAuditorWorkbooks_FromReview` → Run

**What Happens**:
1. Reads `tblGenerationReview` (not tblEntities!)
2. Groups assignments by AuditorID
3. For each auditor:
   - Creates workbook with Index sheet
   - Creates ENT sheet
   - Creates one sheet per unique Jurisdiction
   - Populates matrix with attributes and data validation
   - Saves to `OutputFolder` as `KYC_Workpapers_<BatchID>_<AuditorID>.xlsx`
4. Updates "Generator Status" with completion message
5. Shows summary: "Successfully generated N of N workbook(s)."

**Confirmation Dialog**:
```
Generate auditor workbooks from the review table?

This will create workbooks based on the assignments in the Generation Review table.
```

**Example Output**:
```
Output\KYC_Workpapers_BATCH_2025_001_A001.xlsx
Output\KYC_Workpapers_BATCH_2025_001_A002.xlsx
Output\KYC_Workpapers_BATCH_2025_001_A003.xlsx
```

---

## Validation Checklist

Before running generation, verify:

- [ ] `tblGenerationReview` has at least one row with:
  - Non-blank `GCI`
  - Non-blank `AuditorID`
  - Non-blank `Jurisdiction ID`
- [ ] Named range `rngBatchID` is populated
- [ ] Named range `rngOutputFolder` points to valid directory
- [ ] Review table columns match expected names (case-insensitive)

---

## Troubleshooting

### Issue: "Source table 'tblEntities' not found"

**Cause**: The source table doesn't exist or is named differently.

**Solution**:
1. Verify table exists: **Formulas → Name Manager** → look for "tblEntities"
2. If named differently, edit constant in `modGenerationReview`:
   ```vb
   Private Const SOURCE_TABLE_NAME As String = "YourTableName"
   ```

### Issue: "Table 'tblGenerationReview' not found"

**Cause**: The review table doesn't exist or is on wrong sheet.

**Solution**:
1. Verify sheet name is exactly **"Generation Review"**
2. Verify table name is exactly **"tblGenerationReview"** (check Table Tools → Design → Table Name)
3. If different, edit constants in `modGenerationReview`:
   ```vb
   Private Const REVIEW_SHEET_NAME As String = "Your Sheet Name"
   Private Const REVIEW_TABLE_NAME As String = "YourTableName"
   ```

### Issue: "Source table is missing required columns: [list]"

**Cause**: tblEntities is missing one of: GCI, AuditorID, Jurisdiction ID.

**Solution**:
1. Check column names in tblEntities (exact match, case-insensitive)
2. Add missing columns, OR
3. If columns exist with different names, edit the code to add aliases in `GetCellValue` calls

**Example**: If your table has "Entity ID" instead of "GCI":
```vb
' In PopulateGenerationReview, change:
reviewData(i, 1) = GetCellValue(srcData, i, idxSrc, "Entity ID")  ' was "GCI"
```

### Issue: "Review table is missing required columns: [list]"

**Cause**: tblGenerationReview doesn't have the expected column structure.

**Solution**:
1. Verify column names: GCI, Legal Name, Jurisdiction ID, Jurisdiction, AuditorID, Auditor Name
2. Rename columns to match, OR
3. Edit `WriteToReviewTable` to match your column structure

### Issue: "No valid assignments found in review table"

**Cause**: All rows have blank GCI, AuditorID, or Jurisdiction ID.

**Solution**:
1. Run `PopulateGenerationReview` first
2. Check Immediate Window (Ctrl+G) for "Skipped N row(s) with missing key fields"
3. Fill in blank cells in review table

### Issue: "Named range not found or invalid: rngBatchID"

**Cause**: Named range doesn't exist or is broken.

**Solution**:
1. **Formulas → Name Manager**
2. Check if `rngBatchID` exists and points to valid cell
3. If missing, create it:
   - Select the cell containing Batch ID
   - Name Box (top-left) → type "rngBatchID" → Enter

### Issue: Generated workbook has no sheets / missing sheets

**Cause**: Review table has no assignments for that auditor/jurisdiction, OR attribute filtering excluded all questions.

**Solution**:
1. Check Immediate Window (Ctrl+G) for messages like:
   - "Auditor A001 has 0 assignments"
   - "Attributes for this jurisdiction: 0"
2. Verify `tblAttributes` has rows matching the Jurisdiction IDs in your review table
3. Check `IsRequired` column in tblAttributes (if all "N", nothing generates)

### Issue: Compile error when running macros

**Cause**: Module not properly imported or dependencies missing.

**Solution**:
1. **Debug → Compile VBAProject**
2. Check error message for specific issue
3. Verify both modules are imported:
   - `VBA Script_MatrixLayout_PATCHED` (with Public functions)
   - `modGenerationReview`
4. Verify GenConfig type is defined (should be in MatrixLayout module)

### Issue: Error 450 or Error 13 when generating

**Cause**: Likely related to data validation or named range issues (should be fixed in patched version).

**Solution**:
1. Check `_Log` sheet for detailed error trace
2. Copy Immediate Window contents (Ctrl+G, Ctrl+A, Ctrl+C)
3. Verify you're using the **PATCHED** version of MatrixLayout
4. Check that data validation formulas aren't exceeding 255 characters

---

## Advanced: Incremental Refresh

### Scenario: Preserve Manual Edits When Source Changes

If you've manually edited the review table and want to add new entities without losing edits:

**Workflow**:
1. Run `PopulateGenerationReview` normally (default preserves edits)
2. Existing GCIs keep their current values
3. New GCIs from tblEntities are added
4. Deleted GCIs in source remain in review table (you must delete manually)

**Force Refresh** (overwrites all edits):
1. Edit button macro to:
   ```vb
   Call PopulateGenerationReview(ForceRefresh:=True)
   ```
2. Or run from Immediate Window: `PopulateGenerationReview True`

---

## Monitoring & Logging

### Immediate Window (Recommended)

**How to View**:
1. Open VBA Editor (Alt+F11)
2. Press **Ctrl+G** to show Immediate Window
3. Run macros and watch real-time log

**Example Log Output**:
```
14:35:22 | [Review] === PopulateGenerationReview START ===
14:35:22 | [Review] ForceRefresh: False
14:35:22 | [Review] Source data loaded: 150 rows
14:35:23 | [Review] BuildExistingReviewMap: Preserved 145 existing row(s)
14:35:23 | [Review] WriteToReviewTable: Writing 150 rows x 6 cols
14:35:23 | [Review] Updated Last Refresh: 2025-12-19 14:35:23
14:35:23 | [Review] === PopulateGenerationReview COMPLETE ===

14:40:10 | [Review] === GenerateAuditorWorkbooks_FromReview START ===
14:40:10 | [Review] Config read: BatchID=BATCH_2025_001, OutputFolder=C:\Output
14:40:10 | [Review] BuildAuditorMapFromReview: Built map with 3 auditor(s)
14:40:11 | [Review] Generating workbook for AuditorID: A001
14:40:11 | [Review] Auditor A001 has 52 assignment(s)
14:40:12 | [Review] Found 4 non-ENT jurisdiction(s)
14:40:14 | [Review] Saving workbook to: C:\Output\KYC_Workpapers_BATCH_2025_001_A001.xlsx
14:40:15 | [Review] GenerateOneAuditorWorkbook_FromReview: COMPLETE for A001
14:40:15 | [Review] === GenerateAuditorWorkbooks_FromReview COMPLETE ===
```

### _Log Sheet (Persistent)

All log messages are also written to a `_Log` sheet in Question Library.xlsm.

**How to View**:
1. Find the `_Log` sheet (may be hidden)
2. Review timestamps and messages
3. Filter or sort to find specific issues

**Maintenance**:
- Log grows with each run
- Manually clear old entries: Select rows 2+ → Delete

---

## Expected Column Mapping

### tblEntities (Source) → tblGenerationReview (Target)

| Source Column (tblEntities)  | Target Column (tblGenerationReview) | Notes |
|------------------------------|-------------------------------------|-------|
| GCI                          | GCI                                 | Key field |
| Legal Entity Name OR Legal Name | Legal Name                       | Display only |
| Jurisdiction ID              | Jurisdiction ID                     | Key field |
| Jurisdiction Name OR Jurisdiction | Jurisdiction                   | Display only |
| AuditorID                    | AuditorID                           | Key field |
| Auditor Name                 | Auditor Name                        | Optional |

### Additional Fields in Full tblEntities (Not Used in Review)

These columns exist in tblEntities but are NOT shown in the review table:
- Party Type
- Onboarding Date
- IRR, DRR, Primary FLU
- Case ID # (Aware)

The code adds these as empty values when building the auditor map, so they won't break generation even if not in the review table.

---

## Button Setup Examples

### Example 1: Simple Buttons with Macro Assignment

```
┌─────────────────────────────────┐
│  Generation Review              │
│                                 │
│  Generator Status: Ready        │
│  Last Refresh: 2025-12-19       │
│                                 │
│  [1. Populate Review]           │  ← Macro: PopulateGenerationReview
│  [2. Generate Workbooks]        │  ← Macro: GenerateAuditorWorkbooks_FromReview
│                                 │
│  ┌────────────────────────┐    │
│  │ tblGenerationReview    │    │
│  ├────────────────────────┤    │
│  │ GCI | Legal Name | ... │    │
│  └────────────────────────┘    │
└─────────────────────────────────┘
```

### Example 2: Buttons with Force Refresh Option

Create a third button for force refresh:

**Macro Code** (create in modGenerationReview or a new module):
```vb
Public Sub PopulateGenerationReview_Force()
    Call PopulateGenerationReview(ForceRefresh:=True)
End Sub
```

**Button Layout**:
```
[1a. Populate (Preserve Edits)]  ← PopulateGenerationReview
[1b. Populate (Force Refresh)]   ← PopulateGenerationReview_Force
[2. Generate Workbooks]           ← GenerateAuditorWorkbooks_FromReview
```

---

## Success Criteria

✅ **Populate Review** completed successfully when:
- Message shows: "Review table populated with N assignment(s)"
- tblGenerationReview has rows matching tblEntities count
- "Last Refresh" cell shows current timestamp
- "Generator Status" shows "Ready (review N assignment(s) below)"

✅ **Generate Workbooks** completed successfully when:
- Message shows: "Successfully generated N of N workbook(s)"
- Output folder contains .xlsx files (one per auditor)
- Each workbook has Index sheet + ENT sheet + jurisdiction sheets
- Data validation dropdowns work in matrix cells
- "Generator Status" updated with generation timestamp

---

## FAQ

**Q: Can I add rows manually to the review table?**
A: Yes! Just ensure GCI, AuditorID, and Jurisdiction ID are filled in. Other columns are optional.

**Q: What happens if I delete rows from the review table?**
A: Those GCIs won't be included in generation. This is useful for excluding specific entities.

**Q: Can I run generation multiple times?**
A: Yes. Each run overwrites previous output files with the same name. If you need to keep old versions, rename them before re-running.

**Q: Will this work with the original (unpatched) MatrixLayout?**
A: No. You must use the PATCHED version with Public functions. The original has Private functions that can't be called from another module.

**Q: Can I undo changes to the review table?**
A: If you haven't saved the workbook yet, press Ctrl+Z to undo. Otherwise, re-run `PopulateGenerationReview(ForceRefresh:=True)` to reset to source data.

**Q: What if I have multiple batches?**
A: Change `rngBatchID` named range to point to different batch ID cell. Review table will reflect whatever batch data is in tblEntities.

---

## Support & Diagnostics

If issues persist:

1. **Capture Diagnostics**:
   - Screenshot of error message (include error number)
   - Copy Immediate Window contents (Ctrl+G → Ctrl+A → Ctrl+C)
   - Export `_Log` sheet to CSV
   - Note which step failed (Populate or Generate)

2. **Check Configuration**:
   - Verify table names: `tblEntities`, `tblGenerationReview`, `tblAttributes`, `tblAcceptableDocs`
   - Verify sheet name: "Generation Review"
   - Verify named ranges: `rngBatchID`, `rngOutputFolder`
   - Check column names match expected (GCI, AuditorID, Jurisdiction ID, etc.)

3. **Test Components**:
   - Run `PopulateGenerationReview` alone first
   - Check review table is populated correctly
   - Then run `GenerateAuditorWorkbooks_FromReview`

---

**Document Version**: 1.0
**Last Updated**: 2025-12-19
**Modules Required**: modGenerationReview.bas, VBA Script_MatrixLayout_PATCHED.vb (with Public functions)
