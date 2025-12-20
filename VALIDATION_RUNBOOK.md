# VBA Generator - Patch Validation Runbook

## Summary of Fixes

### Critical Errors Fixed

#### 1. **Error 13: Type Mismatch** - Named Range Array Issue
- **Location**: `GetNamedRangeValue()` function
- **Root Cause**: When a named range refers to multiple cells, `.Value` returns a Variant array, not a single value. Using `CStr()` on an array causes Error 13.
- **Fix**: Added check for `rng.Cells.Count = 1` before accessing `.Value`. If multi-cell range, uses first cell only with warning logged.

#### 2. **Error 450: Wrong number of arguments** - Validation.Add Issue
- **Locations**:
  - `ApplyDvToCell()` in MatrixLayout (lines 371-376)
  - `GenerateAuditorWorkbook()` in SplitGCIsEvenly (lines 706-712, 723-731)
- **Root Cause**: `Validation.Add` with `Type:=xlValidateList` should NOT have an `Operator:=xlBetween` parameter. This is invalid for list validation.
- **Fix**: Removed `Operator:=xlBetween` parameter from all `Validation.Add` calls.

#### 3. **File Format Mismatch**
- **Location**: `GenerateOneAuditorWorkbook()` SaveAs call (line 140)
- **Root Cause**: Using `FileFormat:=xlOpenXMLWorkbookMacroEnabled` (52) but output filename has `.xlsx` extension (non-macro format).
- **Fix**: Changed to `FileFormat:=xlOpenXMLWorkbook` (51) and filename extension to `.xlsx`.

#### 4. **Syntax Error** - Typo in SplitGCIsEvenly
- **Location**: Line 743
- **Root Cause**: `wbNe   w.SaveAs` - extra spaces in variable name would cause compile error.
- **Fix**: Corrected to `wbNew.SaveAs`.

### Enhancements Added

#### Debug Logging System
- Added `DebugLog(msg)` helper function that:
  - Writes timestamped messages to Immediate Window (Ctrl+G in VBA Editor)
  - Optionally creates a `_Log` sheet in the workbook to persist logs
- Integrated logging at key points:
  - Entry/exit of major functions
  - Configuration reads
  - Table loads
  - Workbook generation steps
  - Error conditions

---

## Installation Steps

### Option A: Import Patched Modules (Recommended)

1. **Open Excel Workbook**
   - Open `Question Library.xlsm`

2. **Open VBA Editor**
   - Press `Alt+F11`

3. **Remove Old Modules**
   - In Project Explorer (left pane), find and select the old module(s)
   - Right-click → Remove
   - Choose "No" when asked to export (you already have backups)

4. **Import Patched Modules**
   - File → Import File (or Ctrl+M)
   - Navigate to:
     - `VBA Script_MatrixLayout_PATCHED.vb`
     - `VBA Script_SplitGCIsEvenly_PATCHED.vb` (if using this module)
   - Click "Open" to import

5. **Verify Import**
   - In Project Explorer, you should see the imported modules
   - Double-click to view the code and verify it's the patched version
   - Look for comments like `' FIX:` to confirm

6. **Save Workbook**
   - Press `Ctrl+S` or File → Save
   - Close VBA Editor

### Option B: Manual Copy-Paste (If Import Fails)

1. Open the patched `.vb` file in a text editor
2. Copy entire contents
3. In VBA Editor, create new module or open existing module
4. Delete all existing code in the module
5. Paste the copied code
6. Save

---

## Validation Tests

### Pre-Flight Checklist

Before running the generator, verify these named ranges exist and are correctly defined:

1. **rngBatchID** - Should point to a **single cell** containing the batch ID
2. **rngOutputFolder** - Should point to a **single cell** containing output folder path
3. **rngAuditorID** - (For SplitGCIsEvenly only) Single cell with auditor ID(s)
4. **rngSelectedEntities** - (For SplitGCIsEvenly only) Range containing entity IDs

**How to Check Named Ranges:**
- Go to Formulas tab → Name Manager
- Verify each range points to the correct location
- Ensure single-cell ranges are NOT pointing to multi-cell ranges

### Test 1: Basic Configuration Read

**Objective**: Verify named ranges are read correctly without Error 13.

**Steps**:
1. Open VBA Editor (`Alt+F11`)
2. Press `Ctrl+G` to open Immediate Window
3. In Immediate Window, type:
   ```vb
   ?ThisWorkbook.Names("rngBatchID").RefersToRange.Address
   ```
4. Press Enter - should show address like `$B$5`
5. Repeat for `rngOutputFolder`

**Expected Result**: Single cell addresses displayed without errors.

**If Error Occurs**: Named range may point to multiple cells or an invalid range. Fix in Name Manager.

### Test 2: Run with Debug Logging

**Objective**: Generate workbooks with full debug logging to trace execution.

**Steps** (for MatrixLayout):
1. Ensure your data tables are populated:
   - `tblEntities` has data with AuditorID, GCI, Jurisdiction columns
   - `tblAttributes` has data
   - `tblAcceptableDocs` has data
2. Open VBA Editor (`Alt+F11`)
3. Press `Ctrl+G` to open Immediate Window (keep it visible)
4. Run the generator:
   - Press `Alt+F8` to open Macro dialog
   - Select `GenerateWorkpapers`
   - Click "Run"
5. **Watch Immediate Window** for timestamped log messages like:
   ```
   14:35:22 | === GenerateWorkpapers START ===
   14:35:22 | Config read: BatchID=BATCH_001, OutputFolder=C:\Output
   14:35:22 | Config validated successfully
   14:35:22 | Tables loaded: tblEntities, tblAttributes, tblAcceptableDocs
   ```
6. If the macro runs successfully, check:
   - Output folder for generated `.xlsx` files
   - `_Log` sheet in `Question Library.xlsm` for persistent log

**Expected Result**:
- No runtime errors (13, 450, etc.)
- Log messages in Immediate Window showing progress
- Output files created in specified folder with `.xlsx` extension

**If Error Occurs**:
- Check the error number and description in the message box
- Check Immediate Window for the last successful log entry
- Review `_Log` sheet for full execution trace

### Test 3: Validate Data Validation (DV) in Output

**Objective**: Verify that data validation dropdowns are correctly applied without Error 450.

**Steps**:
1. Open one of the generated `.xlsx` files from the Output folder
2. Navigate to a jurisdiction sheet (e.g., "ENT" or state sheets)
3. Click on a cell in the matrix area (column H and beyond)
4. Look for a dropdown arrow in the cell
5. Click the dropdown - should show options like:
   - Document names from `tblAcceptableDocs`
   - "Fail 1"
   - "Fail 2"
6. Verify no error occurs when selecting an option

**Expected Result**: Dropdowns work correctly with no errors.

**If DV Missing**: Check that `tblAcceptableDocs` has matching Attribute IDs.

### Test 4: File Format Verification

**Objective**: Ensure output files are saved in correct format (.xlsx, not .xlsm).

**Steps**:
1. Navigate to the Output folder
2. Check file extensions - should be `.xlsx` NOT `.xlsm`
3. Right-click a generated file → Properties
4. Verify "Type of file" shows "Microsoft Excel Worksheet (.xlsx)"
5. Open the file
6. Verify no macro security warnings appear (since it's .xlsx with no macros)

**Expected Result**: All files are `.xlsx` format with no macro warnings.

### Test 5: Multi-Cell Named Range Warning (Edge Case)

**Objective**: Test that the code gracefully handles multi-cell named ranges.

**Steps**:
1. Temporarily modify `rngBatchID` to point to a range like `$B$5:$B$7` (multiple cells)
2. Run `GenerateWorkpapers`
3. Check Immediate Window for warning:
   ```
   14:40:15 | WARNING: Named range 'rngBatchID' contains multiple cells. Using first cell only.
   ```
4. Verify generation continues using value from B5
5. **Restore** `rngBatchID` to single cell before production use

**Expected Result**: Warning logged, generation continues with first cell value.

---

## Troubleshooting

### Issue: Still Getting Error 13 on Named Range

**Diagnosis**:
- Named range may be pointing to an invalid reference
- Named range may contain a formula that returns an array

**Solution**:
1. Formulas → Name Manager
2. Select the problematic range
3. Click "Edit"
4. Verify "Refers to" points to a single cell (e.g., `=Sheet1!$B$5`)
5. If it's a formula, replace with direct cell reference
6. Click OK and test again

### Issue: Error 450 Still Occurring

**Diagnosis**:
- Validation formula may be too long (>255 characters)
- External range reference may be invalid

**Solution**:
1. Check `tblAcceptableDocs` - if too many documents per attribute, the comma-separated list may exceed 255 chars
2. Reduce number of acceptable docs, or
3. Modify `ApplyDvToCell()` to use a named range instead of direct formula

### Issue: Generated Files Are Empty or Missing Sheets

**Diagnosis**:
- Data tables may be empty
- Jurisdiction filtering may exclude all records

**Solution**:
1. Check Immediate Window / `_Log` sheet for messages like:
   - "Auditor X has 0 assignments"
   - "Attributes for this jurisdiction: 0"
2. Verify `tblEntities` has rows with matching AuditorID and Jurisdiction ID
3. Verify `tblAttributes` has rows with matching Jurisdiction ID
4. Check IsRequired column - if all "N", no rows will be generated

### Issue: Compile Error When Opening Workbook

**Diagnosis**:
- Module import may have corrupted
- References may be missing

**Solution**:
1. Tools → References in VBA Editor
2. Verify no "MISSING:" entries
3. Ensure "Microsoft Scripting Runtime" is checked (for Dictionary)
4. If corrupt, re-import the patched module

---

## Success Criteria

The patch is successfully applied when:

1. ✅ **No Error 13**: Named ranges read correctly as single values
2. ✅ **No Error 450**: Data validation applied without argument errors
3. ✅ **Correct File Format**: Output files are `.xlsx` (not `.xlsm`)
4. ✅ **Debug Logging Works**: Immediate Window shows timestamped log messages
5. ✅ **Generation Completes**: "Generation complete" message appears
6. ✅ **Output Files Valid**: Generated workbooks open without errors and contain expected data

---

## Rollback Plan

If issues occur after patching:

1. Close Excel completely
2. Open `Question Library.xlsm`
3. In VBA Editor, remove patched modules
4. Import the original `.vb` files (the unpatched versions)
5. Report the issue with screenshots of error messages and Immediate Window log

---

## Contact & Support

If you encounter issues not covered in this runbook:

1. **Capture diagnostics**:
   - Screenshot of error message (including error number)
   - Copy contents of Immediate Window (`Ctrl+A` in Immediate Window, then `Ctrl+C`)
   - Copy contents of `_Log` sheet if it exists
   - Note which test step failed

2. **Provide context**:
   - What data is in your tables (row counts)
   - What named ranges are configured
   - Whether this is first run or previously worked

3. **Include log file**: Share the `_Log` sheet or Immediate Window output

---

## Appendix: Quick Reference - Fixes Applied

| Issue | File | Line(s) | Fix |
|-------|------|---------|-----|
| Error 13: Type mismatch | MatrixLayout | 515-521 | Added check: `If rng.Cells.Count = 1` before `.Value` |
| Error 13: Type mismatch | SplitGCIsEvenly | 117-152 | Same fix in `ReadGeneratorConfig()` |
| Error 450: Wrong args | MatrixLayout | 371-376 | Removed `Operator:=xlBetween` from `Validation.Add` |
| Error 450: Wrong args | SplitGCIsEvenly | 706-712 | Removed `Operator:=xlBetween` from `Validation.Add` |
| Error 450: Wrong args | SplitGCIsEvenly | 723-731 | Removed `Operator:=xlBetween` from `Validation.Add` |
| File format mismatch | MatrixLayout | 140, 533 | Changed to `xlOpenXMLWorkbook` (51) and `.xlsx` extension |
| Syntax error (typo) | SplitGCIsEvenly | 743 | Fixed `wbNe   w` → `wbNew` |
| Enhancement | Both | Top | Added `DebugLog()` function with Immediate Window + `_Log` sheet |

---

**Document Version**: 1.0
**Last Updated**: 2025-12-19
**Patched Files**: VBA Script_MatrixLayout_PATCHED.vb, VBA Script_SplitGCIsEvenly_PATCHED.vb
