# Microsoft Patches Table Implementation

## Overview

The Top Microsoft Patches slide (Slide 16) now displays a table with unique KB patches extracted from the remediation plan CSV, showing:
- **Package Name**: The Windows OS or application name
- **KB Patch**: The KB number (e.g., KB5068865)

## Changes Made

### 1. Template Layout (Already Complete)

Both table slides (15 & 16) now have:
- **Note area at bottom**: Positioned at 4.50" from top, just above footer
- **Table constrained**: Maximum height 3.25" to never overlap note
- **0.15" gap**: Between table and note for clean separation

### 2. New Function: `build_ms_patches_csv()`

**Location**: `core/drive_agent.py`

**Purpose**: Extracts unique KB patches from remediation plan CSV and creates a clean two-column table.

**Features**:
- Case-insensitive column detection
- Finds "Required KB" or "KB Patch" column
- Finds "Package Name", "Host Name", or similar columns
- Deduplicates KB numbers
- Outputs CSV with "Package Name" and "KB Patch" columns
- Sorts by KB number for consistent display

**Returns**: `(csv_path, unique_count)` tuple

### 3. API Integration

**Location**: `api.py`

**Flow**:
```
Fetch remediation plan CSV
    ↓
Count unique KB patches
    ↓
Build MS patches CSV table
    ↓
Create CSVTableAsset (slide 15, shape "ms_patches_table")
    ↓
Inject into patches_csv_assets list
    ↓
Place table on slide
    ↓
Optimize to image
```

**Added for both**:
- Single-sensor reports
- Multi-sensor reports

### 4. Column Detection Logic

The function searches for columns with flexible matching:

**KB Column** (case-insensitive):
- "Required KB"
- "required_kb"
- "KB Patch"
- "kb_patch"

**Package Column** (case-insensitive):
- "Package Name"
- "package_name"
- "Host Name"
- "hostname"
- Any column containing: "package", "host", "name", "os", "system"

**Fallback**: If package column not found, uses KB number for both columns.

## Example Output

From remediation plan CSV like:
```csv
Host Name,Required KB,Status
DC01-WINDOWS-2019,KB5068791,Missing
WS-2022-SERVER,KB5068787,Missing
WIN11-CLIENT-01,KB5068865,Missing
WIN11-CLIENT-02,KB5068865,Missing
```

Generated top_ms_patches.csv:
```csv
Package Name,KB Patch
DC01-WINDOWS-2019,KB5068791
WS-2022-SERVER,KB5068787
WIN11-CLIENT-01,KB5068865
```

**Note**: Duplicates (KB5068865) are automatically removed, showing only unique KBs.

## Table Display

**Slide 16 - Top Microsoft Patches**:
- Black header row: "Package Name" | "KB Patch"
- Alternating blue data rows
- Shows unique KB patches sorted alphabetically
- Note at bottom with patching cautions
- Table rendered as optimized image (via table_optimizer)

## Data Flow

### Single-Sensor Reports

```
1. Fetch remediation_plan CSV from VN folder
2. Count unique KBs → kb_count
3. Build MS patches CSV → ms_patches_asset (slide 15)
4. Build software patches CSV → software_patches_asset (slide 14)
5. Fetch NDR CSV assets
6. Inject both patches CSVs if not from NDR
7. Place all CSV tables
8. Optimize tables to images
```

### Multi-Sensor Reports

```
1. Fetch remediation_plan CSV from VN folder
2. Count unique KBs → kb_count
3. Build MS patches CSV → ms_patches_asset
4. Build software patches CSV → software_patches_asset
5. Fetch per-sensor NDR CSVs
6. Separate sensor CSVs from patches CSVs
7. Inject both patches CSVs into patches_csv_assets
8. Adjust slide indices for multi-sensor layout
9. Place all tables
10. Optimize to images
```

## Files Modified

1. **`core/drive_agent.py`**:
   - Added `build_ms_patches_csv()` function (162 lines)

2. **`api.py`**:
   - Added `ms_patches_asset` variable
   - Added `build_ms_patches_csv` import
   - Added MS patches CSV building after KB count
   - Added MS patches asset injection (multi-sensor path)
   - Added MS patches asset injection (single-sensor path)

3. **`create_template.py`** (already complete):
   - Note area repositioned to bottom on both slides 15 & 16

4. **`core/table_optimizer.py`** (already complete):
   - Updated area constraints for table slides

## Testing

To verify the implementation:

1. **Generate a report** with remediation plan CSV
2. **Check slide 16** for table with KB patches
3. **Verify columns**: "Package Name" and "KB Patch"
4. **Check uniqueness**: No duplicate KB numbers
5. **Verify note position**: At bottom, above footer
6. **Check gap**: Table should not overlap note

## Error Handling

- **No remediation CSV**: MS patches table not created (graceful)
- **Column not found**: Logs warning, returns None
- **Empty KB list**: Logs info, returns None
- **CSV read error**: Logs error, returns None

All errors are non-fatal and allow report generation to continue.

## Future Enhancements

Possible improvements:
- [ ] Add "Severity" column if available in remediation plan
- [ ] Group by Windows version (11, Server 2019, etc.)
- [ ] Add installation date/age information
- [ ] Link to Microsoft KB article URLs
- [ ] Sort by severity instead of KB number
