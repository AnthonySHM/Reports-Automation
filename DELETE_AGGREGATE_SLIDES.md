# Delete Aggregate Vulnerability Slides in Multi-Sensor Reports

## Change Summary

For multi-sensor clients, the aggregate (non-sensor-specific) vulnerability and mitigation slides are now **automatically deleted** instead of being populated with combined data.

## Rationale

In multi-sensor reports:
- Each sensor gets its own vulnerability/mitigation slide pair with sensor labels (e.g., "Internal Network Vulnerabilities - VAPRD")
- The aggregate slides (without sensor labels) are redundant and can confuse readers
- Deleting them keeps the report clean and sensor-specific

## Implementation

**File Modified:** `api.py` lines 594-620

### Previous Behavior (Multi-Sensor):
1. Populate per-sensor vulnerability/mitigation slides (with sensor suffix)
2. Populate aggregate vulnerability/mitigation slides (without sensor suffix) with ALL vulnerabilities from all sensors combined

### New Behavior (Multi-Sensor):
1. Populate per-sensor vulnerability/mitigation slides (with sensor suffix)
2. **Delete aggregate vulnerability/mitigation slides** (without sensor suffix)

## Technical Details

### Slide Identification
Aggregate slides are identified by shape names:
- **Vulnerability slide:** Contains shape named `"internal_vulns_content"` (no sensor suffix)
- **Mitigation slide:** Contains shape named `"internal_mitigation_content"` (no sensor suffix)

### Deletion Logic
```python
# Find aggregate slides
agg_vuln_idx = _find_slide_by_shape_name(prs, "internal_vulns_content")
agg_mit_idx = _find_slide_by_shape_name(prs, "internal_mitigation_content")

# Delete in reverse order to maintain indices
slides_to_delete = [agg_vuln_idx, agg_mit_idx]  # if found
for idx in sorted(slides_to_delete, reverse=True):
    delete_slide(prs, idx)
```

Slides are deleted in reverse order (highest index first) to prevent index shifting issues.

## Example: Elephant (3 Sensors)

### Multi-Sensor Template Structure:
```
Slide 12: Internal Network Vulnerabilities - GAPRD
Slide 13: Internal Network Mitigation - GAPRD
Slide 22: Internal Network Vulnerabilities - VAHQ
Slide 23: Internal Network Mitigation - VAHQ
Slide 32: Internal Network Vulnerabilities - VAPRD
Slide 33: Internal Network Mitigation - VAPRD
Slide 34: Internal Network Vulnerabilities (no sensor label) ❌ DELETE
Slide 35: Internal Network Mitigation (no sensor label) ❌ DELETE
```

### After Deletion:
```
Slide 12: Internal Network Vulnerabilities - GAPRD ✓
Slide 13: Internal Network Mitigation - GAPRD ✓
Slide 22: Internal Network Vulnerabilities - VAHQ ✓
Slide 23: Internal Network Mitigation - VAHQ ✓
Slide 32: Internal Network Vulnerabilities - VAPRD ✓
Slide 33: Internal Network Mitigation - VAPRD ✓
[Aggregate slides removed]
```

## Single-Sensor Reports

**No change** - Single-sensor reports continue to use the non-suffixed slides normally since there's no ambiguity.

## Logging

New log messages help track the deletion:
```
INFO: Marked aggregate vulnerability slide (index 34) for deletion
INFO: Marked aggregate mitigation slide (index 35) for deletion
INFO: Deleted aggregate slide at index 35
INFO: Deleted aggregate slide at index 34
INFO: Nuclei population complete: 3 sensors populated, 0 extra slides inserted, 2 aggregate slides deleted for 'elephant'
```

## Benefits

1. **Cleaner Reports:** No redundant aggregate slides in multi-sensor reports
2. **Better Clarity:** Each sensor's vulnerabilities are clearly separated
3. **Consistent Structure:** Only sensor-labeled slides remain
4. **Easier Navigation:** Readers can quickly identify which sensor has which vulnerabilities

## Files Modified

- `api.py`
  - Lines 594-620: Replaced aggregate slide population with deletion logic
  - Line 43: Already imports `_find_slide_by_shape_name` (no change needed)

## Status

✅ Aggregate slides deleted in multi-sensor reports
✅ Single-sensor reports unaffected
✅ Sensor names preserved in continuation slides
✅ Page numbers auto-updated after deletion

**Server restart required to apply changes.**
