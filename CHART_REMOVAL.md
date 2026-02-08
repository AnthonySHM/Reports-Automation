# Patch Trend Chart Removal

## Changes Made

The patch trend chart has been **removed** from the Required Software Patches slide (Slide 14).

## Files Modified

### 1. `create_template.py`
- **Reverted** `slide_required_patches()` function to original text-only layout
- Removed chart placeholder area
- Removed "Missing Patch Quantity" subtitle
- Restored original spacing and font sizes for bullet points

### 2. `api.py`
- **Removed** chart generation code from patch count processing
- **Removed** `patch_chart_asset` creation
- **Removed** chart placement for both single-sensor and multi-sensor reports
- Chart generation logic completely disabled

### 3. `templates/mantix4_report_template.pptx`
- **Regenerated** without chart placeholder
- Slide 14 now shows only the text content

## Current State

**Slide 14 - Required Software Patches** now contains:
- Black title bar: "Required Software Patches"
- Grey content panel with bullet points:
  - "There are currently [Count] Microsoft KB patches that need to be deployed."
  - "There are [Count] software packages that require updating. Patching should be done with caution to ensure systems and applications are not negatively impacted."
- No chart or graph

## Files Retained (Not Deleted)

The following files remain in the codebase but are **not being used**:

- `core/patch_chart.py` - Chart generator module
- `core/patch_history.py` - History management module
- `PATCH_TREND_CHART.md` - Chart documentation
- `PATCH_TREND_DYNAMIC.md` - Dynamic chart documentation
- `assets/patch_history/` - History data directory (if created)
- `assets/sample_patch_trend.png` - Sample chart image

**These files are inactive and can be deleted if desired.**

## To Re-enable the Chart

If you want to add the chart back in the future, you can:

1. Revert the changes to `create_template.py` (add chart placeholder)
2. Revert the changes to `api.py` (re-enable chart generation)
3. Regenerate the template

The chart generation modules are still available and functional.

## Report Generation

Reports will now generate **without** the patch trend chart:
- Patch counts will still be calculated and displayed in text
- No chart image will be created
- No history will be tracked
- Slide 14 remains text-only

All other functionality remains unchanged.
