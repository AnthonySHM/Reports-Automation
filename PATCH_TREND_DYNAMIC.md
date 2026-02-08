# Dynamic Patch Trend Chart - Implementation Guide

## Overview

The patch trend chart on the Required Software Patches slide (Slide 14) now **automatically updates** with each report generation:

- **Last data point** = Current report period's end date
- **Line values** = Current counts of missing KB patches and software packages
- **Historical data** = Previous report periods (up to 3 data points shown)

## How It Works

### 1. Automatic History Tracking

When a report is generated:
1. Current patch counts are calculated from Drive CSVs
2. History manager loads previous data points for the client
3. Current period's data is added/updated in history
4. Chart is generated with the 3 most recent data points
5. Chart image is placed on slide 14

### 2. History Storage

Patch history is stored in JSON files:
- **Location**: `assets/patch_history/{client_slug}_patch_history.json`
- **Format**:
```json
[
  {
    "date": "2025-11-10",
    "microsoft_count": 6,
    "software_count": 3,
    "updated_at": "2025-11-15T10:30:00"
  },
  {
    "date": "2025-12-08",
    "microsoft_count": 5,
    "software_count": 3,
    "updated_at": "2025-12-10T14:20:00"
  },
  {
    "date": "2026-01-05",
    "microsoft_count": 4,
    "software_count": 2,
    "updated_at": "2026-01-08T09:15:00"
  }
]
```

### 3. Data Flow

```
Report Request
    ↓
Fetch remediation_plan CSV → Count unique Microsoft KBs
Fetch software CSV → Count unique software packages
    ↓
Generate Patch Trend Chart:
  1. Load history for client
  2. Add/update current period's counts
  3. Take last 3 data points
  4. Generate chart image
  5. Save updated history
    ↓
Place chart on slide 14
    ↓
Report Complete
```

## API Integration

### In `api.py`

The chart generation is integrated after patch counts are calculated:

```python
# After counting patches...
if kb_count is not None or sw_count is not None:
    from core.patch_history import generate_patch_trend_chart_for_report
    
    chart_path = generate_patch_trend_chart_for_report(
        client_slug=client_name,
        end_date=end_date_formatted,  # "2026-01-05"
        microsoft_count=kb_count or 0,
        software_count=sw_count or 0,
        output_path=NDR_CACHE_DIR / client_name / "patch_trend.png",
        num_points=3,  # Show last 3 data points
    )
    
    # Create asset for placement
    patch_chart_asset = NDRAsset(
        local_path=chart_path,
        slide_index=13,  # Slide 14 (0-indexed)
        shape_name="patches_trend_chart",
        original_filename="patch_trend.png",
    )
```

### Chart Placement

**Single-Sensor Reports:**
- Chart added to `assets` list
- Placed with NDR images on slide 14

**Multi-Sensor Reports:**
- Chart placed separately (non-sensor-specific)
- Slide index adjusted: `5 + (num_sensors × 10)`

## Chart Features

### Visual Elements

- **Dual trend lines**:
  - Red line: Microsoft KB patches
  - Green line: 3rd party software patches
- **Data points**: Circular markers at each date
- **Grid lines**: Horizontal lines for easy reading
- **Legend**: Bottom center, identifies both data series
- **Title**: "Missing Patch Quantity"
- **Axes**: "Date" (X) and "# of Patches" (Y)

### Date Formatting

Dates are automatically formatted for display:
- **Input**: `"2026-01-05"` (YYYY-MM-DD)
- **Output**: `"January 5"` (Month Day)

### Historical Data Management

- **Maximum history**: 6 data points stored per client
- **Chart displays**: Last 3 data points
- **Auto-trimming**: Oldest entries removed when limit exceeded
- **Duplicate prevention**: Same date updates existing entry

## Usage Examples

### Example 1: First Report for a Client

```python
# No previous history exists
generate_patch_trend_chart_for_report(
    client_slug="elephant",
    end_date="2026-01-05",
    microsoft_count=4,
    software_count=2,
    output_path="assets/ndr_cache/elephant/patch_trend.png"
)

# Result: Chart shows only 1 data point
# History file created with 1 entry
```

### Example 2: Third Report

```python
# History has 2 previous entries (Nov, Dec)
generate_patch_trend_chart_for_report(
    client_slug="elephant",
    end_date="2026-01-05",
    microsoft_count=4,
    software_count=2,
    output_path="assets/ndr_cache/elephant/patch_trend.png"
)

# Result: Chart shows 3 data points (Nov, Dec, Jan)
# Trend lines show progression over time
```

### Example 3: Updating Same Period

```python
# Report regenerated for same period
generate_patch_trend_chart_for_report(
    client_slug="elephant",
    end_date="2026-01-05",
    microsoft_count=3,  # Updated count
    software_count=2,
    output_path="assets/ndr_cache/elephant/patch_trend.png"
)

# Result: January entry updated (not duplicated)
# Chart reflects new count
```

## Manual History Management

### View History

```python
from core.patch_history import PatchHistoryManager

manager = PatchHistoryManager()
history = manager.load_history("elephant")

for entry in history:
    print(f"{entry['date']}: MS={entry['microsoft_count']}, SW={entry['software_count']}")
```

### Add Historical Data

```python
manager = PatchHistoryManager()

# Backfill historical data
historical_data = [
    ("2025-11-10", 6, 3),
    ("2025-12-08", 5, 3),
]

for date, ms, sw in historical_data:
    manager.add_entry("elephant", date, ms, sw)
```

### Clear History

```python
import os
from pathlib import Path

# Remove history file
history_file = Path("assets/patch_history/elephant_patch_history.json")
if history_file.exists():
    os.remove(history_file)
```

## Files Added/Modified

### New Files
- **`core/patch_history.py`**: History management and chart generation
- **`assets/patch_history/`**: Directory for history JSON files (created automatically)

### Modified Files
- **`api.py`**: Integrated chart generation into report workflow
- **`create_template.py`**: Added chart placeholder to slide 14
- **`core/patch_chart.py`**: Chart generator with improved spacing

## Configuration

### Number of Data Points

To show more or fewer data points:

```python
# In api.py, change num_points parameter:
generate_patch_trend_chart_for_report(
    ...
    num_points=5,  # Show last 5 data points instead of 3
)
```

### Maximum History Storage

To store more historical data:

```python
# In patch_history.py:
manager.add_entry(
    client_slug="elephant",
    end_date="2026-01-05",
    microsoft_count=4,
    software_count=2,
    max_history=12,  # Store up to 12 months instead of 6
)
```

## Troubleshooting

### Chart Not Appearing

**Check:**
1. Patch counts calculated successfully?
   - Look for `kb_count` and `sw_count` in logs
2. Chart generation succeeded?
   - Look for "Generated patch trend chart for..." in logs
3. History file created?
   - Check `assets/patch_history/{client}_patch_history.json`

### Wrong Data on Chart

**Check:**
1. History file content
2. End date format (must be YYYY-MM-DD)
3. Logs for "Added history entry" or "Updated history entry"

### Chart Shows Only One Point

**Expected** for first report of a client. History builds over time.

To backfill historical data, use the manual history management methods above.

## Best Practices

1. **Consistent end dates**: Use the same day of month for reports (e.g., 5th of each month)
2. **Regular reporting**: Generate reports at consistent intervals
3. **Backup history files**: Include `assets/patch_history/` in backups
4. **Monitor logs**: Check for chart generation success/errors

## Future Enhancements

Possible improvements:
- [ ] Export history to CSV for analysis
- [ ] Configurable date formats
- [ ] Trend prediction (dotted line for next period)
- [ ] Alert if patches increase instead of decrease
- [ ] Historical data import from external systems
