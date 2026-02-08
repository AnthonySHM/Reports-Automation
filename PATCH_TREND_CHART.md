# Patch Trend Chart Feature

## Overview

The Required Software Patches slide (Slide 14) now includes a trend chart showing the progression of Microsoft KB patches and 3rd party software patches over time.

## Files Added

- **`core/patch_chart.py`**: Module for generating patch trend charts using PIL/Pillow
- **`assets/sample_patch_trend.png`**: Sample chart image (for template/testing)

## Chart Generator API

### `generate_patch_trend_chart()`

Generates a trend chart image showing patch counts over time.

```python
from core.patch_chart import generate_patch_trend_chart

# Example: Generate a chart for the last 3 months
dates = ["November 10", "December 8", "January 5"]
microsoft_counts = [6, 5, 4]
software_counts = [3, 3, 2]

chart_path = generate_patch_trend_chart(
    dates=dates,
    microsoft_counts=microsoft_counts,
    software_counts=software_counts,
    output_path="output/patch_trend.png",
    width=1500,   # pixels
    height=675,   # pixels
)
```

### Parameters

- **`dates`** (list[str]): Date labels for the X-axis (e.g., `["November 10", "December 8"]`)
- **`microsoft_counts`** (list[int]): Microsoft KB patch counts for each date
- **`software_counts`** (list[int]): 3rd party software patch counts for each date
- **`output_path`** (str | Path): Where to save the generated chart image
- **`width`** (int, optional): Image width in pixels (default: 1500)
- **`height`** (int, optional): Image height in pixels (default: 675)

### Returns

- **Path**: Path to the saved chart image

## Integration with Report Generation

### Automatic Chart Generation

To automatically generate and include patch trend charts in reports, you'll need to:

1. **Collect historical patch data** from your vulnerability management system
2. **Generate the chart** using `generate_patch_trend_chart()`
3. **Place the chart** on slide 14 using the existing image placement logic

### Example Integration in `api.py`

```python
from core.patch_chart import generate_patch_trend_chart
from core.drive_agent import place_ndr_images, NDRAsset

# Generate patch trend chart
chart_path = generate_patch_trend_chart(
    dates=patch_dates,
    microsoft_counts=ms_counts,
    software_counts=sw_counts,
    output_path=f"assets/ndr_cache/{client_slug}/patch_trend.png"
)

# Create an NDRAsset for the chart
patch_chart_asset = NDRAsset(
    local_path=chart_path,
    slide_index=13,  # Slide 14 (0-indexed)
    shape_name="patches_trend_chart",
    original_filename="patch_trend.png"
)

# Add to assets list for placement
ndr_assets.append(patch_chart_asset)
```

## Chart Styling

The chart uses the following color scheme to match the reference design:

- **Microsoft KB Patches**: Red line (`#E74C3C`, RGB: 231, 76, 60)
- **3rd Party Software Patches**: Green line (`#52BE80`, RGB: 82, 190, 128)
- **Grid lines**: Light gray (`#D5D8DC`, RGB: 213, 216, 220)
- **Text**: Dark gray (`#2C3E50`, RGB: 44, 62, 80)
- **Background**: White (`#FFFFFF`)

## Chart Features

- **Trend lines** with circular markers at each data point
- **Horizontal grid lines** for easy value reading
- **Y-axis** shows even numbers only (0, 2, 4, 6, 8, 10...)
- **Legend** identifying Microsoft KB vs 3rd Party patches
- **Title** "Missing Patch Quantity"
- **Axis labels** "Date" (X) and "# of Patches" (Y)

## Template Updates

The `slide_required_patches()` function in `create_template.py` has been updated to include:

1. Reduced text area to make room for the chart
2. Chart title: "Missing Patch Quantity"
3. Placeholder area named `patches_trend_chart` for the chart image

## Data Source Recommendations

To populate the chart with real data, you should:

1. **Track patch counts over time** (recommended: monthly or bi-weekly)
2. **Store historical data** in a database or CSV file
3. **Query the last 3-6 data points** for the trend chart
4. **Generate the chart** before creating the report

### Example Data Structure

```python
# Sample historical patch data
patch_history = [
    {"date": "2025-11-10", "microsoft": 6, "software": 3},
    {"date": "2025-12-08", "microsoft": 5, "software": 3},
    {"date": "2026-01-05", "microsoft": 4, "software": 2},
]

# Extract for chart
dates = [entry["date"] for entry in patch_history]
ms_counts = [entry["microsoft"] for entry in patch_history]
sw_counts = [entry["software"] for entry in patch_history]
```

## Testing

To test the chart generator:

```bash
# Generate a sample chart
python -m core.patch_chart
# Output: output/sample_patch_trend.png
```

## Notes

- The chart generator uses **PIL/Pillow** (already in requirements.txt)
- **No matplotlib dependency** required
- Chart images are typically **~50-100 KB** in size
- Charts are **optimized for PowerPoint** (1500x675 pixels @ 96 DPI)
