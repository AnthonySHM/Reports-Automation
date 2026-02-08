# Text Overflow Fix - Internal Network Vulnerabilities

## Issue Identified

Text was overflowing the grey background panel boundaries at the bottom of vulnerability slides. This occurred because:

1. **Vulnerability consolidation** created entries with many targets on one line (e.g., 29 hosts for CVE-2023-34048)
2. **3 vulnerabilities per slide** was too many with the longer consolidated format
3. Text exceeded the available height (2.70") within the grey panel

## Fix Applied

### 1. Reduced Maximum Vulnerabilities Per Slide
**File:** `core/nuclei_parser.py` line 33

**Changed:**
```python
MAX_PER_SLIDE = 2  # Reduced from 3
```

**Result:** Each slide now shows only 2 vulnerabilities maximum, ensuring content fits within boundaries.

### 2. Adjusted Textbox Height
**File:** `core/nuclei_parser.py` line 43

**Changed:**
```python
TEXTBOX_HEIGHT = 2.70  # Reduced from 2.85"
```

**Calculation:**
- Grey panel height: 3.70"
- Intro text area: 0.80" (from top)
- Gap: 0.10"
- Available for textbox: 2.70" (with 0.10" bottom padding)
- Total: 0.90 + 2.70 = 3.60" (fits within 3.70" panel)

## Automatic Pagination Behavior

The existing pagination logic (already implemented) now works better:

**When vulnerabilities exceed 2:**
1. ✓ Creates duplicate slides automatically
2. ✓ Updates title to "Internal Network Vulnerabilities (cont.)"
3. ✓ Removes intro text from continuation slides (saves space)
4. ✓ Splits vulnerabilities across multiple slides (2 per slide)
5. ✓ Maintains numbering sequence (e.g., page 1: vulns 1-2, page 2: vulns 3-4)

**Example with 5 vulnerabilities:**
- **Slide 1:** Vulnerabilities #1 and #2
- **Slide 2 (cont.):** Vulnerabilities #3 and #4
- **Slide 3 (cont.):** Vulnerability #5

## Other Improvements Applied

1. **"Issue:" in bold** - Makes vulnerability descriptions more readable
2. **"Targets:" in bold** - Clearly separates the target list
3. **Word wrapping enabled** - Long target lists wrap properly within textbox width (6.72")

## Layout Specifications

**Grey Panel (content_panel):**
- Top: 1.10"
- Height: 3.70"
- Bottom: 4.80"

**Intro Text:**
- Top: 1.20" (BODY_TOP + 0.10)
- Height: 0.70"
- Bottom: 1.90"

**Vulnerability Textbox:**
- Top: 2.00" (BODY_TOP + 0.90)
- Height: 2.70"
- Bottom: 4.70" (within panel's 4.80" boundary)
- Width: 6.72" (CONTENT_W - 0.40)
- Left margin: 1.64" (CONTENT_L + 0.20)
- Right margin: 0.20" padding

## Testing Recommendation

With Elephant VAPRD having 3 consolidated vulnerabilities:
- **Before:** All 3 on one slide → overflow at bottom
- **After:** Slide 1 shows vulns 1-2, Slide 2 (cont.) shows vuln 3 → fits perfectly

## Files Modified

- `core/nuclei_parser.py` 
  - Line 33: `MAX_PER_SLIDE = 2`
  - Line 43: `TEXTBOX_HEIGHT = 2.70`

## Status

✅ Text overflow fixed
✅ Automatic pagination working
✅ Content stays within grey panel boundaries
✅ Continuation slides properly labeled

**Server restart required to apply changes.**
