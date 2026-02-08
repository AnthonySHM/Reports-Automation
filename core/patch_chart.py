"""
Patch Trend Chart Generator

Generates a trend chart showing Microsoft KB patches and 3rd party software patches
over time using PIL/Pillow (no matplotlib required).
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from PIL import Image, ImageDraw, ImageFont

logger = logging.getLogger("shm.patch_chart")

# Chart styling to match reference
CHART_COLORS = {
    'microsoft': (231, 76, 60),      # Red for Microsoft KB patches
    'software': (82, 190, 128),      # Green for 3rd party software
    'grid': (213, 216, 220),         # Light gray for grid lines
    'text': (44, 62, 80),            # Dark text
    'bg': (255, 255, 255),           # White background
    'axis': (0, 0, 0),               # Black for axes
}


def generate_patch_trend_chart(
    dates: list[str],
    microsoft_counts: list[int],
    software_counts: list[int],
    output_path: str | Path,
    *,
    width: int = 1500,
    height: int = 675,
) -> Path:
    """Generate a patch trend chart image using PIL.
    
    Parameters
    ----------
    dates : list[str]
        List of date strings (e.g., ["November 10", "December 8", "January 5"])
    microsoft_counts : list[int]
        Microsoft KB patch counts for each date
    software_counts : list[int]
        3rd party software patch counts for each date
    output_path : str | Path
        Path where the chart image will be saved
    width : int
        Image width in pixels
    height : int
        Image height in pixels
    
    Returns
    -------
    Path
        Path to the saved chart image
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Create image
    img = Image.new('RGB', (width, height), CHART_COLORS['bg'])
    draw = ImageDraw.Draw(img)
    
    # Try to use a nice font, fallback to default
    try:
        title_font = ImageFont.truetype("arial.ttf", 42)
        label_font = ImageFont.truetype("arial.ttf", 32)
        tick_font = ImageFont.truetype("arial.ttf", 28)
        legend_font = ImageFont.truetype("arial.ttf", 28)
    except:
        title_font = ImageFont.load_default()
        label_font = ImageFont.load_default()
        tick_font = ImageFont.load_default()
        legend_font = ImageFont.load_default()
    
    # Chart margins (increased for better spacing)
    margin_top = 120
    margin_bottom = 150
    margin_left = 140
    margin_right = 100
    
    # Plot area
    plot_left = margin_left
    plot_right = width - margin_right
    plot_top = margin_top
    plot_bottom = height - margin_bottom
    plot_width = plot_right - plot_left
    plot_height = plot_bottom - plot_top
    
    # Title
    title = "Missing Patch Quantity"
    title_bbox = draw.textbbox((0, 0), title, font=title_font)
    title_width = title_bbox[2] - title_bbox[0]
    title_x = (width - title_width) // 2
    draw.text((title_x, 25), title, fill=CHART_COLORS['text'], font=title_font)
    
    # Determine Y-axis range
    max_val = max(max(microsoft_counts), max(software_counts))
    y_max = ((max_val + 1) // 2 + 1) * 2  # Round up to next even number
    y_ticks = list(range(0, y_max + 1, 2))
    
    # Draw Y-axis gridlines and labels
    for y_val in y_ticks:
        y_pos = plot_bottom - (y_val / y_max) * plot_height
        
        # Gridline
        draw.line(
            [(plot_left, y_pos), (plot_right, y_pos)],
            fill=CHART_COLORS['grid'], width=2
        )
        
        # Y-axis tick label (adjusted positioning for better spacing)
        label = str(y_val)
        label_bbox = draw.textbbox((0, 0), label, font=tick_font)
        label_height = label_bbox[3] - label_bbox[1]
        draw.text((margin_left - 80, y_pos - label_height // 2), label, fill=CHART_COLORS['text'], font=tick_font)
    
    # Draw axes
    draw.line([(plot_left, plot_top), (plot_left, plot_bottom)], 
              fill=CHART_COLORS['axis'], width=3)
    draw.line([(plot_left, plot_bottom), (plot_right, plot_bottom)], 
              fill=CHART_COLORS['axis'], width=3)
    
    # Calculate X positions
    n_points = len(dates)
    if n_points > 1:
        x_step = plot_width / (n_points - 1)
        x_positions = [plot_left + i * x_step for i in range(n_points)]
    else:
        x_positions = [plot_left + plot_width / 2]
    
    # Helper function to convert data value to Y pixel position
    def data_to_y(value):
        return plot_bottom - (value / y_max) * plot_height
    
    # Draw lines
    # Microsoft KB patches (red)
    ms_points = [(x_positions[i], data_to_y(microsoft_counts[i])) 
                 for i in range(n_points)]
    for i in range(len(ms_points) - 1):
        draw.line([ms_points[i], ms_points[i + 1]], 
                  fill=CHART_COLORS['microsoft'], width=5)
    
    # 3rd Party Software patches (green)
    sw_points = [(x_positions[i], data_to_y(software_counts[i])) 
                 for i in range(n_points)]
    for i in range(len(sw_points) - 1):
        draw.line([sw_points[i], sw_points[i + 1]], 
                  fill=CHART_COLORS['software'], width=5)
    
    # Draw markers
    marker_radius = 12
    for i in range(n_points):
        # Microsoft marker
        x, y = ms_points[i]
        draw.ellipse(
            [(x - marker_radius, y - marker_radius), 
             (x + marker_radius, y + marker_radius)],
            fill=CHART_COLORS['microsoft']
        )
        
        # Software marker
        x, y = sw_points[i]
        draw.ellipse(
            [(x - marker_radius, y - marker_radius), 
             (x + marker_radius, y + marker_radius)],
            fill=CHART_COLORS['software']
        )
    
    # X-axis labels (dates) - improved spacing
    for i, date in enumerate(dates):
        date_bbox = draw.textbbox((0, 0), date, font=tick_font)
        date_width = date_bbox[2] - date_bbox[0]
        x_center = x_positions[i]
        draw.text((x_center - date_width // 2, plot_bottom + 25), 
                  date, fill=CHART_COLORS['text'], font=tick_font)
    
    # Axis labels
    # Y-axis label (vertically aligned, better positioning)
    y_label = "# of Patches"
    y_label_bbox = draw.textbbox((0, 0), y_label, font=label_font)
    y_label_height = y_label_bbox[3] - y_label_bbox[1]
    draw.text((15, plot_top + plot_height // 2 - y_label_height // 2), y_label, 
              fill=CHART_COLORS['text'], font=label_font)
    
    # X-axis label (improved spacing)
    x_label = "Date"
    x_label_bbox = draw.textbbox((0, 0), x_label, font=label_font)
    x_label_width = x_label_bbox[2] - x_label_bbox[0]
    draw.text((plot_left + plot_width // 2 - x_label_width // 2, height - 85), 
              x_label, fill=CHART_COLORS['text'], font=label_font)
    
    # Legend (improved spacing and positioning)
    legend_y = height - 50
    marker_size = 18
    
    # Calculate total legend width to center it properly
    ms_legend_text = "Microsoft KB Patches"
    sw_legend_text = "3rd Party Software Patches"
    
    ms_text_bbox = draw.textbbox((0, 0), ms_legend_text, font=legend_font)
    ms_text_width = ms_text_bbox[2] - ms_text_bbox[0]
    
    sw_text_bbox = draw.textbbox((0, 0), sw_legend_text, font=legend_font)
    sw_text_width = sw_text_bbox[2] - sw_text_bbox[0]
    
    # Spacing between legend items
    legend_spacing = 60
    
    # Total width of both legend items
    total_legend_width = marker_size + 15 + ms_text_width + legend_spacing + marker_size + 15 + sw_text_width
    
    # Start position to center the legend
    legend_start_x = (width - total_legend_width) // 2
    
    # Microsoft KB Patches legend
    ms_legend_x = legend_start_x
    draw.ellipse(
        [(ms_legend_x, legend_y - marker_size // 2), 
         (ms_legend_x + marker_size, legend_y + marker_size // 2)],
        fill=CHART_COLORS['microsoft']
    )
    draw.text((ms_legend_x + marker_size + 10, legend_y - 12), 
              ms_legend_text, fill=CHART_COLORS['text'], font=legend_font)
    
    # 3rd Party Software legend
    sw_legend_x = ms_legend_x + marker_size + 15 + ms_text_width + legend_spacing
    draw.ellipse(
        [(sw_legend_x, legend_y - marker_size // 2), 
         (sw_legend_x + marker_size, legend_y + marker_size // 2)],
        fill=CHART_COLORS['software']
    )
    draw.text((sw_legend_x + marker_size + 10, legend_y - 12), 
              sw_legend_text, fill=CHART_COLORS['text'], font=legend_font)
    
    # Save image
    img.save(output_path, 'PNG', dpi=(150, 150))
    
    logger.info(
        "Generated patch trend chart → %s (%d dates, MS: %s, SW: %s)",
        output_path, len(dates), microsoft_counts, software_counts
    )
    
    return output_path


def generate_sample_patch_chart(output_path: str | Path) -> Path:
    """Generate a sample patch trend chart with placeholder data.
    
    This is useful for template generation and testing.
    """
    dates = ["November 10", "December 8", "January 5"]
    microsoft_counts = [6, 5, 4]
    software_counts = [3, 3, 2]
    
    return generate_patch_trend_chart(
        dates=dates,
        microsoft_counts=microsoft_counts,
        software_counts=software_counts,
        output_path=output_path,
    )


if __name__ == "__main__":
    # Test the chart generator
    output = Path("output/sample_patch_trend.png")
    generate_sample_patch_chart(output)
    print(f"Sample chart saved to: {output}")
