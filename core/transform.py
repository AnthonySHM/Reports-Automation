"""
Algorithm 1: Intelligent Slide Layout Engine — Affine Transformation Core

Implements 2D affine transformations in homogeneous coordinates to map
source cybersecurity dashboard images into destination PowerPoint placeholders
using 'contain' logic (scale-to-fit without distortion).

Mathematical Model
------------------
    x' = A·x + b

Where:
    A = [[s, 0],    (scaling matrix, uniform scale factor s)
         [0, s]]

    b = [[tx],       (translation vector for centering)
         [ty]]

In homogeneous coordinates:
    [x']   [s  0  tx] [x]
    [y'] = [0  s  ty] [y]
    [1 ]   [0  0   1] [1]

Author: SHM Automation Engine
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Optional

import numpy as np

logger = logging.getLogger("shm.transform")


# ---------------------------------------------------------------------------
# Data Models
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Rect:
    """Axis-aligned rectangle defined by origin + dimensions (in EMU)."""
    x: int
    y: int
    width: int
    height: int

    @property
    def aspect_ratio(self) -> float:
        if self.height == 0:
            raise ValueError("Rectangle has zero height — cannot compute aspect ratio.")
        return self.width / self.height

    def __post_init__(self):
        if self.width < 0 or self.height < 0:
            raise ValueError(f"Dimensions must be non-negative: w={self.width}, h={self.height}")


@dataclass(frozen=True)
class TransformResult:
    """Immutable record of a single affine placement computation.

    Serves as the permanent, auditable record of how raw telemetry
    (source image dimensions) was transformed into the final presentation
    coordinates.
    """
    # Input parameters
    source_width: int
    source_height: int
    box_x: int
    box_y: int
    box_width: int
    box_height: int

    # Derived values
    r_content: float          # source aspect ratio
    r_box: float              # placeholder aspect ratio
    scale_factor: float       # uniform scale  s
    scale_axis: str           # "width" or "height"

    # Affine components
    affine_matrix: np.ndarray  # 3×3 homogeneous transform

    # Final placement (EMU)
    dest_x: int
    dest_y: int
    dest_width: int
    dest_height: int

    # Translation offsets used for centering
    tx: int
    ty: int

    class Config:
        arbitrary_types_allowed = True

    def to_audit_dict(self) -> dict:
        """Serialisable dictionary for logging / JSON export."""
        return {
            "source": {"width": self.source_width, "height": self.source_height},
            "placeholder": {
                "x": self.box_x, "y": self.box_y,
                "width": self.box_width, "height": self.box_height,
            },
            "r_content": round(self.r_content, 6),
            "r_box": round(self.r_box, 6),
            "scale_factor": round(self.scale_factor, 6),
            "scale_axis": self.scale_axis,
            "affine_matrix": self.affine_matrix.tolist(),
            "destination": {
                "x": self.dest_x, "y": self.dest_y,
                "width": self.dest_width, "height": self.dest_height,
            },
            "translation": {"tx": self.tx, "ty": self.ty},
        }


# ---------------------------------------------------------------------------
# Algorithm 1 — Contain Transform
# ---------------------------------------------------------------------------
def compute_contain_transform(
    source: Rect,
    placeholder: Rect,
) -> TransformResult:
    """Compute the 'contain' affine transform for placing *source* inside *placeholder*.

    Steps
    -----
    1. Compute aspect ratios:
        R_content = w_img / h_img
        R_box     = w_box / h_box

    2. Derive uniform scaling factor *s*:
        if R_content > R_box   →  width-constrained   →  s = w_box / w_img
        else                   →  height-constrained  →  s = h_box / h_img

    3. Derive translation (tx, ty) to centre content in placeholder:
        scaled_w = s * w_img
        scaled_h = s * h_img
        tx = box_x + (w_box - scaled_w) / 2
        ty = box_y + (h_box - scaled_h) / 2

    4. Build the 3×3 affine matrix in homogeneous coordinates.

    Parameters
    ----------
    source : Rect
        Dimensions of the source image (x, y ignored; only w, h used).
    placeholder : Rect
        Position and dimensions of the destination placeholder in the slide.

    Returns
    -------
    TransformResult
        Full auditable record of the transformation.
    """
    if source.width == 0 or source.height == 0:
        raise ValueError(f"Source image has zero dimension: {source.width}×{source.height}")

    # Step 1 — Aspect ratios
    r_content: float = source.width / source.height
    r_box: float = placeholder.aspect_ratio

    # Step 2 — Scaling factor
    if r_content > r_box:
        # Width-constrained: image is wider relative to box
        s = placeholder.width / source.width
        scale_axis = "width"
    else:
        # Height-constrained: image is taller relative to box
        s = placeholder.height / source.height
        scale_axis = "height"

    # Step 3 — Scaled dimensions & centering translation
    scaled_w = int(round(s * source.width))
    scaled_h = int(round(s * source.height))

    tx = placeholder.x + (placeholder.width - scaled_w) // 2
    ty = placeholder.y + (placeholder.height - scaled_h) // 2

    # Step 4 — Affine matrix (homogeneous coordinates)
    #   [s  0  tx]
    #   [0  s  ty]
    #   [0  0   1]
    affine = np.array([
        [s,  0.0, tx],
        [0.0, s,  ty],
        [0.0, 0.0, 1.0],
    ], dtype=np.float64)

    result = TransformResult(
        source_width=source.width,
        source_height=source.height,
        box_x=placeholder.x,
        box_y=placeholder.y,
        box_width=placeholder.width,
        box_height=placeholder.height,
        r_content=r_content,
        r_box=r_box,
        scale_factor=s,
        scale_axis=scale_axis,
        affine_matrix=affine,
        dest_x=tx,
        dest_y=ty,
        dest_width=scaled_w,
        dest_height=scaled_h,
        tx=tx - placeholder.x,  # offset relative to box origin
        ty=ty - placeholder.y,
    )

    logger.debug(
        "Contain transform: %d×%d → %d×%d  (s=%.4f, axis=%s, tx=%d, ty=%d)",
        source.width, source.height, scaled_w, scaled_h,
        s, scale_axis, result.tx, result.ty,
    )

    return result


def apply_transform_to_point(affine: np.ndarray, x: float, y: float) -> tuple[float, float]:
    """Apply a 3×3 affine matrix to a 2D point (utility for verification)."""
    point = np.array([x, y, 1.0])
    result = affine @ point
    return float(result[0]), float(result[1])
