"""
Tests for Algorithm 1: Intelligent Slide Layout Engine

Validates the mathematical model:
  - Aspect ratio computation
  - Scaling factor derivation (width-constrained vs height-constrained)
  - Centering translation
  - Affine matrix construction
  - Edge cases (square images, extreme ratios, zero dimensions)
"""

import math
import pytest
import numpy as np

from core.transform import (
    Rect,
    TransformResult,
    compute_contain_transform,
    apply_transform_to_point,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
EMU_PER_INCH = 914400


def inches(n: float) -> int:
    return int(n * EMU_PER_INCH)


# ---------------------------------------------------------------------------
# Aspect Ratio Tests
# ---------------------------------------------------------------------------
class TestRect:
    def test_aspect_ratio_landscape(self):
        r = Rect(0, 0, 1600, 900)
        assert abs(r.aspect_ratio - 16 / 9) < 1e-6

    def test_aspect_ratio_portrait(self):
        r = Rect(0, 0, 900, 1600)
        assert r.aspect_ratio < 1.0

    def test_aspect_ratio_square(self):
        r = Rect(0, 0, 1000, 1000)
        assert r.aspect_ratio == 1.0

    def test_zero_height_raises(self):
        r = Rect(0, 0, 1000, 0)
        with pytest.raises(ValueError, match="zero height"):
            _ = r.aspect_ratio

    def test_negative_dimension_raises(self):
        with pytest.raises(ValueError):
            Rect(0, 0, -100, 200)


# ---------------------------------------------------------------------------
# Contain Transform — Width-Constrained
# ---------------------------------------------------------------------------
class TestContainWidthConstrained:
    """Image is wider than the box → scale by width."""

    def test_wide_image_scales_by_width(self):
        source = Rect(0, 0, inches(10), inches(5))       # 2:1
        box = Rect(inches(1), inches(1), inches(8), inches(6))  # 4:3

        result = compute_contain_transform(source, box)

        assert result.scale_axis == "width"
        # Scaled width should equal box width
        assert result.dest_width == box.width
        # Scaled height < box height (letterboxed)
        assert result.dest_height < box.height

    def test_wide_image_centered_vertically(self):
        source = Rect(0, 0, 2000, 1000)
        box = Rect(100, 100, 1000, 1000)

        result = compute_contain_transform(source, box)

        assert result.scale_axis == "width"
        # Horizontal: should fill box width exactly
        assert result.dest_x == box.x
        # Vertical: should be centered — offset = (1000 - 500) / 2 = 250
        expected_ty = (box.height - result.dest_height) // 2
        assert result.ty == expected_ty


# ---------------------------------------------------------------------------
# Contain Transform — Height-Constrained
# ---------------------------------------------------------------------------
class TestContainHeightConstrained:
    """Image is taller than the box → scale by height."""

    def test_tall_image_scales_by_height(self):
        source = Rect(0, 0, inches(5), inches(10))       # 1:2
        box = Rect(inches(1), inches(1), inches(8), inches(6))  # 4:3

        result = compute_contain_transform(source, box)

        assert result.scale_axis == "height"
        assert result.dest_height == box.height
        assert result.dest_width < box.width

    def test_tall_image_centered_horizontally(self):
        source = Rect(0, 0, 1000, 2000)
        box = Rect(100, 100, 1000, 1000)

        result = compute_contain_transform(source, box)

        assert result.scale_axis == "height"
        expected_tx = (box.width - result.dest_width) // 2
        assert result.tx == expected_tx


# ---------------------------------------------------------------------------
# Contain Transform — Perfect Fit / Square
# ---------------------------------------------------------------------------
class TestContainPerfectFit:
    def test_same_aspect_ratio_fills_exactly(self):
        source = Rect(0, 0, 1600, 900)
        box = Rect(100, 100, 3200, 1800)

        result = compute_contain_transform(source, box)

        assert result.dest_width == box.width
        assert result.dest_height == box.height
        assert result.tx == 0
        assert result.ty == 0

    def test_square_into_square(self):
        source = Rect(0, 0, 500, 500)
        box = Rect(0, 0, 1000, 1000)

        result = compute_contain_transform(source, box)

        assert result.dest_width == 1000
        assert result.dest_height == 1000
        assert result.scale_factor == 2.0


# ---------------------------------------------------------------------------
# Affine Matrix Verification
# ---------------------------------------------------------------------------
class TestAffineMatrix:
    def test_matrix_shape(self):
        source = Rect(0, 0, 1000, 500)
        box = Rect(200, 300, 800, 600)

        result = compute_contain_transform(source, box)

        assert result.affine_matrix.shape == (3, 3)
        # Last row must be [0, 0, 1]
        np.testing.assert_array_equal(result.affine_matrix[2], [0, 0, 1])

    def test_origin_maps_to_dest_origin(self):
        """Applying the affine to (0, 0) should give (dest_x, dest_y)."""
        source = Rect(0, 0, 2000, 1000)
        box = Rect(500, 500, 1000, 1000)

        result = compute_contain_transform(source, box)
        x, y = apply_transform_to_point(result.affine_matrix, 0, 0)

        assert abs(x - result.dest_x) < 1
        assert abs(y - result.dest_y) < 1

    def test_far_corner_maps_correctly(self):
        """Applying the affine to (w_img, h_img) should give (dest_x + dest_w, dest_y + dest_h)."""
        source = Rect(0, 0, 2000, 1000)
        box = Rect(500, 500, 1000, 1000)

        result = compute_contain_transform(source, box)
        x, y = apply_transform_to_point(result.affine_matrix, source.width, source.height)

        assert abs(x - (result.dest_x + result.dest_width)) < 2
        assert abs(y - (result.dest_y + result.dest_height)) < 2


# ---------------------------------------------------------------------------
# Edge Cases
# ---------------------------------------------------------------------------
class TestEdgeCases:
    def test_zero_source_raises(self):
        source = Rect(0, 0, 0, 500)
        box = Rect(0, 0, 1000, 1000)
        with pytest.raises(ValueError, match="zero dimension"):
            compute_contain_transform(source, box)

    def test_extreme_aspect_ratio(self):
        """Very wide panoramic image into a square box."""
        source = Rect(0, 0, 100000, 100)
        box = Rect(0, 0, 1000, 1000)

        result = compute_contain_transform(source, box)

        assert result.scale_axis == "width"
        assert result.dest_width == 1000
        assert result.dest_height == 1  # nearly flat

    def test_audit_dict_serialisable(self):
        source = Rect(0, 0, 1000, 500)
        box = Rect(100, 100, 800, 600)

        result = compute_contain_transform(source, box)
        audit = result.to_audit_dict()

        assert "scale_factor" in audit
        assert "affine_matrix" in audit
        assert isinstance(audit["affine_matrix"], list)
        assert len(audit["affine_matrix"]) == 3
