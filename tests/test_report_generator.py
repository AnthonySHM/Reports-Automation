"""
Integration tests for the ReportGenerator pipeline.

Creates temporary images and manifests to validate the full flow:
manifest → image resolution → transform → PPTX output → audit trail.
"""

import json
import tempfile
from pathlib import Path

import pytest
from PIL import Image

from core.report_generator import (
    ReportGenerator,
    ImageSourceError,
    SlideLayoutError,
)


@pytest.fixture
def workspace(tmp_path):
    """Create a temporary workspace with sample images and manifest."""
    assets = tmp_path / "assets"
    assets.mkdir()
    manifests = tmp_path / "manifests"
    manifests.mkdir()
    output = tmp_path / "output"
    output.mkdir()

    # Create test images of known sizes
    images = {
        "wide.png": (1920, 1080),
        "tall.png": (1080, 1920),
        "square.png": (1000, 1000),
    }
    for name, (w, h) in images.items():
        img = Image.new("RGB", (w, h), color="blue")
        img.save(str(assets / name))

    # Create manifest
    manifest = manifests / "test.csv"
    manifest.write_text(
        "slide_index,placeholder_id,image_path,box_x_emu,box_y_emu,box_w_emu,box_h_emu,label\n"
        "0,-1,wide.png,457200,914400,8229600,4572000,Wide Dashboard\n"
        "1,-1,tall.png,457200,914400,4114800,4572000,Tall Chart\n"
        "2,-1,square.png,457200,914400,8229600,4572000,Square Logo\n",
        encoding="utf-8",
    )

    return {
        "root": tmp_path,
        "assets": assets,
        "manifests": manifests,
        "output": output,
        "manifest_path": manifest,
    }


class TestReportGeneratorPipeline:
    def test_full_generation(self, workspace):
        gen = ReportGenerator(
            template_path=None,
            output_path=workspace["output"] / "report.pptx",
            assets_dir=workspace["assets"],
        )
        result = gen.generate(workspace["manifest_path"])

        assert result.exists()
        assert result.suffix == ".pptx"

        # Audit trail written
        audit_path = result.with_suffix(".audit.json")
        assert audit_path.exists()

        audit = json.loads(audit_path.read_text(encoding="utf-8"))
        assert audit["summary"]["total"] == 3
        assert audit["summary"]["ok"] == 3
        assert audit["summary"]["errors"] == 0

    def test_missing_image_logged_not_fatal(self, workspace):
        """A missing image should produce an error record, not crash."""
        manifest = workspace["manifests"] / "bad_image.csv"
        manifest.write_text(
            "slide_index,placeholder_id,image_path,box_x_emu,box_y_emu,box_w_emu,box_h_emu,label\n"
            "0,-1,nonexistent.png,457200,914400,8229600,4572000,Missing Image\n",
            encoding="utf-8",
        )

        gen = ReportGenerator(
            template_path=None,
            output_path=workspace["output"] / "report_bad.pptx",
            assets_dir=workspace["assets"],
        )
        result = gen.generate(manifest)

        assert result.exists()
        assert len(gen.audit_trail) == 1
        assert gen.audit_trail[0].status == "error"
        assert "not found" in gen.audit_trail[0].error_message.lower()

    def test_audit_contains_transform_details(self, workspace):
        gen = ReportGenerator(
            template_path=None,
            output_path=workspace["output"] / "report_audit.pptx",
            assets_dir=workspace["assets"],
        )
        gen.generate(workspace["manifest_path"])

        ok_records = [r for r in gen.audit_trail if r.status == "ok"]
        assert len(ok_records) == 3

        for record in ok_records:
            audit = record.transform.to_audit_dict()
            assert "scale_factor" in audit
            assert "affine_matrix" in audit
            assert audit["scale_factor"] > 0


class TestManifestLoading:
    def test_json_manifest(self, workspace):
        manifest = workspace["manifests"] / "test.json"
        manifest.write_text(json.dumps([
            {
                "slide_index": 0,
                "placeholder_id": -1,
                "image_path": "wide.png",
                "box_x_emu": 457200,
                "box_y_emu": 914400,
                "box_w_emu": 8229600,
                "box_h_emu": 4572000,
                "label": "JSON Wide Dashboard",
            }
        ]), encoding="utf-8")

        gen = ReportGenerator(
            template_path=None,
            output_path=workspace["output"] / "json_report.pptx",
            assets_dir=workspace["assets"],
        )
        result = gen.generate(manifest)
        assert result.exists()
        assert gen.audit_trail[0].status == "ok"

    def test_missing_columns_raises(self, workspace):
        manifest = workspace["manifests"] / "bad.csv"
        manifest.write_text("col_a,col_b\n1,2\n", encoding="utf-8")

        gen = ReportGenerator(
            template_path=None,
            output_path=workspace["output"] / "bad.pptx",
            assets_dir=workspace["assets"],
        )
        with pytest.raises(ValueError, match="missing columns"):
            gen.generate(manifest)
