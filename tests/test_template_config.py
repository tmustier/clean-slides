from pathlib import Path

import pytest
import yaml

from clean_slides.template_config import load_template_config


def _config_path() -> Path:
    return Path(__file__).resolve().parents[1] / "clean_slides" / "template-config.yaml"


def test_default_template_config_valid():
    config = load_template_config(_config_path())
    assert "midnight" in config.section("colors")
    assert "slide_width_emu" in config.section("layout")


def test_missing_required_key_raises(tmp_path: Path):
    data = yaml.safe_load(_config_path().read_text())
    data["colors"].pop("midnight", None)

    bad_path = tmp_path / "bad-template.yaml"
    bad_path.write_text(yaml.safe_dump(data))

    with pytest.raises(ValueError) as excinfo:
        load_template_config(bad_path)

    message = str(excinfo.value)
    assert "Missing colors.midnight" in message
