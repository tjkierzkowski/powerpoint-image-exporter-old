import pytest
from pathlib import Path
from pptx import Presentation

# import pptx_export.pptx_export


@pytest.fixture
def fake_path(tmp_path):
    fake_file_path = tmp_path / "presentations"
    return fake_file_path


@pytest.fixture
def custom_path(tmp_path: Path) -> Path:
    custom_path = tmp_path / "presentation_images"
    return custom_path


@pytest.fixture
def default_path(tmp_path: Path) -> Path:
    custom_path = tmp_path / "lecture_images"
    return custom_path


@pytest.fixture
def fake_file():
    return 'a_valid_file.pptx'


@pytest.fixture
def minimal_pres(tmp_path):
    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "Hello, World!"
    subtitle.text = "python-pptx was here!"
    minimal_file = tmp_path / 'minimal.pptx'
    prs.save(minimal_file)
    return minimal_file


@pytest.fixture
def valid_presentation_name(custom_path, fake_file):
    from pptx_export.pptx_export import PowerPointImageExporter
    ppt_stub_file = custom_path / fake_file
    pres = PowerPointImageExporter(ppt_stub_file)
    return pres
