import pytest
from pathlib import Path

from pptx_export.pptx_export import PowerPointImageExporter


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
def valid_presentation_name(custom_path):
    ppt_stub_file = custom_path / 'a_valid_file.pptx'
    pres = PowerPointImageExporter(ppt_stub_file)
    return pres


def test_if_file_is_none():
    with pytest.raises(ValueError):
        PowerPointImageExporter(None)


def test_if_file_is_not_a_file(fake_path):
    fake_path.mkdir()
    with pytest.raises(ValueError):
        PowerPointImageExporter(fake_path)


def test_if_file_is_not_pptx(fake_path):
    fake_path.mkdir()
    not_pptx_file = fake_path / 'fakefile.ppt'
    with pytest.raises(ValueError):
        PowerPointImageExporter(not_pptx_file)


def test_if_powerpoint_name_is_safely_created(tmp_path):
    expected = "Fall_2020_Block_1_Lab_Review_w_labels_1-3_and_osteo-1"
    pptx_file = tmp_path / 'Fall 2020 Block 1 Lab Review w labels 1-3 and osteo-1.pptx'
    ppie = PowerPointImageExporter(pptx_file)
    actual = ppie.safe_presentation_name
    assert expected == actual


def test_presentation_images_directory_is_already_created(custom_path, valid_presentation_name):
    custom_path.mkdir()
    assert valid_presentation_name.create_directory_for_images(custom_path) is None and custom_path.exists()


def test_presentation_images_directory_is_created(custom_path, valid_presentation_name):
    assert not custom_path.exists()
    valid_presentation_name.create_directory_for_images(custom_path)
    assert custom_path.exists()


def test_presentation_images_directory_throws_error_without_path(custom_path, valid_presentation_name):
    assert not custom_path.exists()
    with pytest.raises(ValueError):
        valid_presentation_name.create_directory_for_images()


def test_presentation_images_directory_will_not_be_overwritten(custom_path, valid_presentation_name):
    custom_path.mkdir()
    with pytest.raises(ValueError):
        valid_presentation_name.create_directory_for_images()


def test_meaningful_filename_output(valid_presentation_name):
    assert 'None_slide_None_image_None.None' == valid_presentation_name.meaningful_filename(None, None, None, None)


