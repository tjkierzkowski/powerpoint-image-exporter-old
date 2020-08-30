import pytest
from unittest.mock import Mock
from pptx import Presentation
from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE_TYPE

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


# TODO migrate this to an integration test

@pytest.fixture(scope="session")
def actual_presentation_path():
    """Load an actual .pptx file with images to test against"""
    project_name = 'powerpointImageExporter'
    presentation_under_test = 'stub_tester.pptx'
    project_dir = [project for project in Path.home().rglob(project_name) if project.is_dir()]
    if not project_dir:
        raise ValueError(f"Could not find project directory '{project_name}' from your user's home directory")
    project_root = project_dir[0]
    presentations = [pres for pres in project_root.rglob(presentation_under_test) if pres.is_file()]
    if not presentations:
        raise ValueError(f"Could not find the actual presentation '{presentation_under_test}' within the project "
                         f"directory {project_root.resolve()}")
    return str(presentations[0].resolve())


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


# TODO separate these out to integration tests

def test_powerpoint_images_are_the_same_as_the_number_in_the_file(custom_path, actual_presentation_path):
    pres_file = actual_presentation_path
    pptx_export = PowerPointImageExporter(pres_file)
    pres = Presentation(pres_file)
    result_pictures = [picture_info
                       for picture_info in pptx_export.iter_by_shape(pres, MSO_SHAPE_TYPE.PICTURE)
                       ]
    result_amount = len(result_pictures)
    pptx_export.create_directory_for_images(custom_path)
    expected_images = [shape
                       for slide in pres.slides
                       for shape in slide.shapes
                       if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert len(expected_images) == result_amount


def test_if_all_images_from_an_actual_presentation_are_extracted_to_directory(default_path, actual_presentation_path):
    real_presentation = actual_presentation_path
    pptx_export = PowerPointImageExporter(real_presentation)
    pptx_export.create_directory_for_images(default_path)
    pptx_export.copy_images_to_directory()
    result_pictures = [files
                       for files in default_path.iterdir()]
    expected_images = [shape
                       for slide in Presentation(real_presentation).slides
                       for shape in slide.shapes
                       if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert len(expected_images) == len(result_pictures) and len(expected_images) > 0
