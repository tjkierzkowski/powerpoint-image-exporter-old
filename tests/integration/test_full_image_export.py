import pytest
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx_export.pptx_export import PowerPointImageExporter


@pytest.fixture
def default_path(tmp_path: Path) -> Path:
    custom_path = tmp_path / "lecture_images"
    return custom_path


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


def test_powerpoint_images_are_the_same_as_the_number_in_the_file(tmp_path, actual_presentation_path):
    custom_path = tmp_path / "presentation_images"
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