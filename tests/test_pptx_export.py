from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pytest

from pptx_export.pptx_export import PowerPointImageExporter


def test_if_file_is_none():
    with pytest.raises(ValueError):
        PowerPointImageExporter(None)


def test_if_file_is_not_a_file(fake_path):
    fake_path.mkdir()
    with pytest.raises(ValueError):
        PowerPointImageExporter(fake_path)


def test_if_file_is_not_pptx(fake_path):
    fake_path.mkdir()
    not_pptx_file = fake_path / "fakefile.ppt"
    with pytest.raises(ValueError):
        PowerPointImageExporter(not_pptx_file)


def test_if_powerpoint_name_is_safely_created(tmp_path):
    expected = "Fall_2020_Block_1_Lab_Review_w_labels_1-3_and_osteo-1"
    pptx_file = tmp_path / "Fall 2020 Block 1 Lab Review w labels 1-3 and osteo-1.pptx"
    ppie = PowerPointImageExporter(pptx_file)
    actual = ppie.safe_presentation_name
    assert expected == actual


def test_presentation_images_directory_is_already_created(
    custom_path, valid_presentation_name
):
    custom_path.mkdir()
    assert (
        valid_presentation_name.create_directory_for_images(custom_path) is None
        and custom_path.exists()
    )


def test_presentation_images_directory_is_created(tmp_path, valid_presentation_name):
    assert len([p for p in tmp_path.iterdir()]) < 1
    valid_presentation_name.create_directory_for_images(str(tmp_path / "fake_output"))
    assert len([p for p in tmp_path.iterdir()]) != 0


def test_presentation_images_directory_throws_error_without_path(
    custom_path, valid_presentation_name
):
    assert not custom_path.exists()
    with pytest.raises(ValueError):
        valid_presentation_name.create_directory_for_images()


def test_presentation_images_directory_will_not_be_overwritten(
    custom_path, valid_presentation_name
):
    custom_path.mkdir()
    with pytest.raises(ValueError):
        valid_presentation_name.create_directory_for_images()


def test_meaningful_filename_output(valid_presentation_name):
    assert (
        "None_slide_None_image_None.None"
        == valid_presentation_name.meaningful_filename(None, None, None, None)
    )


def test_if_no_presentation_is_provided_to_iter_by_shape(valid_presentation_name):
    pres = valid_presentation_name
    with pytest.raises(ValueError):
        for _ in pres.iter_by_shape(None, MSO_SHAPE_TYPE):
            pass


def test_if_no_shape_enum_is_provided_to_iter_by_shape(
    minimal_pres, valid_presentation_name
):
    pres = valid_presentation_name
    with pytest.raises(ValueError):
        for item in pres.iter_by_shape(Presentation(minimal_pres), None):
            print(item)


def test_if_image_directory_path_is_none(custom_path, minimal_pres):
    pptx_exporter = PowerPointImageExporter(minimal_pres)
    pptx_exporter.default_image_path = custom_path
    pptx_exporter.image_directory_path = None
    pptx_exporter.copy_images_to_directory()
    assert pptx_exporter.image_directory_path is not None


def test_if_image_directory_path_has_files_in_it(custom_path, minimal_pres):
    custom_path.mkdir()
    new_file = custom_path / "another_fake.jpeg"
    new_file.write_text("stuff")
    pptx_exporter = PowerPointImageExporter(minimal_pres)
    pptx_exporter.image_directory_path = custom_path
    assert len([p for p in custom_path.iterdir()]) != 0
    with pytest.raises(ValueError):
        pptx_exporter.copy_images_to_directory()
