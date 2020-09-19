import click

from pptx_export.pptx_export import DEFAULT_DIR, PowerPointImageExporter
from . import __version__


@click.command()
@click.argument(
    "pptx_file_path",
    metavar="<pptx file>",
    required=1,
    type=click.Path(exists=True, file_okay=True, dir_okay=False),
)
@click.option(
    "-o",
    "--output-dir",
    "output_directory",
    metavar="<output directory>",
    help="full or relative path of either an empty or to be created "
    "output directory for images.",
    default=DEFAULT_DIR,
    show_default="f{DEFAULT_DIR}",
)
@click.version_option(version=__version__)
def main(pptx_file_path, output_directory):
    """Export all images from a powerpoint lecture (.pptx) into a directory

    pptx_file_path: full or relative path to the pptx file
    """
    exporter = PowerPointImageExporter(pptx_file_path)
    exporter.create_directory_for_images(output_directory)
    exporter.copy_images_to_directory()
