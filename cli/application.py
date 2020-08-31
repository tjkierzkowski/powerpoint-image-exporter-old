import argparse
import logging

from pptx_export.pptx_export import DEFAULT_DIR, PowerPointImageExporter


class CommandLineApplication:

    def __init__(self):
        self.cli_app = argparse.ArgumentParser(description="Export all images from a powerpoint lecture (.pptx) into a "
                                                           "directory")
        self.cli_app.add_argument("pptx_file_path",
                                  metavar="<pptx file>",
                                  help="full or relative path to the pptx file")
        self.cli_app.add_argument("-o",
                                  "--output-dir",
                                  dest="output_directory",
                                  metavar="<output directory>",
                                  help="full or relative path of either an empty or to be created output directory for "
                                       f"images. Defaults to creating a new local directory '{DEFAULT_DIR}' ",
                                  default=DEFAULT_DIR)

    def execute(self):
        arguments = self.cli_app.parse_args()
        pptx_exporter = None
        try:
            pptx_exporter = PowerPointImageExporter(arguments.pptx_file_path)
        except Exception as ex:
            logging.exception("An exception occurred when verifying your pptx file", exc_info=True)
        try:
            if arguments.output_directory is not None:
                pptx_exporter.create_directory_for_images(arguments.output_directory)
        except Exception as ex:
            logging.exception(f"An exception occurred when creating the custom directory", exc_info=True)
        try:
            pptx_exporter.copy_images_to_directory()
        except Exception as ex:
            logging.exception("An exception occurred when exporting the images", exc_info=True)

