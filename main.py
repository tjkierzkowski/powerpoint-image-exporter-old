from pptx_export.pptx_export import PowerPointImageExporter
from pathlib import Path
# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.i




# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    desktop_path = str(Path.home()) + '/Desktop/'
    test_file = Path(desktop_path + 'Fall 2020 Block 1 Lab Review w labels 1-3 and osteo-1.pptx')
    # project_file = '/home/tomk/PycharmProjects/powerpointImageExporter/tests/resources/stub_tester.pptx'
    project_file = '//tests/resources/shorter_test.pptx'
    pptx_img_exporter = PowerPointImageExporter(project_file)

    new_directory = desktop_path + 'test_output'
    print(f'\n{new_directory}\n')

    pptx_img_exporter.copy_images_to_directory()





# See PyCharm help at https://www.jetbrains.com/help/pycharm/
