from cx_Freeze import setup, Executable

base = None

executables = [Executable("GUI_3.py", base=base)]

packages = ["idna","openpyxl","tkinter","dictionary"]
options = {
    'build_exe': {
        'packages':packages,
    },
}

setup(
    name = "<RTI WAWA ASSET TAGGER>",
    options = options,
    version = "<.11>",
    description = '<any description>',
    executables = executables
)