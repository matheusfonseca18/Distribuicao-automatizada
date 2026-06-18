import sys
import os
from cx_Freeze import setup, Executable

def get_customtkinter_path():
    import customtkinter
    return os.path.dirname(customtkinter.__file__)

build_exe_options = {
    "packages": [
        "os",
        "json",
        "random",
        "datetime",
        "itertools",
        "tkinter",
        "customtkinter",
        "CTkMessagebox",
        "PIL",
        "pandas",
        "openpyxl",
    ],
    "includes": [
        "logica_interface",
        "logica_negocio",
        "apelidos_utils",
    ],
    "include_files": [
        "link.png",
        (get_customtkinter_path(), "customtkinter"),
    ],
    "include_msvcr": True,
}

base = None
if sys.platform == "win32":
    base = "gui"

setup(
    name="Distribuidor",
    version="1.0",
    description="Distribuidor de atividades - 1.0",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "interface.py",
            base=base,
            target_name="Distribuidor.exe",
            icon=None
        )
    ],
)
