import sys
import os
from cx_Freeze import setup, Executable

def get_customtkinter_path():
    import customtkinter
    return os.path.dirname(customtkinter.__file__)

build_exe_options = {
    "packages": ["os", "customtkinter", "logica_interface", "logica_negocio", "PIL"],
    "include_files": ["link.png", (get_customtkinter_path(), "customtkinter")],
}

base = None
if sys.platform == "win32":
    base = "gui"

setup(
    name="Distribuidor",
    version="0.3",
    description="Distribuidor de atividades",
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
