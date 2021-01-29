import sys
from cx_Freeze import setup, Executable


base = None
if sys.platform == "win32":
    base = "Win32GUI"
    icon="Logo.ico"
executables = [
        Executable("RoboTela.py",icon=icon ,base=base)
]

buildOptions = dict(
        packages = [],
        includes = [],
        include_files = [],
        excludes = []
)

setup(
    name = "WhatsRobot",
    version = "1.0",
    description = "Robô para envio de mensagens de cobranças JCNet.",
    options = dict(build_exe = buildOptions),
    executables = executables
 )
