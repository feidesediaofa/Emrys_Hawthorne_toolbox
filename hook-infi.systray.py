import pkg_resources.py2_warn
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules('pkg_resources')