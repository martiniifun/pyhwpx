# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

import os
import sys
from unittest.mock import MagicMock
from pathlib import Path


# sys.path.insert(0, str(Path(__file__).resolve().parents[0]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
sys.path.insert(0, os.path.abspath("../.."))



project = 'pyhwpx'
copyright = '2025, ilco'
author = 'ilco'
release = '0.44.9'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = [
    'sphinx.ext.duration',
    'sphinx.ext.autodoc',
    'sphinx.ext.autosummary',
    'sphinx.ext.todo',
]
templates_path = ['_templates']
exclude_patterns = []

language = 'ko'

# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = "sphinx_rtd_theme" # 'furo'  # 'alabaster'
html_static_path = ['_static']  # docs/source/_static, docs/build/html/_static
html_theme_options = {
    'collapse_navigation': False,
    'sticky_navigation': True,
    'navigation_depth': 4,
    'style_external_links': True
}
# autodoc_preserve_defaults = True
autodoc_member_order = 'bysource'
autoclass_member_order = 'bysource'
autosummary_generate = True
autodoc_inherit_docstrings = True
on_rtd = os.environ.get('READTHEDOCS') == 'True'
# if on_rtd:
MOCK_MODULES = ["win32com", "win32com.client", "pythoncom", "pywintypes"]
sys.modules.update((mod_name, MagicMock()) for mod_name in MOCK_MODULES)
autodoc_mock_imports = ["pywin32", "numpy", "pandas", "pyperclip", "pillow"]

autosummary_ignore_module_all = False
autodoc_default_options = {
    "members": True,
    "show-inheritance": True,
}
