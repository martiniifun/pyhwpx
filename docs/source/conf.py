# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

import os
import sys
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
html_static_path = ['build/html/_static']
# autodoc_preserve_defaults = True
# autodoc_member_order = 'bysource'
# autoclass_member_order = 'bysource'
autosummary_generate = True
autodoc_inherit_docstrings = True
on_rtd = os.environ.get('READTHEDOCS') == 'True'
if on_rtd:
    autodoc_mock_imports = ["win32com", "numpy"]

# autodoc_mock_imports = ["pyhwpx"]
autosummary_ignore_module_all = False
autodoc_default_options = {
    "members": True,
    "show-inheritance": True,
}