import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.abspath('..'))

project = 'Course Management Toolkit'
author = 'Duc A. Hoang'
version = '0.1.1'
release = version

extensions = [
    'sphinx.ext.autodoc',
    'sphinx.ext.napoleon',
    'sphinx.ext.viewcode',
    'sphinx.ext.autosummary',
]

autosummary_generate = True

templates_path = ['_templates']
exclude_patterns = ['_build']

autodoc_typehints = 'description'

html_theme = 'alabaster'
html_static_path = ['_static']
today_fmt = '%Y-%m-%d'
html_last_updated_fmt = '%Y-%m-%d'

copyright = f"{datetime.now().year}, {author}"
