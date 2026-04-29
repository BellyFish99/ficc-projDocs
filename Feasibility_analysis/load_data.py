"""
load_data.py
============
Shared loader for the YAML SSoT pipeline.

Reads ficc_data.yaml (meta + roles) and merges all module files from
the modules/ subfolder into a single data dict identical in shape to
the old single-file format:
  {
    'meta':    {...},
    'roles':   [...],
    'modules': [...]   # sorted by module 'no'
  }
"""

import glob
import os
import yaml

_HERE = os.path.dirname(os.path.abspath(__file__))
MASTER_PATH = os.path.join(_HERE, 'ficc_data.yaml')
MODULES_DIR = os.path.join(_HERE, 'modules')


def load():
    with open(MASTER_PATH, encoding='utf-8') as f:
        data = yaml.safe_load(f)

    module_files = sorted(glob.glob(os.path.join(MODULES_DIR, '*.yaml')))

    modules = []
    for path in module_files:
        with open(path, encoding='utf-8') as f:
            doc = yaml.safe_load(f)
        modules.append(doc['module'])

    modules.sort(key=lambda m: m['no'])
    data['modules'] = modules
    return data
