#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Author: lhys
# File  : os.py

import os

def make_path_exists(path, path_type='d'):
    if path_type == 'd' and not os.path.exists(path):
        os.makedirs(path)
    if path_type == 'f' and not os.path.exists(path):
        with open(path, 'w') as f:
            f.write('')