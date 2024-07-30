#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Author: lhys
# File  : global.py

import os
import time
import random
import logging


@LogWrapper(retry=False)
def show_tips(title, *msg, level='error', out='log'):
    level = level.lower()

    if level in ['info', 'error', 'warning']:
        func = eval(f'messagebox.show{level}')
        func(format_string(title), format_string(*msg))
        lock_output(msg, level=level, out=out)
    else:
        raise ValueError(f'Input level "{level}" is invalid, not in ["info", "error", "warning"].')

cache_path = 'cache'

def set_cache_path(work_path):
    global cache_path
    cache_path = os.path.join(os.path.dirname(work_path), 'cache')

def init_cache_path():
    if not os.path.exists(cache_path):
        os.makedirs(cache_path)