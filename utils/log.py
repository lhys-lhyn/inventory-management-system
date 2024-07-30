#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Author: lhys
# File  : log.py

import time
import logging
import tkinter as tk
from threading import Lock
from collections.abc import Iterable

class Record:

    def __init__(self, log_text, log_file='process.log'):
        logging.basicConfig(
            filename=log_file,
            filemode='w',
            format='%(asctime)s - %(pathname)s[line: %(lineno)d] - %(levelname)s - %(message)s',
            level=logging.INFO
        )
        self.lock = Lock()
        self.log_text = log_text
        self.encoding_error_flag = False

    @staticmethod
    def format_string(*items, extract=True):
        return ' '.join([
            Record.format_string(*item) if isinstance(item, Iterable) and type(item) != str and extract else str(item)
            for item in items
            ]).replace('\n', '')

    @staticmethod
    def get_format_current_time():
        return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))

    def lock_logging(self, msg, level='info'):
        level = level.lower()
        if level in ['error', 'warning', 'info', 'debug']:
            try:
                if not self.encoding_error_flag:
                    with open('process.log', encoding='utf-8') as f:
                        f.read()
                    # with open('process1.log', 'w') as f:
                    #     f.write(f'{Record.get_format_current_time()} - {level} - ' + msg + '\n')
                func = getattr(logging, level)
                func(msg)
            except UnicodeEncodeError as e:
                self.encoding_error_flag = True
                print(e)
            return
        else:
            raise ValueError(f'Input level "{level}" is invalid, not in ["error", "warning", "info", "debug"].')

    def lock_print(self, msg, out=None):
        if out is not None:
            out.write(msg)
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def lock_output(self, *msg, level='info', out='log', msg_extract=True):
        with self.lock:
            level = level.lower()
            msg = Record.format_string(*msg, extract=msg_extract)
            if 'log' in out or 'all' in out:
                self.lock_logging(msg, level=level)
            if 'text' in out or 'all' in out:
                self.lock_print(msg)
            if level == 'error':
                tk.messagebox.showerror('错误', msg)

    def LogWrapper(self, retry=True, pop_up=False):
        def wrapper(func):
            def inner(*args, **kwargs):
                flag = True
                while flag or retry:
                    try:
                        return func(*args, **kwargs)
                    except Exception as e:
                        flag = False
                        message = str(e).replace("\n", " ")
                        self.lock_output(f'type: {e.__class__.__name__}, '
                                         f'func: {func.__name__}, args:{args}, kwargs: {kwargs}, message: {message}',
                                         level='error', out='all')
                        if pop_up:
                            tk.messagebox.showerror('错误', '程序出了一些错误，请查看日志文件。')

            return inner

        return wrapper


if __name__ == '__main__':
    recorder = Record()
