# -*- coding=utf-8 -*-
# @Time    : 2021/6/18 20:37
# @Author  : lhys
# @FileName: sql_tools.py

import sqlite3

class SqliteOperation:

    def __init__(self, db_name, recorder):
        self.driver = sqlite3.connect(db_name)
        self.recorder = recorder

    def return_driver(self):
        return self.driver

    def exec_sql(self, sql, *args, mode='single'):
        try:
            cursor = self.driver.cursor()
            # 执行 sql 语句
            if mode == 'single':
                result = cursor.execute(sql, *args).fetchall()
            else:
                result = cursor.executemany(sql, *args).fetchall()
            # 提交 sql 语句
            self.driver.commit()
            self.recorder.lock_output(f"Execute '{sql}' successfully, args: [{', '.join(map(str, args))}], result: {result}")
            return result, True
        except Exception as e:
            sql = sql.replace('\n', '')
            self.recorder.lock_output(f"Execute '{sql}' error, args: [{', '.join(map(str, args))}], {e}", level='error', out='all')
            return [], False

    def _concat_fields(self, fields, extra=''):
        if not fields:
            return ''
        if isinstance(fields, str):
            # 如果字段名包含括号，可能是函数调用，如 "MIN(day)"
            if '(' in fields and ')' in fields:
                # 分割函数名和字段名
                func_name, field_name = fields.split('(')
                field_name = field_name.strip(') ')  # 去掉右括号和多余的空格
                # 添加反引号并重组
                return f"{func_name}(`{field_name}`)"
            elif '*' not in fields:
                # 普通字段名，直接添加反引号
                return f"`{fields}`"
            else:
                return fields
        else:
            return '(`' + '`, `'.join([kw + extra for kw in fields]) + '`)'

    def _concat_value_string(self, num, key=''):
        return ', '.join([(key + str(i) + '=' if key else '') + '?' for i in range(num)])

    def get_column_names(self, table_name):
        result, _ = self.exec_sql(f"PRAGMA table_info({table_name})")
        columns = [column[1] for column in result]
        return columns

    def insert(self, table, data, keywords=[], mode='single'):
        kw_sql = self._concat_fields(keywords)
        value_sql = self._concat_value_string(len(data) if mode == 'single' else len(data[0]))
        sql = f'''INSERT INTO {table} {kw_sql} VALUES ({value_sql})'''
        return self.exec_sql(sql, data, mode=mode)

    def delete(self, table, limit=None):
        sql = f'delete from {table} '
        # 如果有限制条件就加上
        if limit: sql += limit
        return self.exec_sql(sql)

    def modify(self, table, keywords, data, *args, limit=''):
        limit = f'where {limit}' if limit else ''
        # 同上，拼接 sql 语句
        sql = f'update {table} set {self._concat_fields(keywords, " = ?")} {limit}'
        return self.exec_sql(sql, data)

    def search(self, table, *args, field='*', limit=''):
        kw_sql = self._concat_fields(field)
        # limit = f'where {limit}' if limit else ''
        # 拼接 sql 字符串
        sql = f'select {kw_sql} from {table} {limit}'
        return self.exec_sql(sql, *args)

    def union_search(self, table1, table2, *args, field1='*', field2='*',  limit=''):
        kw1, kw2 = map(self._concat_fields, [field1, field2])
        sql = f'''
            SELECT {kw1} FROM {table1} {limit} 
            UNION 
            SELECT {kw2} FROM {table2} {limit}
        '''
        return self.exec_sql(sql, *args)

    def close(self):
        self.driver.close()

    def __enter__(self):
        # 返回实例化对象
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        if exc_type is not None:
            self.recorder.lock_output(f'exc_type:{exc_type}, exc_val:{exc_val}.', level='error')
            return False
        return True