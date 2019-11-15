# -*- coding: utf-8 -*-
# Перечень ВСЕХ пользователей

import sys, argparse
from _datetime import datetime, timedelta, date
import os
from mysql.connector import MySQLConnection, Error
from collections import OrderedDict
import openpyxl
from pymongo import MongoClient
import psycopg2

from lib import read_config, fine_phone

MYSQL_DBS = ['passport_service', 'saturn_crm', 'saturn_fin']

cfg = read_config(filename='menty.ini', section='postgresql')
conn = psycopg2.connect(**cfg)

#Наследование по Организациям
division_closure = {}
cursorA = conn.cursor()
cursorA.execute('SELECT descendant, ancestor, depth FROM division_closure -- WHERE depth > 0 ORDER BY depth DESC')
for row in cursorA:
    division_closure[row[0]] = {}
    division_closure[row[0]]['потомок'] = row[0]
    division_closure[row[0]]['предок'] = row[1]
    division_closure[row[0]]['глубина вложенности'] = row[2]

# Права по всем Организациям
divisions = {}
cursor = conn.cursor()
cursor.execute('SELECT id, title, product_access, access_model FROM division')
for row in cursor:
    divisions[row[0]] = {}
    divisions[row[0]]['title'] = row[1]
    divisions[row[0]]['product_access'] = row[2]
    divisions[row[0]]['access_model'] = row[3]
    if divisions[row[0]]['access_model'] == 300:
        divisions[row[0]]['product_access'] = ['Полный доступ']
    elif divisions[row[0]]['access_model'] == 100:
        divisions[row[0]]['product_access'] = {'0': 'Ошибка при расчете наследования'}
for division in divisions:
    if divisions[division]['access_model'] == 100:
        product_access_tek = divisions[division]['product_access']
        depth = division_closure[division]['глубина вложенности']
        division_tek = division
        for i in range(depth, 0, -1):
            if product_access_tek['0'] == 'Ошибка при расчете наследования':
                division_tek = division_closure[division_tek]['предок']
                product_access_tek = divisions[division_tek]['product_access']
            else:
                divisions[division]['product_access'] = product_access_tek
                break

q=0



wb_rez = openpyxl.Workbook(write_only=True)
ws_rez = wb_rez.create_sheet('Список пользователей')
ws_rez.append(['id', 'Фамилия', 'Имя', 'Отчество', 'Телефон', 'E-mail', 'Логин пользователя', 'СНИЛС', 'Подразделение',
               'Доступ к продуктам','паспорт-титульник', 'паспорт-регистрация','СНИЛС'])

conn = psycopg2.connect(**cfg)
tables_cursor = conn.cursor()
sql = "SELECT table_name FROM information_schema.tables WHERE table_schema NOT IN " \
      "('information_schema','pg_catalog') AND table_schema IN('public', 'myschema')"
tables_cursor.execute(sql)
cursor = conn.cursor()
ws_rez.append(wb_rez.create_sheet('ЧтобыРаботало'))
ws_rez.append(wb_rez.create_sheet('Postgresql'))
for table_row in tables_cursor:
    cursor.execute('SELECT * FROM ' + table_row[0] + ' LIMIT 2')
    col_names = [desc[0] for desc in cursor.description]
    cursor.execute('SELECT * FROM ' + table_row[0] + ' ORDER BY ' + col_names[0] + ' DESC LIMIT 5')
    ws_rez[len(ws_rez) - 1].append([])
    ws_rez[len(ws_rez) - 1].append([table_row[0]])
    ws_rez[len(ws_rez) - 1].append(col_names)
    cycled = False
    for row in cursor:
        col_rez = []
        cycled = True
        for j, col_name in enumerate(col_names):
            if row[j]:
                if str(type(row[j])).find('str') < 0:
                    col_rez.append(str(row[j]))
                else:
                    col_rez.append(row[j])
            else:
                col_rez.append('--пусто--')
        ws_rez[len(ws_rez) - 1].append(col_rez)
    if not cycled:
        col_rez = []
        for j, col_name in enumerate(col_names):
            col_rez.append('--пусто--')
        ws_rez[len(ws_rez) - 1].append(col_rez)

"""
wb_rez.save(datetime.now().strftime('%Y-%m-%d_%H-%M') + 'databases.xlsx')
"""