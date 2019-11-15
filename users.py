# -*- coding: utf-8 -*-
# Перечень ВСЕХ пользователей

import openpyxl
from pymongo import MongoClient
import psycopg2

from lib import read_config, fine_phone

MYSQL_DBS = ['passport_service', 'saturn_crm', 'saturn_fin']

cfg = read_config(filename='menty.ini', section='postgresql')
conn = psycopg2.connect(**cfg)

# Наследование по Организациям
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

# Таблицы access_group
cursorB = conn.cursor()
cursorB.execute('SELECT ua.user_id, ag.title, ag.content FROM user_access_group AS ua '
                'LEFT JOIN access_group AS ag ON ua.access_group_id = ag.id')
user_access_groups = {}
for row in cursorB:
    if user_access_groups.get(row[0]):
        user_access_groups[row[0]] += ', ' + str(row[1])
    else:
        user_access_groups[row[0]] = str(row[1])

# Все пользователи
cursorC = conn.cursor()
cursorC.execute('SELECT ac.id, ac.roles, ac.division_id, ac.passport_main_id, ac.passport_registration_id, '
                'ac.insurance_document_id, ac.lastname, ac.name, ac.middlename, ac.phone, ac.email, ac.username, '
                'ac.insurance, at.title '
                'FROM account AS ac LEFT JOIN access_template AS at ON ac.access_template_id = at.id')
acc = {}
for row in cursorC:
    acc[row[0]] = {}
    if row[2]:
        acc[row[0]]['Подразделение'] =  divisions[row[2]]['title']
        acc[row[0]]['Доступ к продуктам'] = ' '.join(divisions[row[2]]['product_access'])
    else:
        acc[row[0]]['Подразделение'] =  'Не выбрано'
        acc[row[0]]['Доступ к продуктам'] = 'Нет'
    if row[3]:
        acc[row[0]]['паспорт-титульник'] = 'Есть'
    else:
        acc[row[0]]['паспорт-титульник'] = 'Нет'
    if row[4]:
        acc[row[0]]['паспорт-регистрация'] = 'Есть'
    else:
        acc[row[0]]['паспорт-регистрация'] = 'Нет'
    if row[5]:
        acc[row[0]]['скан СНИЛС'] = 'Есть'
    else:
        acc[row[0]]['скан СНИЛС'] = 'Нет'
    acc[row[0]]['Фамилия'] = row[6]
    acc[row[0]]['Имя'] = row[7]
    acc[row[0]]['Отчество'] = row[8]
    acc[row[0]]['Телефон'] = row[9]
    acc[row[0]]['E-mail'] = row[10]
    acc[row[0]]['Логин пользователя'] = row[11]
    acc[row[0]]['СНИЛС'] = row[12]
    acc[row[0]]['Должность'] = row[13]
    if user_access_groups.get(row[0]):
        acc[row[0]]['Роль'] = user_access_groups[row[0]]
    else:
        acc[row[0]]['Роль'] = ''

wb_rez = openpyxl.Workbook(write_only=True)
ws_rez = wb_rez.create_sheet('Список пользователей')
ws_rez.append(['id', 'Фамилия', 'Имя', 'Отчество', 'Телефон', 'E-mail', 'Логин пользователя', 'СНИЛС', 'Должность',
               'Роль', 'Подразделение', 'паспорт-титульник', 'паспорт-регистрация','скан СНИЛС', 'Доступ к продуктам'])
for a in acc:
    ws_rez.append([a, acc[a]['Фамилия'], acc[a]['Имя'], acc[a]['Отчество'], acc[a]['Телефон'],
                   acc[a]['E-mail'], acc[a]['Логин пользователя'], acc[a]['СНИЛС'], acc[a]['Должность'],
                   acc[a]['Роль'], acc[a]['Подразделение'], acc[a]['паспорт-титульник'], acc[a]['паспорт-регистрация'],
                   acc[a]['скан СНИЛС'], acc[a]['Доступ к продуктам']])
wb_rez.save('users.xlsx')
