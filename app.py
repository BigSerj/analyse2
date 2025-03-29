# This file is deprecated. All functionality moved to appiframe.py

import os
from flask import Flask, render_template, request, send_file, jsonify, abort, Response
from markupsafe import Markup
import requests
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from datetime import datetime, timedelta
import json
import threading
import math
from time import sleep, time
from openpyxl.styles.colors import Color
from copy import copy

app = Flask(__name__)

# Add port configuration
port = int(os.environ.get("PORT", 10000))

# Загрузка конфигурации
with open('config.py', 'r') as config_file:
    exec(config_file.read())

BASE_URL = 'https://api.moysklad.ru/api/remap/1.2'

processing_cancelled = False
processing_lock = threading.Lock()
current_status = {'total': 0, 'processed': 0}
api_request_times = []

# Добавляем корневой маршрут
@app.route('/')
def index():
    return render_template('iframe.html', 
                         stores=get_stores(),
                         product_groups=get_product_groups(),
                         product_groups_json=json.dumps(prepare_groups_for_js(get_product_groups())),
                         render_group_options=render_group_options)

# Добавляем маршрут iframe (теперь он будет дублировать корневой)
@app.route('/iframe')
def iframe():
    return index()

def prepare_groups_for_js(groups):
    result = []
    for group in groups:
        group_data = {
            'id': group['id'],
            'name': group['name'],
            'children': prepare_groups_for_js(group.get('children', [])),
            'hasChildren': bool(group.get('children'))
        }
        result.append(group_data)
    return result

# Копируем все остальные функции и маршруты из appiframe.py
@app.route('/process', methods=['POST'])
def process():
    # Копируем содержимое функции process из appiframe.py
    ...

@app.route('/cancel', methods=['POST'])
def cancel_processing():
    # Копируем содержимое функции cancel_processing из appiframe.py
    ...

@app.route('/status-stream')
def status_stream():
    # Копируем содержимое функции status_stream из appiframe.py
    ...

def get_stores():
    """Получает список складов из МойСклад"""
    url = f"{BASE_URL}/entity/store"
    headers = {
        'Authorization': f'Bearer {MOYSKLAD_TOKEN}',
        'Accept': 'application/json;charset=utf-8'
    }
    
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            stores = response.json()['rows']
            return [{'id': store['id'], 'name': store['name']} for store in stores]
        else:
            error_message = f"Ошибка при получении списка складов: {response.status_code}"
            print(error_message)
            return []
    except Exception as e:
        print(f"Ошибка при получении списка складов: {str(e)}")
        return []

def get_product_groups():
    """Получает список групп товаров из МойСклад"""
    url = f"{BASE_URL}/entity/productfolder"
    headers = {
        'Authorization': f'Bearer {MOYSKLAD_TOKEN}',
        'Accept': 'application/json;charset=utf-8'
    }
    
    try:
        all_groups = []
        offset = 0
        limit = 1000

        while True:
            params = {
                'offset': offset,
                'limit': limit,
                'expand': 'productFolder'
            }
            
            response = requests.get(url, headers=headers, params=params)
            if response.status_code == 200:
                data = response.json()
                all_groups.extend(data['rows'])
                if len(data['rows']) < limit:
                    break
                offset += limit
            else:
                error_message = f"Ошибка при получении списка групп товаров: {response.status_code}"
                print(error_message)
                return []

        return build_group_hierarchy(all_groups)
    except Exception as e:
        print(f"Ошибка при получении списка групп товаров: {str(e)}")
        return []

def build_group_hierarchy(groups):
    """Строит иерархию групп товаров"""
    group_dict = {}
    root_groups = []

    # Первый проход - создаем словарь всех групп
    for group in groups:
        group_id = group['id']
        group_dict[group_id] = {
            'id': group_id,
            'name': group['name'],
            'children': [],
            'parent': None,
            'level': 0,
            'has_children': False
        }

    # Второй проход - строим иерархию
    for group in groups:
        group_id = group['id']
        parent_folder = group.get('productFolder', {})
        
        if parent_folder and parent_folder.get('meta'):
            parent_id = parent_folder['meta']['href'].split('/')[-1]
            
            if parent_id in group_dict:
                group_dict[group_id]['parent'] = parent_id
                group_dict[parent_id]['children'].append(group_dict[group_id])
                group_dict[parent_id]['has_children'] = True
        else:
            root_groups.append(group_dict[group_id])

    # Сортируем группы
    root_groups.sort(key=lambda x: x['name'])
    
    def set_levels(groups, level=0):
        for group in groups:
            group['level'] = level
            if group['children']:
                group['has_children'] = True
                set_levels(group['children'], level + 1)
                group['children'].sort(key=lambda x: x['name'])

    set_levels(root_groups)
    return root_groups

def render_group_options(groups, level=0):
    """Рендерит HTML-опции для групп товаров"""
    result = []
    for group in groups:
        indent = '—' * level
        has_children = '1' if group.get('children') and len(group['children']) > 0 else '0'
        
        option_html = (
            f'<option value="{group["id"]}" '
            f'data-level="{level}" '
            f'data-has-children="{has_children}" '
            f'data-parent="{group.get("parent", "")}" '
            f'style="margin-left: {level * 20}px">'
            f'{indent} {group["name"]}'
            f'</option>'
        )
        
        result.append(option_html)
        
        if group.get('children'):
            child_options = render_group_options(group['children'], level + 1)
            result.extend(child_options)
    
    return '\n'.join(result)

# Копируем все остальные функции из appiframe.py

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port)
