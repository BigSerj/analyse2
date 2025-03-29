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

# Копируем все остальные функции из appiframe.py
def get_stores():
    # Копируем содержимое функции get_stores
    ...

def get_product_groups():
    # Копируем содержимое функции get_product_groups
    ...

# ... и так далее для всех остальных функций

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port)
