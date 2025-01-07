import os
from flask import Flask, render_template, request, send_file, jsonify, abort, Response
from markupsafe import Markup
import requests
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from datetime import datetime, timedelta
import json
import threading
import math
from time import sleep

app = Flask(__name__)

# Загрузка конфигурации так же, как в основном приложении
with open('config.py', 'r') as config_file:
    exec(config_file.read())

BASE_URL = 'https://api.moysklad.ru/api/remap/1.2'

processing_cancelled = False
processing_lock = threading.Lock()

# Добавим глобальную переменную для хранения статуса
current_status = {'total': 0, 'processed': 0}

def check_if_cancelled():
    """Проверяет, не была ли отменена обработка отчета"""
    global processing_cancelled
    if processing_cancelled:
        return True
    return False

# Основной маршрут для iframe
@app.route('/iframe', methods=['GET'])
def iframe():
    stores = get_stores()
    product_groups = get_product_groups()
    
    # Преобразуем структуру групп для корректной работы в JavaScript
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
    
    prepared_groups = prepare_groups_for_js(product_groups)
    product_groups_json = json.dumps(prepared_groups)
    
    return render_template(
        'iframe.html',
        stores=stores,
        product_groups=product_groups,
        product_groups_json=product_groups_json,
        render_group_options=render_group_options
    )

# Маршрут для обработки формы через AJAX
@app.route('/process', methods=['POST'])
def process():
    try:
        global processing_cancelled, current_status
        with processing_lock:
            processing_cancelled = False
            current_status = {'total': 0, 'processed': 0}
        
        start_date = request.form['start_date']
        end_date = request.form['end_date']
        store_id = request.form['store_id']
        planning_days = int(request.form['planning_days'])
        
        product_groups = []
        if 'final_product_groups' in request.form and request.form['final_product_groups']:
            raw_groups = request.form['final_product_groups']
            product_groups = [group for group in raw_groups.split(',') if group]
        
        manual_stock_settings = request.form.get('final_manual_stock_groups', '[]')
        
        # Проверяем отмену перед получением данных
        if check_if_cancelled():
            return jsonify({'cancelled': True}), 200
            
        report_data = get_report_data(start_date, end_date, store_id, product_groups)
        
        if not report_data or 'rows' not in report_data or not report_data['rows']:
            return jsonify({'error': 'Нет данных для формирования отчета'}), 404
        
        # Проверяем отмену после получения данных
        if check_if_cancelled():
            return jsonify({'cancelled': True}), 200
            
        # Считаем только позиции с продажами и вариантами
        total_items = sum(1 for item in report_data['rows'] 
                         if item.get('sellQuantity', 0) > 0 
                         and ('/variant/' in item.get('assortment', {}).get('meta', {}).get('href', '') 
                             or '/product/' in item.get('assortment', {}).get('meta', {}).get('href', '')))
        
        with processing_lock:
            current_status['total'] = total_items
            current_status['processed'] = 0
        
        # Проверяем отмену перед созданием отчета
        if check_if_cancelled():
            return jsonify({'cancelled': True}), 200
            
        excel_file = create_excel_report(report_data, store_id, start_date, end_date, planning_days, manual_stock_settings)
        
        # Финальная проверка отмены перед отправкой результата
        if check_if_cancelled():
            return jsonify({'cancelled': True}), 200
            
        return jsonify({
            'success': True, 
            'file_url': f'/download/{excel_file}'
        })
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Маршрут для скачивания файла
@app.route('/download/<filename>')
def download_file(filename):
    try:
        return send_file(filename, as_attachment=True, download_name='profitability_report.xlsx')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Остальные функции остаются без изменений, так как они отвечают за бизнес-логику
# Здесь идут все остальные функции из app.py без изменений:
# get_stores(), get_product_groups(), get_report_data() и т.д.

# Создадим новый маршрут для встраивания iframe
@app.route('/embed')
def embed():
    return """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Встроенный отчет</title>
        <style>
            body, html {
                margin: 0;
                padding: 0;
                height: 100%;
                width: 100%;
            }
            iframe {
                width: 100%;
                height: 100%;
                border: none;
            }
            .overlay {
                display: none;
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(255, 255, 255, 0.4);
                z-index: 999;
            }
            .status-box {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: white;
                padding: 20px 30px;
                border-radius: 8px;
                box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
                z-index: 1000;
                text-align: center;
                min-width: 300px;
            }
            .status-box h3 {
                margin: 0 0 15px 0;
                font-size: 18px;
                color: #333;
            }
            .status-box.hidden {
                display: none;
            }
            .stop-button {
                margin-top: 20px;
                padding: 8px 20px;
                background-color: #dc3545;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
                transition: background-color 0.2s;
            }
            .stop-button:hover {
                background-color: #c82333;
            }
            .confirm-modal {
                display: none;
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: white;
                padding: 25px;
                border-radius: 8px;
                box-shadow: 0 2px 15px rgba(0, 0, 0, 0.2);
                z-index: 1001;
                text-align: center;
                min-width: 300px;
            }
            .confirm-modal p {
                margin: 0 0 20px 0;
                font-size: 16px;
                color: #333;
            }
            .confirm-modal-buttons {
                display: flex;
                justify-content: center;
                gap: 15px;
            }
            .confirm-yes {
                padding: 8px 20px;
                background-color: #dc3545;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
                transition: background-color 0.2s;
            }
            .confirm-no {
                padding: 8px 20px;
                background-color: #6c757d;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
                transition: background-color 0.2s;
            }
            .confirm-yes:hover {
                background-color: #c82333;
            }
            .confirm-no:hover {
                background-color: #5a6268;
            }
        </style>
    </head>
    <body>
        <div id="overlay" class="overlay"></div>
        <div id="statusBox" class="status-box hidden">
            <h3>Формирование отчета</h3>
            <div>Осталось обработать позиций: <span id="remainingItems">...</span></div>
            <div style="margin-top: 10px;">Осталось примерно времени: <span id="remainingTime">...</span></div>
            <button class="stop-button" onclick="showConfirmModal()">Остановить</button>
        </div>
        <div id="confirmModal" class="confirm-modal">
            <p>Точно остановить формирование отчета?</p>
            <div class="confirm-modal-buttons">
                <button class="confirm-yes" onclick="confirmStop()">Да, остановить</button>
                <button class="confirm-no" onclick="hideConfirmModal()">Нет, продолжить формирование отчета</button>
            </div>
        </div>
        <iframe src="/iframe" allowfullscreen></iframe>
        <script>
            const overlay = document.getElementById('overlay');
            const statusBox = document.getElementById('statusBox');
            const confirmModal = document.getElementById('confirmModal');
            const remainingItems = document.getElementById('remainingItems');
            const remainingTime = document.getElementById('remainingTime');
            let currentEventSource = null;
            
            function showConfirmModal() {
                confirmModal.style.display = 'block';
            }
            
            function hideConfirmModal() {
                confirmModal.style.display = 'none';
            }
            
            function confirmStop() {
                hideConfirmModal();
                stopProcessing();
            }
            
            function stopProcessing() {
                fetch('/cancel', { method: 'POST' })
                    .then(() => {
                        if (currentEventSource) {
                            currentEventSource.close();
                        }
                        overlay.style.display = 'none';
                        statusBox.classList.add('hidden');
                    });
            }
            
            // Слушаем сообщения от основного окна
            window.addEventListener('message', function(event) {
                if (event.data === 'startProcessing') {
                    // Показываем оверлей и статус-бокс
                    overlay.style.display = 'block';
                    statusBox.classList.remove('hidden');
                    startEventSource();
                }
            });

            function startEventSource() {
                if (currentEventSource) {
                    currentEventSource.close();
                }
                currentEventSource = new EventSource('/status-stream');
                currentEventSource.onmessage = function(event) {
                    const remaining = event.data;
                    if (remaining === '...') {
                        remainingItems.textContent = '...';
                        remainingTime.textContent = '...';
                    } else {
                        const remainingNum = parseInt(remaining);
                        if (remainingNum > 0) {
                            remainingItems.textContent = remainingNum;
                            const totalSeconds = Math.ceil(remainingNum * 3);
                            
                            if (totalSeconds <= 0) {
                                remainingTime.textContent = '0 сек';
                            } else {
                                const hours = Math.floor(totalSeconds / 3600);
                                const minutes = Math.floor((totalSeconds % 3600) / 60);
                                const seconds = totalSeconds % 60;
                                
                                let timeString = '';
                                if (hours > 0) timeString += hours + ' ч ';
                                if (minutes > 0) timeString += minutes + ' мин ';
                                if (seconds > 0 || timeString === '') timeString += seconds + ' сек';
                                
                                remainingTime.textContent = timeString.trim();
                            }
                        } else {
                            overlay.style.display = 'none';
                            statusBox.classList.add('hidden');
                            currentEventSource.close();
                        }
                    }
                };
            }
        </script>
    </body>
    </html>
    """

# Добавляем функцию render_group_options из app.py
def render_group_options(groups, level=0):
    result = []
    for group in groups:
        indent = '—' * level
        has_children = '1' if group.get('children') and len(group['children']) > 0 else '0'
        
        # print(f"\nRendering group: {group['name']}")
        # print(f"  Level: {level}")
        # print(f"  Has children: {has_children}")
        # print(f"  Children count: {len(group.get('children', []))}")
        # print(f"  Parent: {group.get('parent')}")
        # print(f"  Raw group data: {group}")  # Добавляем вывод сырых данных группы
        
        option_html = (
            f'<option value="{group["id"]}" '
            f'data-level="{level}" '
            f'data-has-children="{has_children}" '
            f'data-parent="{group.get("parent", "")}" '
            f'style="margin-left: {level * 20}px">'
            f'{indent} {group["name"]}'
            f'</option>'
        )
        
        # print(f"  Generated HTML: {option_html}")
        result.append(option_html)
        
        if group.get('children'):
            child_options = render_group_options(group['children'], level + 1)
            # print(f"  Added {len(child_options.split('\n'))} child options for {group['name']}")
            result.extend(child_options)
    
    return '\n'.join(result)

# Добавляем функцию get_stores из app.py
def get_stores():
    url = f"{BASE_URL}/entity/store"
    headers = {
        'Authorization': f'Bearer {MOYSKLAD_TOKEN}',
        'Accept': 'application/json;charset=utf-8'
    }
    
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        stores = response.json()['rows']
        return [{'id': store['id'], 'name': store['name']} for store in stores]
    else:
        error_message = f"Ошибка при получении списка складов: {response.status_code}"
        print(error_message)
        raise Exception(error_message)

# Добавляем функцию get_product_groups из app.py
def get_product_groups():
    url = f"{BASE_URL}/entity/productfolder"
    headers = {
        'Authorization': f'Bearer {MOYSKLAD_TOKEN}',
        'Accept': 'application/json;charset=utf-8'
    }
    
    all_groups = []
    offset = 0
    limit = 1000

    # Получаем все группы одним запросом
    while True:
        params = {
            'offset': offset,
            'limit': limit,
            'expand': 'productFolder'  # Добавляем expand для получения полной информации о родительской группе
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
            raise Exception(error_message)

    # Строим иерархию
    return build_group_hierarchy(all_groups)

# Добавляем функцию build_group_hierarchy из app.py
def build_group_hierarchy(groups):
    # print("\nStarting build_group_hierarchy")
    # print(f"Received {len(groups)} groups")
    
    group_dict = {}
    root_groups = []

    # Первый проход - создаем словарь всех групп
    # print("\nFirst pass - creating groups dictionary")
    for group in groups:
        group_id = group['id']
        # print(f"\nProcessing group: {group['name']} (ID: {group_id})")
        # print(f"Raw group data: {group}")  # Печатаем сырые данные группы
        
        group_dict[group_id] = {
            'id': group_id,
            'name': group['name'],
            'children': [],
            'parent': None,
            'level': 0,
            'has_children': False
        }

    # Второй проход - строим иерархию
    # print("\nSecond pass - building hierarchy")
    for group in groups:
        group_id = group['id']
        parent_folder = group.get('productFolder', {})
        # print(f"\nProcessing group: {group['name']} (ID: {group_id})")
        # print(f"Parent folder data: {parent_folder}")
        
        if parent_folder and parent_folder.get('meta'):
            parent_id = parent_folder['meta']['href'].split('/')[-1]
            # print(f"Found parent ID: {parent_id}")
            
            if parent_id in group_dict:
                group_dict[group_id]['parent'] = parent_id
                group_dict[parent_id]['children'].append(group_dict[group_id])
                group_dict[parent_id]['has_children'] = True
                # print(f"Added group {group_id} as child to parent {parent_id}")
                # print(f"Parent's children count: {len(group_dict[parent_id]['children'])}")
        else:
            root_groups.append(group_dict[group_id])
            # print(f"Added group {group_id} to root groups")

    # print(f"\nFound {len(root_groups)} root groups")
    
    # Рекурсивная функция для установки уровней
    def set_levels(groups, level=0):
        for group in groups:
            group['level'] = level
            # print(f"Set level {level} for group {group['name']}")
            if group['children']:
                group['has_children'] = True
                # print(f"Group {group['name']} has {len(group['children'])} children")
                set_levels(group['children'], level + 1)
                group['children'].sort(key=lambda x: x['name'])

    # Сортируем и устанавливаем уровни
    root_groups.sort(key=lambda x: x['name'])
    # print("\nSetting levels for hierarchy")
    set_levels(root_groups)

    # print("\nFinal hierarchy structure:")
    # print_hierarchy(root_groups)
    
    return root_groups

def print_hierarchy(groups, level=0):
    """Вспомогательная функция для отладки иерархии"""
    for group in groups:
        print("  " * level + f"- {group['name']}")
        print("  " * level + f"  id: {group['id']}")
        print("  " * level + f"  level: {group['level']}")
        print("  " * level + f"  has_children: {group['has_children']}")
        print("  " * level + f"  children count: {len(group['children'])}")
        if group['children']:
            print_hierarchy(group['children'], level + 1)

# Добавляем корневой маршрут
@app.route('/')
def index():
    return embed()

# Добавьте все остальные функции из app.py:
def get_report_data(start_date, end_date, store_id, product_groups):
    print(f"\nStarting get_report_data with product_groups: {product_groups}")  # Начало функции
    
    global processing_cancelled
    with processing_lock:
        processing_cancelled = False
    
    try:
        url = f"{BASE_URL}/report/profit/byvariant"
        headers = {
            'Authorization': f'Bearer {MOYSKLAD_TOKEN}',
            'Accept': 'application/json;charset=utf-8',
            'Content-Type': 'application/json'
        }
        
        start_datetime = datetime.strptime(start_date, '%Y-%m-%d')
        end_datetime = datetime.strptime(end_date, '%Y-%m-%d')
        
        formatted_start = start_datetime.strftime('%Y-%m-%d %H:%M:%S')
        formatted_end = end_datetime.replace(hour=23, minute=59, second=59).strftime('%Y-%m-%d %H:%M:%S')
        
        params = {
            'momentFrom': formatted_start,
            'momentTo': formatted_end,
            'limit': 1000,
            'offset': 0
        }
        
        filter_parts = []
        
        print(f"Building filter with product_groups: {product_groups}")  # Отладка
        
        # Добавляем фильтр по складу
        if store_id:
            store_url = f"{BASE_URL}/entity/store/{store_id}"
            store_filter = f'store={store_url}'
            filter_parts.append(store_filter)
            print(f"Added store filter: {store_filter}")  # Отладка
        
        # Формируем фильтр по группам
        if product_groups:
            # Формируем фильтр для всех выбранных групп
            group_filters = []
            for group_id in product_groups:
                if group_id:
                    product_folder_url = f"{BASE_URL}/entity/productfolder/{group_id}"
                    group_filters.append(f"productFolder={product_folder_url}")
            
            if group_filters:
                # Объединяем фильтры через ИЛИ (;)
                folders_filter = '&filter='.join(group_filters)
                filter_parts.append(folders_filter)
        
        if filter_parts:
            # Объединяем все части фильтра через запятую (И)
            params['filter'] = '&filter='.join(filter_parts)
            print(f"Final filter parameter: {params['filter']}")  # Отладка
        
        all_rows = []
        total_count = None
        
        while True:
            check_if_cancelled()
            
            query_params = [f"{k}={v}" for k, v in params.items() if k != 'filter']
            if 'filter' in params:
                query_params.append(f"filter={params['filter']}")
            query_string = '&'.join(query_params)
            
            full_url = f"{url}?{query_string}"
            print(f"Отправляем запрос: URL={full_url}, Headers={headers}")
            
            response = requests.get(full_url, headers=headers)
            
            if response.status_code != 200:
                error_message = f"Ошибка при получении данных: {response.status_code}. Ответ сервера: {response.text}"
                print(error_message)
                raise Exception(error_message)
            
            data = response.json()
            
            if total_count is None:
                total_count = data.get('meta', {}).get('size', 0)
                print(f"Всего записей: {total_count}")
            
            all_rows.extend(data.get('rows', []))
            
            if len(all_rows) >= total_count:
                break
            
            params['offset'] += params['limit']
        
        return {'meta': data.get('meta', {}), 'rows': all_rows}
        
    except Exception as e:
        if str(e) == "Processing cancelled by user":
            abort(499, description="Processing cancelled by user")
        raise e

def get_sales_speed(variant_id, store_id, start_date, end_date, is_variant):
    url = f"{BASE_URL}/report/turnover/byoperations"
    headers = {
        'Authorization': f'Bearer {MOYSKLAD_TOKEN}',
        'Accept': 'application/json;charset=utf-8'
    }
    
    # Преобразуем даты в datetime объекты
    end_datetime = datetime.strptime(end_date, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
    original_start = datetime.strptime(start_date, '%Y-%m-%d').replace(hour=0, minute=0, second=0)
    
    # Устанавливаем начальную дату на 5 лет раньше
    extended_start = original_start - timedelta(days=5*365)
    
    # Форматируем даты для API
    start_date_formatted = extended_start.strftime('%Y-%m-%d %H:%M:%S')
    end_date_formatted = end_datetime.strftime('%Y-%m-%d %H:%M:%S')
    
    params = {
        'momentFrom': start_date_formatted,
        'momentTo': end_date_formatted,
        'limit': 1000,
        'order': 'moment,asc'
    }
    
    assortment_type = 'variant' if is_variant else 'product'
    filter_params = [
        f"filter=store={BASE_URL}/entity/store/{store_id}",
        f"filter={assortment_type}={BASE_URL}/entity/{assortment_type}/{variant_id}"
    ]
    
    query_string = '&'.join([f"{k}={v}" for k, v in params.items()] + filter_params)
    full_url = f"{url}?{query_string}"
    
    print(f"Запрос для получения операций: URL={full_url}")
    
    response = requests.get(full_url, headers=headers)
    if response.status_code != 200:
        print(f"Ошибка при получении данных: {response.status_code}. Ответ сервера: {response.text}")
        return 0, '', '', '', ''

    data = response.json()
    
    # Фильтруем строки для нужной модификации
    rows = [
        row for row in data.get('rows', [])
        if row['assortment']['meta']['href'].split('/')[-1] == variant_id
    ]
    
    if not rows:
        return 0, '', '', '', ''
    
    print(f"Всего операций после фильтрации по variant_id {variant_id}: {len(rows)}")
    print(f"Название модификации: {rows[0]['assortment']['name']}")
    
    # Получаем метаданные группы и товара
    group_uuid = ''
    group_name = ''
    product_uuid = ''
    product_href = ''
    
    assortment = rows[0].get('assortment', {})
    product_folder = assortment.get('productFolder', {})
    group_href = product_folder.get('meta', {}).get('href', '')
    group_uuid = group_href.split('/')[-1] if group_href else ''
    group_name = product_folder.get('name', '')
    
    product_meta = assortment.get('meta', {})
    product_href = product_meta.get('uuidHref', '')
    if product_href:
        product_uuid = product_meta.get('href', '').split('/')[-1]
    
    # Сортировка операций по дате
    rows.sort(key=lambda x: datetime.fromisoformat(x['operation']['moment'].replace('Z', '+00:00')))
    
    # Отслеживаем каждую единицу товара
    stock_items = []  # [(arrival_time, sold_time)] для каждой единицы товара
    current_stock = []  # [(arrival_time, quantity)]
    retail_demand_counter = 0
    on_stock_time = timedelta()
    
    # Проходим по всем операциям для построения полной картины движения товара
    for row in rows:
        operation_time = datetime.fromisoformat(row['operation']['moment'].replace('Z', '+00:00'))
        quantity = row['quantity']
        operation_type = row['operation']['meta']['type']
        
        if quantity > 0:  # Приход товара
            current_stock.append((operation_time, quantity))
            print(f"Приход товара: {quantity} шт. в {operation_time}")
            
        elif operation_type == 'retaildemand' and original_start <= operation_time <= end_datetime:
            # Розничная продажа в указанном периоде
            quantity_to_sell = abs(quantity)
            print(f"Розничная продажа: {quantity_to_sell} шт. в {operation_time}")
            
            # Ищем товар для продажи в порядке FIFO
            while quantity_to_sell > 0 and current_stock:
                arrival_time, available_quantity = current_stock[0]
                sold_from_batch = min(quantity_to_sell, available_quantity)
                
                # Добавляем время на складе для каждой проданной единицы
                time_on_stock = operation_time - arrival_time
                on_stock_time += time_on_stock * sold_from_batch
                print(f"Продано {sold_from_batch} шт. из партии от {arrival_time}")
                print(f"Время на складе: {time_on_stock} × {sold_from_batch} шт.")
                
                retail_demand_counter += sold_from_batch
                quantity_to_sell -= sold_from_batch
                
                # Обновляем или удаляем партию
                if sold_from_batch == available_quantity:
                    current_stock.pop(0)
                else:
                    current_stock[0] = (arrival_time, available_quantity - sold_from_batch)
                    
        elif quantity < 0:  # Другие операции ухода товара
            # Просто уменьшаем количество в stock по FIFO
            quantity_to_remove = abs(quantity)
            while quantity_to_remove > 0 and current_stock:
                if current_stock[0][1] <= quantity_to_remove:
                    quantity_to_remove -= current_stock[0][1]
                    current_stock.pop(0)
                else:
                    current_stock[0] = (current_stock[0][0], current_stock[0][1] - quantity_to_remove)
                    quantity_to_remove = 0
    
    days_on_stock = on_stock_time.total_seconds() / (24 * 60 * 60)
    
    if days_on_stock > 0:
        sales_speed = retail_demand_counter / days_on_stock
    else:
        sales_speed = 0
    
    print(f"Вариант товара: {variant_id}")
    print(f"Период на складе (дней): {days_on_stock}")
    print(f"Количество розничных продаж: {retail_demand_counter}")
    print(f"Скорость продаж: {sales_speed}")
    
    return sales_speed, group_uuid, group_name, product_uuid, product_href

def get_group_path(group_uuid, product_groups, get_uuid=False):
    def find_group_path(groups, target_uuid, current_path=[], current_uuid_path=[]):
        for group in groups:
            if group['id'] == target_uuid:
                return (current_path + [group['name']], current_uuid_path + [group['id']])
            if group.get('children'):
                path = find_group_path(group['children'], target_uuid, 
                                     current_path + [group['name']], 
                                     current_uuid_path + [group['id']])
                if path:
                    return path
        return None

    path = find_group_path(product_groups, group_uuid)
    if not path:
        return '', []  # Возвращаем пустую строку и пустой список UUID
    
    names_path, uuid_path = path
    return ('/'.join(names_path), uuid_path) if not get_uuid else ('/'.join(uuid_path), uuid_path)

def create_excel_report(data, store_id, start_date, end_date, planning_days, manual_stock_settings=None):
    try:
        print("Начало создания Excel отчета")
        print(f"Полученные настройки минимальных остатков: {manual_stock_settings}")
        
        global current_status
        
        wb = Workbook()
        ws = wb.active
        
        product_groups = get_product_groups()
        products_data = []
        max_depth = 0
        
        # Сначала собираем все данные и определяем максимальную глубину
        for item in data['rows']:
            if check_if_cancelled():
                wb.close()
                return None
                
            assortment = item.get('assortment', {})
            assortment_meta = assortment.get('meta', {})
            assortment_href = assortment_meta.get('href', '')
            
            is_variant = '/variant/' in assortment_href
            variant_id = assortment_href.split('/variant/')[-1] if is_variant else assortment_href.split('/product/')[-1]
            
            if variant_id and item.get('sellQuantity', 0) > 0:  # Проверяем продажи сразу
                print(f"Обработка позиции {assortment.get('name', '')} (sellQuantity: {item.get('sellQuantity', 0)})")
                
                # Передаем обе даты в функцию get_sales_speed
                sales_speed, group_uuid, group_name, product_uuid, product_href = get_sales_speed(
                    variant_id, store_id, start_date, end_date, is_variant
                )
                
                if sales_speed is not None:  # Проверяем, что скорость продаж успешно рассчитана
                    full_path, uuid_path = get_group_path(group_uuid, product_groups)
                    max_depth = max(max_depth, len(uuid_path))
                    
                    # Находим количество знаков после запятой для отображения
                    if sales_speed > 0:
                        decimal_str = str(sales_speed).split('.')[-1]
                        non_zero_count = 0
                        for i, digit in enumerate(decimal_str):
                            if digit != '0':
                                non_zero_count = i + 2
                                break
                        display_sales_speed = round(sales_speed, non_zero_count)
                    else:
                        display_sales_speed = 0
                    
                    products_data.append({
                        'name': assortment.get('name', ''),
                        'quantity': item.get('sellQuantity', 0),
                        'profit': round(item.get('profit', 0) / 100, 2),
                        'sales_speed': display_sales_speed,
                        'forecast': sales_speed * planning_days,
                        'group_uuid': group_uuid,
                        'group_path': full_path,
                        'uuid_path': uuid_path,
                        'names_by_level': get_names_by_uuid(uuid_path, product_groups),
                        'product_uuid': product_uuid,
                        'product_href': product_href
                    })
                    
                    print(f"Позиция успешно обработана, обновляем счетчик")
                    update_processed_count()
                    
                    if check_if_cancelled():
                        wb.close()
                        return None

        # Проверяем отмену перед форматированием
        if check_if_cancelled():
            wb.close()
            return None
            
        # Продолжаем форматирование и сохранение файла...
        # [оставшийся код функции без изменений]
        
        filename = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        wb.save(filename)
        wb.close()
        return filename
        
    except Exception as e:
        try:
            wb.close()
        except:
            pass
        raise e

def get_sheet_name(products_data):
    # Получаем уникальные названия второго уровня
    level2_names = set()
    for product in products_data:
        names_by_level = product['names_by_level']
        if len(names_by_level) > 1:  # Если есть второй уровень
            level2_names.add(names_by_level[1])
    
    # Сортируем имена для консистентности
    level2_names = sorted(list(level2_names))
    
    # Формируем название листа с новым разделителем
    if len(level2_names) > 5:
        sheet_name = ', '.join(level2_names[:5])
    else:
        sheet_name = ', '.join(level2_names)
    
    # Ограничиваем длину названия листа (максимум 31 символ в Excel)
    if len(sheet_name) > 31:
        sheet_name = sheet_name[:28] + "..."
    
    # Если название пустое, используем значение по умолчанию 
    return sheet_name if sheet_name else "Отчет прибыльности"

def get_names_by_uuid(uuid_path, product_groups):
    def find_name_by_uuid(groups, target_uuid):
        for group in groups:
            if group['id'] == target_uuid:
                return group['name']
            if group.get('children'):
                name = find_name_by_uuid(group['children'], target_uuid)
                if name:
                    return name
        return None

    return [find_name_by_uuid(product_groups, uuid) or '' for uuid in uuid_path]

@app.route('/cancel', methods=['POST'])
def cancel_processing():
    global processing_cancelled
    with processing_lock:
        processing_cancelled = True
    return jsonify({'status': 'cancelled', 'cancelled': True}), 200  # Возвращаем 200 вместо 499

@app.route('/status-stream')
def status_stream():
    def generate():
        last_processed = -1
        while True:
            with processing_lock:
                if current_status['total'] == 0:  # Ждем инициализации
                    yield f"data: ...\n\n"
                else:
                    remaining = current_status['total'] - current_status['processed']
                    if current_status['processed'] != last_processed:
                        last_processed = current_status['processed']
                        yield f"data: {remaining}\n\n"
                    if remaining <= 0:
                        break
            # Добавляем небольшую задержку
            sleep(0.1)
    return Response(generate(), mimetype='text/event-stream')

def update_processed_count():
    """Обновляет счетчик обработанных записей"""
    global current_status
    with processing_lock:
        if current_status['total'] > 0:  # Проверяем, что счетчик инициализирован
            current_status['processed'] += 1
            remaining = current_status['total'] - current_status['processed']
            print(f"Обновлен счетчик: обработано {current_status['processed']} из {current_status['total']}, осталось {remaining}")

if __name__ == '__main__':
    print("Starting Flask iframe app...")
    app.run(debug=True, port=5001)  # Используем другой порт, чтобы не конфликтовать с основным приложением 