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
from time import sleep, time
from openpyxl.styles.colors import Color

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
        return send_file(filename, as_attachment=True, download_name=filename)
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
                            
                            // Получаем значение search_days из iframe
                            const iframe = document.querySelector('iframe');
                            const searchDays = parseInt(iframe.contentWindow.document.getElementById('search_days').value) || 300;
                            
                            // Рассчитываем время на одну позицию в зависимости от количества дней
                            let timePerItem;
                            if (searchDays <= 300) {
                                timePerItem = 1.4; // базовое время для 300 дней
                            } else if (searchDays <= 365) {
                                // Для диапазона 300-365 дней: ~0.0062 сек/день
                                timePerItem = 1.4 + ((searchDays - 300) * 0.0062);
                            } else {
                                // Для диапазона более 365 дней: первые 65 дней по 0.0062 сек/день, остальные по 0.0025 сек/день
                                timePerItem = 1.4 + (65 * 0.0062) + ((searchDays - 365) * 0.0025);
                            }
                            
                            // Для меньшего количества дней пропорционально уменьшаем время
                            if (searchDays < 300) {
                                timePerItem = (timePerItem * searchDays) / 300;
                            }
                            
                            const totalSeconds = Math.ceil(remainingNum * timePerItem);
                            
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
            
            try:
                response = requests.get(full_url, headers=headers, timeout=30)
            except requests.exceptions.Timeout:
                print("Timeout при получении данных отчета")
                raise Exception("Timeout при получении данных отчета")
            except requests.exceptions.RequestException as e:
                print(f"Ошибка при получении данных отчета: {str(e)}")
                raise e
            
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
        
        # Собираем все variant_ids для одного запроса
        variant_ids = []
        for item in data.get('rows', []):
            assortment = item.get('assortment', {})
            assortment_meta = assortment.get('meta', {})
            assortment_href = assortment_meta.get('href', '')
            
            is_variant = '/variant/' in assortment_href
            variant_id = assortment_href.split('/variant/')[-1] if is_variant else assortment_href.split('/product/')[-1]
            if variant_id:
                variant_ids.append(variant_id)
        
        # Делаем один запрос для всех вариантов
        operations_data = get_bulk_operations(variant_ids, store_id, start_date, end_date)
        
        return {'meta': data.get('meta', {}), 'rows': all_rows}
        
    except Exception as e:
        if str(e) == "Processing cancelled by user":
            abort(499, description="Processing cancelled by user")
        raise e

def get_bulk_operations(variant_ids, store_id, start_date, end_date):
    url = f"{BASE_URL}/report/turnover/byoperations"
    headers = {
        'Authorization': f'Bearer {MOYSKLAD_TOKEN}',
        'Accept': 'application/json;charset=utf-8'
    }
    
    # Формируем фильтр для всех вариантов через OR
    variant_filters = [f"variant={BASE_URL}/entity/variant/{vid}" for vid in variant_ids]
    filter_query = " || ".join(variant_filters)
    
    params = {
        'momentFrom': start_date,
        'momentTo': end_date,
        'filter': f"store={BASE_URL}/entity/store/{store_id};({filter_query})",
        'limit': 1000
    }
    
    return requests.get(url, headers=headers, params=params).json()

def get_sales_speed(variant_id, store_id, start_date, end_date, is_variant):
    print(f"\nНачало расчета скорости продаж для варианта {variant_id}")
    total_start_time = time()
    
    url = f"{BASE_URL}/report/turnover/byoperations"
    headers = {
        'Authorization': f'Bearer {MOYSKLAD_TOKEN}',
        'Accept': 'application/json;charset=utf-8'
    }
    
    # Преобразуем даты в datetime объекты
    date_conversion_start = time()
    end_datetime = datetime.strptime(end_date, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
    original_start = datetime.strptime(start_date, '%Y-%m-%d').replace(hour=0, minute=0, second=0)
    # Используем значение из параметра search_days
    search_days = int(request.form.get('search_days', 300))
    extended_start = original_start - timedelta(days=search_days)
    start_date_formatted = extended_start.strftime('%Y-%m-%d %H:%M:%S')
    end_date_formatted = end_datetime.strftime('%Y-%m-%d %H:%M:%S')
    print(f"Время на конвертацию дат: {(time() - date_conversion_start):.3f} сек")
    
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
    
    # Замер времени API запроса
    api_request_start = time()
    try:
        response = requests.get(full_url, headers=headers, timeout=30)
    except requests.exceptions.Timeout:
        print(f"Timeout при запросе операций для варианта {variant_id}")
        return 0, '', '', '', ''
    except requests.exceptions.RequestException as e:
        print(f"Ошибка при запросе операций для варианта {variant_id}: {str(e)}")
        return 0, '', '', '', ''
    
    api_request_time = time() - api_request_start
    print(f"Время выполнения API запроса: {api_request_time:.3f} сек")
        
    if response.status_code != 200:
        print(f"Ошибка при получении данных: {response.status_code}. Ответ сервера: {response.text}")
        return 0, '', '', '', ''

    # Замер времени парсинга JSON
    json_parse_start = time()
    data = response.json()
    json_parse_time = time() - json_parse_start
    print(f"Время парсинга JSON: {json_parse_time:.3f} сек")
    
    if not data or 'rows' not in data:
        print(f"Получен пустой ответ для варианта {variant_id}")
        return 0, '', '', '', ''
    
    # Замер времени фильтрации строк
    filter_start = time()
    rows = [
        row for row in data.get('rows', [])
        if row['assortment']['meta']['href'].split('/')[-1] == variant_id
    ]
    filter_time = time() - filter_start
    print(f"Время фильтрации строк: {filter_time:.3f} сек")
    
    if not rows:
        return 0, '', '', '', ''
    
    # Получение метаданных
    metadata_start = time()
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
    metadata_time = time() - metadata_start
    print(f"Время обработки метаданных: {metadata_time:.3f} сек")
    
    # Замер времени сортировки
    sort_start = time()
    rows.sort(key=lambda x: datetime.fromisoformat(x['operation']['moment'].replace('Z', '+00:00')))
    sort_time = time() - sort_start
    print(f"Время сортировки операций: {sort_time:.3f} сек")
    
    # Замер времени расчета скорости продаж
    calculation_start = time()
    stock_items = []
    current_stock = []
    retail_demand_counter = 0
    on_stock_time = timedelta()
    
    for row in rows:
        operation_time = datetime.fromisoformat(row['operation']['moment'].replace('Z', '+00:00'))
        quantity = row['quantity']
        operation_type = row['operation']['meta']['type']
        
        if quantity > 0:
            current_stock.append((operation_time, quantity))
            
        elif operation_type == 'retaildemand' and original_start <= operation_time <= end_datetime:
            quantity_to_sell = abs(quantity)
            
            while quantity_to_sell > 0 and current_stock:
                arrival_time, available_quantity = current_stock[0]
                sold_from_batch = min(quantity_to_sell, available_quantity)
                
                time_on_stock = operation_time - arrival_time
                on_stock_time += time_on_stock * sold_from_batch
                
                retail_demand_counter += sold_from_batch
                quantity_to_sell -= sold_from_batch
                
                if sold_from_batch == available_quantity:
                    current_stock.pop(0)
                else:
                    current_stock[0] = (arrival_time, available_quantity - sold_from_batch)
                    
        elif quantity < 0:
            quantity_to_remove = abs(quantity)
            while quantity_to_remove > 0 and current_stock:
                if current_stock[0][1] <= quantity_to_remove:
                    quantity_to_remove -= current_stock[0][1]
                    current_stock.pop(0)
                else:
                    current_stock[0] = (current_stock[0][0], current_stock[0][1] - quantity_to_remove)
                    quantity_to_remove = 0
    
    days_on_stock = on_stock_time.total_seconds() / (24 * 60 * 60)
    sales_speed = retail_demand_counter / days_on_stock if days_on_stock > 0 else 0
    
    calculation_time = time() - calculation_start
    print(f"Время расчета скорости продаж: {calculation_time:.3f} сек")
    
    total_time = time() - total_start_time
    print(f"\nОбщее время выполнения get_sales_speed: {total_time:.3f} сек")
    print(f"Разбивка времени выполнения:")
    print(f"- API запрос: {api_request_time:.3f} сек ({(api_request_time/total_time*100):.1f}%)")
    print(f"- Парсинг JSON: {json_parse_time:.3f} сек ({(json_parse_time/total_time*100):.1f}%)")
    print(f"- Фильтрация: {filter_time:.3f} сек ({(filter_time/total_time*100):.1f}%)")
    print(f"- Метаданные: {metadata_time:.3f} сек ({(metadata_time/total_time*100):.1f}%)")
    print(f"- Сортировка: {sort_time:.3f} сек ({(sort_time/total_time*100):.1f}%)")
    print(f"- Расчет: {calculation_time:.3f} сек ({(calculation_time/total_time*100):.1f}%)")
    
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

def calculate_group_quantities(products_data):
    """
    Рассчитывает суммарное количество для каждой группы на основе всех товаров в ней
    и её подгруппах.
    """
    group_quantities = {}  # uuid -> quantity

    # Сначала собираем все товары по группам
    for product in products_data:
        # Для каждого уровня в пути группы
        for i in range(len(product['uuid_path'])):
            group_uuid = product['uuid_path'][i]
            if group_uuid not in group_quantities:
                group_quantities[group_uuid] = 0
            # Добавляем количество товара к каждой группе в пути
            group_quantities[group_uuid] += product['quantity']

    return group_quantities

def calculate_group_profits(products_data):
    """
    Рассчитывает среднюю прибыльность для каждой группы на основе всех товаров в ней
    и её подгруппах.
    """
    group_profits = {}  # uuid -> (total_profit, count)

    # Сначала собираем все товары по группам
    for product in products_data:
        # Для каждого уровня в пути группы
        for i in range(len(product['uuid_path'])):
            group_uuid = product['uuid_path'][i]
            if group_uuid not in group_profits:
                group_profits[group_uuid] = [0, 0]  # [сумма прибыли, количество товаров]
            # Добавляем прибыль товара и увеличиваем счетчик
            group_profits[group_uuid][0] += product['profit']
            group_profits[group_uuid][1] += 1

    # Вычисляем средние значения
    average_profits = {}
    for group_uuid, (total_profit, count) in group_profits.items():
        average_profits[group_uuid] = round(total_profit / count, 2) if count > 0 else 0

    return average_profits

def calculate_group_sales_speed(products_data):
    """
    Рассчитывает среднюю скорость продаж для каждой группы на основе всех товаров в ней
    и её подгруппах.
    """
    group_speeds = {}  # uuid -> (total_speed, count)

    # Сначала собираем все товары по группам
    for product in products_data:
        # Для каждого уровня в пути группы
        for i in range(len(product['uuid_path'])):
            group_uuid = product['uuid_path'][i]
            if group_uuid not in group_speeds:
                group_speeds[group_uuid] = [0, 0]  # [сумма скоростей, количество товаров]
            # Добавляем скорость продаж товара и увеличиваем счетчик
            group_speeds[group_uuid][0] += product['sales_speed']
            group_speeds[group_uuid][1] += 1

    # Вычисляем средние значения
    average_speeds = {}
    for group_uuid, (total_speed, count) in group_speeds.items():
        average_speeds[group_uuid] = round(total_speed / count, 2) if count > 0 else 0

    return average_speeds

def calculate_group_profitability(products_data):
    """
    Рассчитывает среднюю прибыльность группы на основе всех товаров в ней
    и её подгруппах.
    """
    group_profitability = {}  # uuid -> (total_profitability, count)

    # Сначала собираем все товары по группам
    for product in products_data:
        # Для каждого уровня в пути группы
        for i in range(len(product['uuid_path'])):
            group_uuid = product['uuid_path'][i]
            if group_uuid not in group_profitability:
                group_profitability[group_uuid] = [0, 0]  # [сумма прибыльности группы, количество товаров]
            # Добавляем прибыльность группы (произведение прибыли на скорость) и увеличиваем счетчик
            group_profitability[group_uuid][0] += (product['profit'] * product['sales_speed'])
            group_profitability[group_uuid][1] += 1

    # Вычисляем средние значения
    average_profitability = {}
    for group_uuid, (total_profitability, count) in group_profitability.items():
        average_profitability[group_uuid] = round(total_profitability / count, 2) if count > 0 else 0

    return average_profitability

def create_hierarchical_sort_key(product):
    """
    Создает ключ сортировки, который обеспечивает правильное иерархическое отображение.
    Сортирует группы по возрастанию названия.
    """
    path_components = []
    names_by_level = product['names_by_level']
    
    for i, (uuid, name) in enumerate(zip(product['uuid_path'], names_by_level)):
        # Создаем кортеж из уровня, имени и uuid для сортировки
        path_components.append((i, name, uuid))
    
    return path_components

def create_excel_report(data, store_id, start_date, end_date, planning_days, manual_stock_settings=None):
    try:
        print("Начало создания Excel отчета")
        print(f"Полученные настройки минимальных остатков: {manual_stock_settings}")
        
        global current_status
        
        wb = Workbook()
        ws = wb.active
        
        # Определяем палитру цветов в начале
        color_palette = ['F2F2F2', 'E6E6E6', 'D9D9D9', 'CCCCCC', 'BFBFBF']
        
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
            
        print(f"Максимальная глубина групп: {max_depth}")

        # Рассчитываем количества и средние значения для всех групп
        group_quantities = calculate_group_quantities(products_data)
        group_profits = calculate_group_profits(products_data)
        group_sales_speeds = calculate_group_sales_speed(products_data)
        group_profitability = calculate_group_profitability(products_data)

        # Сортируем данные с использованием иерархического ключа
        products_data.sort(key=create_hierarchical_sort_key)

        # Формируем заголовки с учетом реальной глубины, начиная со второго уровня
        group_level_headers = [f'Уровень {i+2}' for i in range(max_depth-1)] if max_depth > 1 else []
        headers = group_level_headers + [
            'UUID',  # Изменено название столбца
            'Наименование', 'Количество проданного', 'Средняя прибыльность товара', 'Скорость продаж', 
            'Прибыльность группы', '30', 'Мин.остаток'
        ]
        
        # Записываем заголовки
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(color="000000", bold=True)
            # Применяем форматирование для всех заголовков
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Инициализируем current_row здесь
        current_row = 2
        current_uuid_path = []
        written_groups = set()  # Множество для отслеживания уже записанных групп

        # Сначала записываем все данные для листа "Анализ1"
        for product in products_data:
            uuid_path = product['uuid_path']
            names_by_level = product['names_by_level']
            
            # Записываем строки групп, если путь изменился
            for i, uuid in enumerate(uuid_path):
                group_key = f"{i}_{uuid}"  # Создаем уникальный ключ для группы с учетом уровня
                if (i >= len(current_uuid_path) or uuid != current_uuid_path[i]) and group_key not in written_groups:
                    if i > 0:
                        ws.cell(row=current_row, column=i, value=names_by_level[i])
                        # Записываем количество для группы
                        ws.cell(row=current_row, column=max_depth+2, value=group_quantities.get(uuid, 0))
                        # Записываем среднюю прибыльность для группы
                        profit_cell = ws.cell(row=current_row, column=max_depth+3, value=group_profits.get(uuid, 0))
                        profit_cell.number_format = '0.00'
                        # Записываем среднюю скорость продаж для группы
                        speed_cell = ws.cell(row=current_row, column=max_depth+4, value=group_sales_speeds.get(uuid, 0))
                        speed_cell.number_format = '0.00'
                        # Записываем среднюю прибыльность группы
                        group_profit_cell = ws.cell(row=current_row, column=max_depth+5, value=group_profitability.get(uuid, 0))
                        group_profit_cell.number_format = '0.00'
                        
                        color_index = min(i - 1, len(color_palette) - 1)
                        fill_color = color_palette[color_index]
                        for col in range(1, ws.max_column + 1):
                            cell = ws.cell(row=current_row, column=col)
                            cell.fill = PatternFill(start_color=fill_color, 
                                                  end_color=fill_color, 
                                                  fill_type='solid')
                        
                        if current_row > 2:
                            uuid_cell = ws.cell(row=current_row, column=max_depth, value=uuid)
                            uuid_cell.alignment = Alignment(horizontal='left', shrink_to_fit=False)
                        
                        current_row += 1
                        written_groups.add(group_key)
            
            # При записи UUID товара
            if current_row > 2:
                uuid_cell = ws.cell(row=current_row, column=max_depth)
                if product['product_href']:
                    uuid_cell.value = product['product_uuid']
                    uuid_cell.hyperlink = product['product_href']
                    uuid_cell.font = Font(color="0000FF", underline="single")
                    uuid_cell.alignment = Alignment(horizontal='left', shrink_to_fit=False)
            
            # Записываем данные продукта
            ws.cell(row=current_row, column=max_depth+1, value=product['name'])
            ws.cell(row=current_row, column=max_depth+2, value=product['quantity'])
            profit_cell = ws.cell(row=current_row, column=max_depth+3, value=product['profit'])
            profit_cell.number_format = '0.00'
            speed_cell = ws.cell(row=current_row, column=max_depth+4, value=product['sales_speed'])
            speed_cell.number_format = '0.00'
            group_profit = round(product['profit'] * product['sales_speed'], 2)
            group_profit_cell = ws.cell(row=current_row, column=max_depth+5, value=group_profit)
            group_profit_cell.number_format = '0.00'
            
            current_row += 1
            current_uuid_path = uuid_path

        # Теперь, когда у нас есть все данные, добавляем формулы
        name_col = get_column_letter(max_depth+1)      # Столбец "Наименование"
        quantity_col = get_column_letter(max_depth+2)  # Столбец "Количество проданного"
        profit_col = get_column_letter(max_depth+3)    # Столбец "Средняя прибыльность товара"
        speed_col = get_column_letter(max_depth+4)     # Столбец "Скорость продаж"
        group_profit_col = get_column_letter(max_depth+5)  # Столбец "Прибыльность группы"
        forecast_col_letter = get_column_letter(max_depth+6)  # Столбец "30"
        round_col_letter = get_column_letter(max_depth+7)     # Столбец "Мин.остаток"

        # Устанавливаем значение 30 в первой ячейке столбца forecast
        ws.cell(row=1, column=max_depth+6, value=30)
        
        # Добавляем формулы в каждую строку, где есть значение в столбце "Наименование"
        for row in range(2, current_row):
            name_cell = ws.cell(row=row, column=max_depth+1).value
            if name_cell:
                # Формула для столбца "30" с абсолютной ссылкой на ячейку с числом
                forecast_formula = f'={speed_col}{row}*{forecast_col_letter}$1'
                forecast_cell = ws.cell(row=row, column=max_depth+6, value=forecast_formula)
                forecast_cell.number_format = '0.00'
                
                # Формула для столбца округления вверх
                round_formula = f'=CEILING({forecast_col_letter}{row})'
                round_cell = ws.cell(row=row, column=max_depth+7, value=round_formula)
                round_cell.number_format = '0.00'

        # После записи всех данных и перед форматированием добавляем группировку
        ws.sheet_properties.outlinePr.summaryBelow = False  # Устанавливаем кнопку группировки сверху

        # Функция для определения диапазонов групп
        def find_groups(ws, start_row, end_row, max_depth):
            groups = []  # [(start_row, end_row, level, group_name)]
            
            # Проходим по каждой строке
            for row in range(start_row, end_row + 1):
                # Проверяем каждый уровень
                for level in range(1, max_depth + 1):
                    value = ws.cell(row=row, column=level).value
                    if value is not None:
                        # Находим конец группы (последнюю строку перед следующей группой того же или более высокого уровня)
                        end_group_row = row
                        for next_row in range(row + 1, end_row + 1):
                            # Проверяем, не началась ли новая группа того же или более высокого уровня
                            found_higher_level = False
                            for check_level in range(1, level + 1):
                                if ws.cell(row=next_row, column=check_level).value is not None:
                                    found_higher_level = True
                                    break
                            if found_higher_level:
                                end_group_row = next_row - 1
                                break
                            end_group_row = next_row
                        
                        groups.append((row, end_group_row, level, value))
            
            return groups

        # Находим все группы
        groups = find_groups(ws, 2, current_row - 1, max_depth)

        # Сортируем группы по уровню (от большего к меньшему)
        # и по позиции (сверху вниз)
        groups.sort(key=lambda x: (-x[2], x[0]))

        # Применяем группировку
        for start_row, end_row, level, group_name in groups:
            if start_row < end_row:  # Группируем если есть что группировать
                # Группируем все строки под группой
                for row in range(start_row + 1, end_row + 1):
                    current_level = ws.row_dimensions[row].outline_level
                    ws.row_dimensions[row].outline_level = current_level + 1 if current_level is not None else 1
                    ws.row_dimensions[row].hidden = False

        # Отключаем группировку для заголовка
        ws.row_dimensions[1].outline_level = 0 

        # Форматирование
        ws.freeze_panes = ws['A2']  # Закрепляем первую строку
        
        # Устанавливаем ширину для столбца "Наименование"
        ws.column_dimensions[name_col].width = 90

        # Устанавливаем выравнивание по левому краю для заголовков от столбца A до "Наименование"
        for col in range(1, ws[name_col + '1'].column + 1):
            header_cell = ws.cell(row=1, column=col)
            header_cell.alignment = Alignment(horizontal='left', vertical='center', shrink_to_fit=True)

        # Устанавливаем ширину и форматирование для указанных столбцов
        fixed_width_columns = [
            quantity_col, profit_col, speed_col, group_profit_col,
            forecast_col_letter, round_col_letter
        ]
        
        # Устанавливаем ширину столбцов и форматируем все ячейки
        for col_letter in fixed_width_columns:
            ws.column_dimensions[col_letter].width = 15
            # Форматируем заголовок
            header_cell = ws.cell(row=1, column=ws[col_letter + '1'].column)
            header_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            # Форматируем все ячейки в столбце
            for cell in ws[col_letter]:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # Автоподбор ширины для остальных столбцов
        for column in ws.columns:
            column_letter = get_column_letter(column[0].column)
            
            # Пропускаем столбцы с фиксированной шириной
            if column_letter in fixed_width_columns or column_letter == name_col:
                continue
                
            # Если это столбец "UUID" (max_depth)
            if column[0].column == max_depth:
                ws.column_dimensions[column_letter].width = 3
                # Применяем настройки отображения ко всем ячейкам в столбце
                for cell in column:
                    if cell.row == 1:  # For the header cell
                        cell.alignment = Alignment(horizontal='left', vertical='center', shrink_to_fit=True)
                    else:  # For other cells
                        if isinstance(cell.hyperlink, str):  # Если есть ссылка
                            cell.font = Font(color="0000FF", underline="single")
                        cell.alignment = Alignment(horizontal='left', shrink_to_fit=True)
                continue

            max_length = 0
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width

        # После сбора всех данных и перед созданием заголовков
        ws.title = "Анализ1"
        print(f"Название листа: {ws.title}")
        
        # Создаем лист "Анализ2" и копируем в него данные с новой логикой
        ws2 = wb.create_sheet("Анализ2")
        
        # Копируем заголовки и их форматирование, пропуская столбец "Наименование"
        target_col = 1
        for col in range(1, ws.max_column + 1):
            # Пропускаем столбец "Наименование"
            if col == max_depth + 1:
                continue
                
            source_cell = ws.cell(row=1, column=col)
            target_cell = ws2.cell(row=1, column=target_col)
            
            # Копируем значение и форматирование заголовка
            target_cell.value = source_cell.value
            target_cell.font = Font(bold=True, color=Color(rgb='00000000'))  # Черный цвет
            
            # Устанавливаем выравнивание в зависимости от столбца
            if col <= max_depth:  # Для столбцов до UUID включительно
                target_cell.alignment = Alignment(horizontal='left', shrink_to_fit=True)
            else:
                target_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            # Копируем ширину столбца
            source_letter = get_column_letter(col)
            target_letter = get_column_letter(target_col)
            if source_letter in ws.column_dimensions:
                ws2.column_dimensions[target_letter].width = ws.column_dimensions[source_letter].width
            
            target_col += 1
        
        # Закрепляем первую строку
        ws2.freeze_panes = ws2['A2']
        
        # Сбрасываем переменные для второго листа
        target_row = 2
        current_uuid_path = []
        written_groups = set()
        
        # Копируем данные из первого листа
        for source_row in range(2, current_row):
            # Проверяем, есть ли в строке данные в столбцах "Уровень №"
            has_group_data = False
            for col in range(1, max_depth):
                if ws.cell(row=source_row, column=col).value is not None:
                    has_group_data = True
                    break
            
            # Копируем только строки с группами
            if has_group_data:
                # Копируем всю строку с данными и форматированием
                target_col = 1
                for col in range(1, ws.max_column + 1):
                    # Пропускаем столбец "Наименование"
                    if col == max_depth + 1:
                        continue
                        
                    source_cell = ws.cell(row=source_row, column=col)
                    target_cell = ws2.cell(row=target_row, column=target_col)
                    
                    # Копируем значение
                    target_cell.value = source_cell.value
                    
                    # Копируем форматирование
                    if source_cell.fill and hasattr(source_cell.fill, 'start_color') and source_cell.fill.start_color:
                        try:
                            fill_color = source_cell.fill.start_color.rgb or 'FFFFFF'
                            target_cell.fill = PatternFill(
                                start_color=fill_color,
                                end_color=fill_color,
                                fill_type='solid'
                            )
                        except:
                            pass
                    
                    # Копируем шрифт
                    try:
                        font_color = None
                        if source_cell.font and source_cell.font.color:
                            if hasattr(source_cell.font.color, 'rgb'):
                                font_color = Color(rgb=source_cell.font.color.rgb)
                            elif isinstance(source_cell.font.color, str):
                                font_color = Color(rgb=source_cell.font.color)
                        
                        target_cell.font = Font(
                            bold=bool(source_cell.font.bold) if source_cell.font else False,
                            color=font_color,
                            underline=source_cell.font.underline if source_cell.font else None
                        )
                    except:
                        target_cell.font = Font(bold=False, color=Color(rgb='00000000'))
                    
                    # Копируем выравнивание
                    if col <= max_depth:  # Для столбцов до UUID включительно
                        target_cell.alignment = Alignment(horizontal='left', shrink_to_fit=True)
                    else:
                        target_cell.alignment = Alignment(
                            horizontal='center',
                            vertical='center',
                            wrap_text=True
                        )
                    
                    # Копируем формат чисел
                    if source_cell.number_format:
                        target_cell.number_format = source_cell.number_format
                    
                    # Копируем формулы, если они есть, корректируя номера столбцов
                    if isinstance(source_cell.value, str) and source_cell.value.startswith('='):
                        try:
                            formula = source_cell.value.replace(str(source_row), str(target_row))
                            # Корректируем ссылки на столбцы в формулах
                            if col > max_depth + 1:
                                old_col = get_column_letter(col)
                                new_col = get_column_letter(target_col)
                                formula = formula.replace(old_col, new_col)
                            target_cell.value = formula
                        except:
                            pass
                    
                    target_col += 1
                
                target_row += 1
        
        # Создаем лист с информацией
        info_ws = wb.create_sheet("Инфо")  # Создаем лист "Инфо" последним в книге, убираем индекс 0
        
        # Получаем название магазина для информационного листа
        stores = get_stores()
        store_name = next((store['name'] for store in stores if store['id'] == store_id), store_id)
        
        # Форматируем даты для отображения
        start_date_obj = datetime.strptime(start_date, '%Y-%m-%d')
        end_date_obj = datetime.strptime(end_date, '%Y-%m-%d')
        start_date_formatted = start_date_obj.strftime('%d.%m.%Y')
        end_date_formatted = end_date_obj.strftime('%d.%m.%Y')
        
        # Подготавливаем данные для информационного листа
        info_data = [
            ["Параметры анализа", ""],
            ["", ""],
            ["Период анализа:", f"с {start_date_formatted} по {end_date_formatted}"],
            ["Количество дней для поиска поступлений на склад до рассматриваемого периода:", f"{request.form.get('search_days', '300')} дней"],
            ["Количество дней анализа:", f"{(end_date_obj - start_date_obj).days + 1} дней"],
            ["", ""],
            ["Магазин:", store_name],
            ["", ""],
            ["Период планирования:", f"{planning_days} дней"],
            ["", ""],
            ["Выбранные группы товаров:", ""],
        ]
        
        # Добавляем информацию о выбранных группах
        selected_groups = []
        for product in products_data:
            if product['group_path'] and product['group_path'] not in selected_groups:
                selected_groups.append(product['group_path'])
        
        # Получаем название группы для имени файла и находим общий путь
        group_name_for_file = "Все группы"
        common_path = ""
        if selected_groups:
            # Разбиваем первый путь на компоненты для начального сравнения
            path_parts = selected_groups[0].split('/')
            
            # Находим общие части пути для всех групп
            for i in range(len(path_parts)):
                current_part = '/'.join(path_parts[:i+1])
                is_common = all(group.startswith(current_part) for group in selected_groups)
                if is_common:
                    common_path = current_part
                else:
                    break
            
            # Если общий путь найден, используем его последнюю часть для имени файла
            if common_path:
                group_name_for_file = common_path.split('/')[-1]
        
        # Добавляем общий путь в информационный лист
        if common_path:
            info_data.append(["", common_path + "/"])
        else:
            info_data.append(["", "Все группы"])
        
        # Записываем данные в лист
        for row_idx, row_data in enumerate(info_data, 1):
            for col_idx, value in enumerate(row_data, 1):
                cell = info_ws.cell(row=row_idx, column=col_idx, value=value)
                # Форматирование для заголовков
                if row_idx == 1 or (col_idx == 1 and value):
                    cell.font = Font(bold=True)
                # Выравнивание
                cell.alignment = Alignment(vertical='center')
                if col_idx == 1:
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
        
        # Устанавливаем ширину столбцов
        info_ws.column_dimensions['A'].width = 30
        info_ws.column_dimensions['B'].width = 60
        
        # Добавляем время создания отчета
        current_time = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
        time_row = len(info_data) + 2
        info_ws.cell(row=time_row, column=1, value="Отчет сформирован:").font = Font(bold=True)
        info_ws.cell(row=time_row, column=1).alignment = Alignment(horizontal='right', vertical='center')
        info_ws.cell(row=time_row, column=2, value=current_time).alignment = Alignment(horizontal='left', vertical='center')
        
        # Делаем активным лист "Анализ1"
        wb.active = wb["Анализ1"]
        
        # Формируем имя файла
        filename = f"{start_date_formatted}-{end_date_formatted} - {group_name_for_file} - {store_name}.xlsx"
        
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