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

def configure_app(app):
    """Конфигурирует Flask приложение, добавляя все необходимые маршруты и функции"""
    # Перемещаем все глобальные переменные внутрь функции
    global processing_cancelled, processing_lock, current_status, api_request_times
    global BASE_URL, MOYSKLAD_TOKEN
    
    # Инициализируем глобальные переменные
    processing_cancelled = False
    processing_lock = threading.Lock()
    current_status = {'total': 0, 'processed': 0}
    api_request_times = []
    
    # Загружаем конфигурацию
    with open('config.py', 'r') as config_file:
        exec(config_file.read())
    
    BASE_URL = 'https://api.moysklad.ru/api/remap/1.2'
    
    # Регистрируем все маршруты
    @app.route('/')
    def index():
        return render_template('iframe.html', 
                             stores=get_stores(),
                             product_groups=get_product_groups(),
                             product_groups_json=json.dumps(prepare_groups_for_js(get_product_groups())),
                             render_group_options=render_group_options)
    
    @app.route('/iframe')
    def iframe():
        return index()
    
    @app.route('/process', methods=['POST'])
    def process():
        try:
            global processing_cancelled, current_status, api_request_times
            with processing_lock:
                processing_cancelled = False
                current_status = {'total': 0, 'processed': 0}
                api_request_times = []  # Сбрасываем список времени запросов
            
            start_date = request.form['start_date']
            end_date = request.form['end_date']
            store_id = request.form['store_id']
            planning_days = int(request.form['planning_days'])
            
            product_groups = []
            if 'final_product_groups' in request.form:
                raw_groups = request.form.get('final_product_groups', '')
                if raw_groups:
                    product_groups = [group.strip() for group in raw_groups.split(',') if group.strip()]
                print(f"Обработанные группы: {product_groups}")
            
            # Проверяем отмену перед получением данных
            if check_if_cancelled():
                return jsonify({'cancelled': True}), 200
                
            try:
                report_data = get_report_data(start_date, end_date, store_id, product_groups)
                print("Получены данные отчета:", report_data is not None)
                if report_data:
                    print(f"Количество строк в отчете: {len(report_data.get('rows', []))}")
            except Exception as e:
                print(f"Ошибка при получении данных отчета: {str(e)}")
                import traceback
                print("Полный стек ошибки:")
                print(traceback.format_exc())
                return jsonify({'error': str(e)}), 500
            
            if not report_data or 'rows' not in report_data or not report_data['rows']:
                return jsonify({'error': 'Нет данных для формирования отчета'}), 404
            
            # Проверяем отмену после получения данных
            if check_if_cancelled():
                return jsonify({'cancelled': True}), 200
                
            try:
                # Считаем только позиции с продажами и вариантами
                total_items = sum(1 for item in report_data['rows'] 
                                if item.get('sellQuantity', 0) > 0 
                                and ('/variant/' in item.get('assortment', {}).get('meta', {}).get('href', '') 
                                    or '/product/' in item.get('assortment', {}).get('meta', {}).get('href', '')))
                
                print(f"Всего позиций для обработки: {total_items}")
                
                with processing_lock:
                    current_status['total'] = total_items
                    current_status['processed'] = 0
                
                excel_file = create_excel_report(report_data, store_id, start_date, end_date, planning_days)
                
                if excel_file is None:
                    return jsonify({'cancelled': True}), 200
                    
                return jsonify({
                    'success': True,
                    'file_url': f'/download/{excel_file}'
                })
                
            except Exception as e:
                print(f"Ошибка при обработке данных: {str(e)}")
                import traceback
                print("Полный стек ошибки:")
                print(traceback.format_exc())
                return jsonify({'error': str(e)}), 500
                
        except Exception as e:
            print(f"Общая ошибка в process: {str(e)}")
            import traceback
            print("Полный стек ошибки:")
            print(traceback.format_exc())
            return jsonify({'error': str(e)}), 500
    
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
                            # Вычисляем среднее время запроса
                            avg_time = sum(api_request_times) / len(api_request_times) if api_request_times else 0.9
                            status_data = {
                                'remaining': remaining,
                                'processed': current_status['processed'],
                                'total': current_status['total'],
                                'avg_request_time': avg_time
                            }
                            yield f"data: {json.dumps(status_data)}\n\n"
                        if remaining <= 0:
                            break
                sleep(0.1)
        return Response(generate(), mimetype='text/event-stream')
    
    @app.route('/download/<filename>')
    def download_file(filename):
        try:
            return send_file(filename, as_attachment=True, download_name=filename)
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    
    # Остальные функции остаются без изменений
    ...

# Остальной код файла остается без изменений

# Заменяем блок if __name__ == '__main__':
if __name__ == '__main__':
    print("This application should be run through a WSGI server")