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

# Add this near the top of the file after app creation
port = int(os.environ.get("PORT", 10000))

# Загрузка конфигурации
with open('config.py', 'r') as config_file:
    exec(config_file.read())

BASE_URL = 'https://api.moysklad.ru/api/remap/1.2'

processing_cancelled = False
processing_lock = threading.Lock()
current_status = {'total': 0, 'processed': 0}
api_request_times = []

# ... rest of your existing code ...

# Modify the if __name__ == '__main__': block at the bottom
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port)
