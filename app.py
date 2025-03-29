# This file is deprecated. All functionality moved to appiframe.py

import os
from flask import Flask
from appiframe import configure_app

app = Flask(__name__)

# Add port configuration
port = int(os.environ.get("PORT", 10000))

# Конфигурируем приложение, используя функции из appiframe
configure_app(app)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port)
