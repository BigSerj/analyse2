# This file is deprecated. All functionality moved to appiframe.py

import os
from flask import Flask, render_template, request, send_file, jsonify, abort, Response

app = Flask(__name__)

# Add this near the top of the file after app creation
port = int(os.environ.get("PORT", 10000))

# ... rest of your existing code ...

# Modify the if __name__ == '__main__': block at the bottom
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port)
