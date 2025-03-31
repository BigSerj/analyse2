import shutil
import os

# Создаем директории если их нет
os.makedirs('web/js', exist_ok=True)
os.makedirs('web/css', exist_ok=True)

# Копируем iframe.html как index.html
shutil.copy('templates/iframe.html', 'web/index.html')

# Если есть CSS файлы в templates, копируем их
if os.path.exists('templates/styles.css'):
    shutil.copy('templates/styles.css', 'web/css/styles.css')

# Если есть JS файлы в templates, копируем их
if os.path.exists('templates/script.js'):
    shutil.copy('templates/script.js', 'web/js/script.js') 