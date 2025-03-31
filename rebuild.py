import os
import shutil
import subprocess

def rebuild_app():
    print("Начинаем пересборку приложения...")
    
    # Удаляем старые файлы сборки
    print("Удаляем старые файлы...")
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    if os.path.exists('build'):
        shutil.rmtree('build')
    for file in os.listdir('.'):
        if file.endswith('.spec'):
            os.remove(file)
            
    # Запускаем PyInstaller
    print("Запускаем сборку...")
    cmd = [
        '.venv/bin/pyinstaller',
        '--windowed',
        '--add-data=web:web',
        '--add-data=config.py:.',
        '--name=appiframe',
        '--clean',
        '--debug=all',
        '--hidden-import=eel',  # Добавляем явную зависимость от eel
        '--collect-all=eel',    # Собираем все файлы eel
        'appiframe.py'
    ]
    
    try:
        subprocess.run(cmd, check=True)
        print("\nСборка успешно завершена!")
        print("Приложение находится в папке dist/appiframe/")
    except subprocess.CalledProcessError as e:
        print(f"\nОшибка при сборке: {e}")
    except Exception as e:
        print(f"\nНеожиданная ошибка: {e}")

if __name__ == '__main__':
    rebuild_app() 