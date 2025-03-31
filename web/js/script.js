// Глобальные переменные для отслеживания состояния
let isProcessing = false;
let statusCheckInterval = null;

// Инициализация при загрузке страницы
document.addEventListener('DOMContentLoaded', async function() {
    // Загружаем список складов
    await loadStores();
    // Загружаем группы товаров
    await loadProductGroups();
});

// Загрузка списка складов
async function loadStores() {
    try {
        const result = await eel.get_stores()();
        if (result.success) {
            const storeSelect = document.getElementById('store_id');
            storeSelect.innerHTML = '<option value="">Выберите склад</option>';
            result.stores.forEach(store => {
                const option = document.createElement('option');
                option.value = store.id;
                option.textContent = store.name;
                storeSelect.appendChild(option);
            });
        } else {
            showError(result.message);
        }
    } catch (error) {
        showError('Ошибка при загрузке списка складов');
        console.error(error);
    }
}

// Загрузка групп товаров
async function loadProductGroups() {
    try {
        const result = await eel.get_product_groups()();
        if (result.success) {
            // Обновляем дерево групп товаров
            updateProductGroupsTree(result.groups);
        } else {
            showError(result.message);
        }
    } catch (error) {
        showError('Ошибка при загрузке групп товаров');
        console.error(error);
    }
}

// Генерация отчета
async function generateReport(event) {
    event.preventDefault();
    
    if (isProcessing) {
        return;
    }
    
    // Собираем параметры из формы
    const params = {
        start_date: document.getElementById('start_date').value,
        end_date: document.getElementById('end_date').value,
        store_id: document.getElementById('store_id').value,
        planning_days: document.getElementById('planning_days').value,
        product_groups: getSelectedProductGroups()
    };
    
    // Валидация параметров
    if (!validateParams(params)) {
        return;
    }
    
    try {
        isProcessing = true;
        showLoading();
        startStatusCheck();
        
        const result = await eel.generate_report(params)();
        
        if (result.success) {
            if (result.filename && result.filedata) {
                downloadBase64File(result.filedata, result.filename);
                showSuccess(result.message || `Файл ${result.filename} готов к скачиванию`);
            } else {
                showSuccess(result.message || "Отчет сгенерирован");
            }
        } else {
            showError(result.message || "Ошибка при генерации отчета");
        }
    } catch (error) {
        showError('Ошибка при генерации отчета');
        console.error(error);
    } finally {
        isProcessing = false;
        hideLoading();
        stopStatusCheck();
    }
}

// Функция для скачивания файла из base64
function downloadBase64File(base64Data, filename) {
    const byteCharacters = atob(base64Data);
    const byteNumbers = new Array(byteCharacters.length);
    
    for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
}

// Проверка статуса обработки
async function checkStatus() {
    try {
        const status = await eel.get_processing_status()();
        updateProgressDisplay(status);
    } catch (error) {
        console.error('Ошибка при получении статуса:', error);
    }
}

// Запуск проверки статуса
function startStatusCheck() {
    if (!statusCheckInterval) {
        statusCheckInterval = setInterval(checkStatus, 1000);
    }
}

// Остановка проверки статуса
function stopStatusCheck() {
    if (statusCheckInterval) {
        clearInterval(statusCheckInterval);
        statusCheckInterval = null;
    }
}

// Отмена генерации отчета
async function cancelReport() {
    try {
        await eel.cancel_processing()();
        showInfo('Генерация отчета отменена');
    } catch (error) {
        console.error('Ошибка при отмене генерации:', error);
    }
}

// Вспомогательные функции для UI
function showLoading() {
    document.getElementById('loading').style.display = 'block';
    document.getElementById('overlay').style.display = 'block';
}

function hideLoading() {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('overlay').style.display = 'none';
}

function showSuccess(message) {
    // Реализация отображения успешного сообщения
}

function showError(message) {
    // Реализация отображения ошибки
}

function showInfo(message) {
    // Реализация отображения информационного сообщения
}

function updateProgressDisplay(status) {
    // Обновление отображения прогресса
    const progressElement = document.getElementById('progress');
    if (status.total > 0) {
        const percent = Math.round((status.processed / status.total) * 100);
        progressElement.textContent = `Обработано ${status.processed} из ${status.total} (${percent}%)`;
    }
} 