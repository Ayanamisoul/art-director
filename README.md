# Арт-директор

Приложение для подготовки презентаций PowerPoint и HTML-лендингов по текстовому брифу. Для презентации можно загрузить шаблон PPTX или PDF и фотографии. Для лендинга доступны три варианта оформления. Готовые материалы можно открыть в браузере и экспортировать в PDF.

Учебный проект Московского Политеха для Актиона.

## Запуск в Windows

Команды выполняются в PowerShell из корня проекта. Для генерации текста нужен ключ Z.ai. Экспорт презентаций в PDF и предпросмотр слайдов используют установленный Microsoft PowerPoint. Для экспорта лендингов в PDF нужен Google Chrome.

1. Создайте виртуальное окружение и установите зависимости.

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\python.exe -m pip install fastapi uvicorn python-multipart "mcp==2.3.0" python-pptx Pillow PyMuPDF
   ```

2. При первом запуске скопируйте `.env.example` в `.env` и заполните `ZAI_API_KEY`. Параметры `ZAI_MODEL` и `ZAI_BASE_URL` позволяют выбрать модель и адрес API.

3. Запустите приложение и откройте [веб-интерфейс](http://127.0.0.1:8000).

   ```powershell
   .\.venv\Scripts\python.exe -m uvicorn web:app --app-dir agent --host 127.0.0.1 --port 8000
   ```

В интерфейсе выберите формат, укажите шаблон презентации или оформление лендинга, добавьте изображения и отправьте бриф в чат. Приложение сохраняет результаты в `data/artifacts/`, загруженные изображения и шаблоны — в `data/assets/` и `data/templates/`, историю чатов — в `data/sessions/`. Git игнорирует каталог `data/` и файл `.env`.

## Проверка

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests
```

Для отдельной проверки подключения MCP и экспорта HTML в PDF выполните `.\.venv\Scripts\python.exe test_mcp.py`.
