# Image to DOCX Backend

Backend-сервис для автоматической подготовки штрих-кодов Ozon FBO к печати на обычном принтере.

Проект был создан для продавцов маркетплейсов (Ozon / FBO), которые получают штрих-коды в формате PDF, но печатают их на обычном A4-принтере.

Сервис:

- принимает изображения и PDF через API
- конвертирует PDF-страницы в изображения
- автоматически поворачивает штрих-коды на 90°
- размещает 4 штрих-кода на одном листе A4
- генерирует готовый DOCX-файл для печати

Это позволяет значительно экономить бумагу и упростить процесс подготовки отправлений.

---

# Problem

При работе по модели FBS продавцы маркетплейсов часто получают штрих-коды в PDF-формате, где каждый штрих-код занимает отдельный лист.

При печати на обычном принтере это приводит к:

- большому расходу бумаги
- неудобству печати
- лишним затратам

Проект автоматически оптимизирует размещение штрих-кодов и подготавливает их к компактной печати.

---

# How It Works

Проект является backend-частью веб-сервиса.

Frontend отправляет файлы через HTTP API, backend:

1. принимает файлы
2. обрабатывает изображения и PDF
3. генерирует DOCX-документ
4. возвращает готовый файл пользователю

Сам backend не предоставляет пользовательский интерфейс.

Для полноценной работы требуется frontend или API-клиент.

---

# Возможности

- Загрузка нескольких изображений и PDF-файлов
- Конвертация PDF-страниц в изображения
- Автоматический поворот изображений
- Генерация DOCX-документа
- Размещение 4 изображений на одном листе A4
- Подготовка штрих-кодов для маркетплейсов
- Поддержка интеграции с frontend через CORS

---

# Технологии

- Python
- Flask
- Flask-CORS
- Pillow
- python-docx
- pdf2image
- Gunicorn

---

# Структура проекта

```text
image-to-docx-backend/
├── app.py
├── requirements.txt
├── Procfile
├── gunicorn.conf.py
└── uploads/
```

---

# Установка проекта

## 1. Клонировать репозиторий

```bash
git clone https://github.com/Timur-Kalabin/image-to-docx-backend.git
cd image-to-docx-backend
```

## 2. Создать виртуальное окружение

```bash
python -m venv venv
```

## 3. Активировать виртуальное окружение

### macOS / Linux

```bash
source venv/bin/activate
```

### Windows

```bash
venv\Scripts\activate
```

## 4. Установить зависимости

```bash
pip install -r requirements.txt
```

---

# Зависимости

```txt
flask
flask-cors
pillow
python-docx
pdf2image
gunicorn
```

> Для работы `pdf2image` может потребоваться установленный Poppler.

## macOS

```bash
brew install poppler
```

## Ubuntu / Debian

```bash
sudo apt update
sudo apt install poppler-utils
```

---

# Локальный запуск

```bash
python app.py
```

Сервер будет доступен по адресу:

```text
http://localhost:5001
```

---

# Запуск через Gunicorn

```bash
gunicorn --config gunicorn.conf.py app:app
```

---

# API

## POST `/api/upload`

Загружает изображения или PDF-файлы и возвращает готовый DOCX-документ.

---

# Request

Тип запроса:

```text
multipart/form-data
```

Поле:

```text
files
```

Поддерживается загрузка нескольких файлов.

---

# Пример cURL-запроса

```bash
curl -X POST http://localhost:5001/api/upload \
  -F "files=@image1.jpg" \
  -F "files=@image2.png" \
  -F "files=@document.pdf" \
  --output Images.docx
```

---

# Успешный ответ

Сервер возвращает файл:

```text
Images.docx
```

MIME type:

```text
application/vnd.openxmlformats-officedocument.wordprocessingml.document
```

---

# Ошибки

## Файлы не выбраны

```json
{
  "error": "Файлы не выбраны"
}
```

## Неподдерживаемый формат

```json
{
  "error": "Нет подходящих изображений или PDF"
}
```

## Ошибка обработки

```json
{
  "error": "Ошибка при обработке: ..."
}
```

---

# Ограничения

- Максимальный размер запроса: 32 MB
- Поддерживаемые форматы:
  - PNG
  - JPG
  - JPEG
  - GIF
  - BMP
  - PDF
- PDF конвертируется в изображения с DPI 200
- Изображения автоматически поворачиваются на 90°
- На одном листе A4 размещается до 4 изображений

---

# CORS

Backend разрешает запросы с frontend-домена:

```text
https://timur-kalabin.github.io
```

---

# Конфигурация Gunicorn

```python
timeout = 120
workers = 1
max_requests = 100
preload_app = True
```

---

# Deployment

Проект подготовлен для запуска через Gunicorn.

## Procfile

```text
web: gunicorn --config gunicorn.conf.py app:app
```
