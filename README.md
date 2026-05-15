# Factory Tracker — Инструкция по установке

## Структура проекта

```
factory_tracker/
├── app.py              # Flask-приложение
├── db.py               # SQLite база
├── excel_parser.py     # Парсер Excel из 1С
├── sheets.py           # Запись в Google Sheets
├── requirements.txt    # Python-зависимости
├── service_account.json  # (создаёшь сам — ключ Google)
├── templates/
│   ├── index.html      # Список заказов
│   └── order.html      # Позиции заказа
└── uploads/            # Временная папка для Excel
```

---

## 1. Установка на Orange Pi / Raspberry Pi

```bash
# Обновить систему
sudo apt update && sudo apt upgrade -y

# Python и pip
sudo apt install python3 python3-pip -y

# Установить зависимости
cd ~/factory_tracker
pip3 install -r requirements.txt
```

---

## 2. Google Sheets — настройка service account

1. Открыть https://console.cloud.google.com
2. Создать проект (или выбрать существующий)
3. Включить **Google Sheets API**: APIs & Services → Enable APIs → Google Sheets API
4. Создать service account: APIs & Services → Credentials → Create Credentials → Service Account
5. Дать имя, нажать Done
6. Зайти в созданный service account → вкладка Keys → Add Key → JSON
7. Скачанный файл переименовать в `service_account.json` и положить в папку `factory_tracker/`
8. Скопировать email service account (вида `name@project.iam.gserviceaccount.com`)
9. Открыть Google таблицу → Настройки доступа → вставить этот email с правами **Редактор**

---

## 3. Прописать ID таблицы в sheets.py

Открыть `sheets.py`, найти строку:
```python
SPREADSHEET_ID = "YOUR_SPREADSHEET_ID"
```

Вставить ID из URL таблицы:
```
https://docs.google.com/spreadsheets/d/  ТУТ_ВОТ_ЭТО  /edit
```

---

## 4. Запуск

```bash
cd ~/factory_tracker
python3 app.py
```

Открыть в браузере: `http://localhost:5000`
Или с другого устройства в сети: `http://<IP-адрес-малинки>:5000`

---

## 5. Автозапуск при включении (systemd)

```bash
sudo nano /etc/systemd/system/factory.service
```

Вставить:
```ini
[Unit]
Description=Factory Tracker
After=network.target

[Service]
User=pi
WorkingDirectory=/home/pi/factory_tracker
ExecStart=/usr/bin/python3 /home/pi/factory_tracker/app.py
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl enable factory
sudo systemctl start factory
```

---

## 6. Chromium в kiosk-режиме (тачскрин)

```bash
sudo nano /etc/xdg/autostart/kiosk.desktop
```

Вставить:
```ini
[Desktop Entry]
Type=Application
Name=Kiosk
Exec=chromium-browser --kiosk --noerrdialogs --disable-infobars http://localhost:5000
```

После перезагрузки браузер откроется автоматически на весь экран.

---

## Как использовать

1. Скачать Excel-файл из 1С, закинуть через интерфейс (кнопка загрузки или drag & drop)
2. Выбрать заказ из списка
3. По мере готовности позиций нажимать **+1 готово** или **N шт** (для произвольного кол-ва)
4. Данные сохраняются локально и автоматически уходят в Google Sheets

---

## Если WiFi пропал

Данные сохраняются в локальный SQLite. При восстановлении сети — зайти на `/sync` или перезапустить приложение, несинхронизированные записи отправятся автоматически.
