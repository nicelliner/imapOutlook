# IMAP Outlook Fetcher

Python-скрипт для выгрузки всех писем из почтовых ящиков Outlook через IMAP с OAuth2-аутентификацией. Поддерживает работу с несколькими аккаунтами одновременно.

## Возможности

- Авторизация через OAuth2 (refresh_token + client_id)
- Подключение к Outlook IMAP (outlook.office365.com) по XOAUTH2
- Поддержка нескольких почтовых ящиков в одном конфиге
- Читаемая текстовая выгрузка писем с заголовками, телом и списком вложений
- Автоопределение email из токена (или указание вручную)
- Fallback: если нет plain text — извлекается текст из HTML

## Требования

- Python 3.9+
- Библиотека `requests`
- Azure App Registration с разрешением `IMAP.AccessAsUser.All`
- Действующий `refresh_token`

## Установка

```bash
pip3 install -r requirements.txt
cp config.example.json config.json
```

## Настройка

Отредактируй `config.json`:

```json
{
    "output_dir": "./emails",
    "accounts": [
        {
            "client_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            "refresh_token": "M.C552_BAY...",
            "email": "user1@outlook.com"
        },
        {
            "client_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            "refresh_token": "M.C552_BAY...",
            "email": "user2@outlook.com"
        }
    ]
}
```

| Поле | Описание |
|------|----------|
| `output_dir` | Папка для сохранения `.txt` файлов. `null` — вывод в консоль |
| `accounts` | Массив аккаунтов для обработки |
| `client_id` | ID приложения из Azure App Registration |
| `refresh_token` | OAuth2 refresh token |
| `email` | Email аккаунта (если не указан — скрипт попробует определить автоматически) |

## Использование

### Из конфига (все аккаунты)

```bash
python3 imap_outlook.py
```

### С сохранением в папку

```bash
python3 imap_outlook.py --output-dir ./emails
```

Для каждого аккаунта создается файл `{email}.txt` в указанной папке.

### Один аккаунт через аргументы

```bash
python3 imap_outlook.py --client-id YOUR_ID --refresh-token YOUR_TOKEN --email user@outlook.com
```

## Формат выгрузки

```
Почтовый ящик: user@outlook.com
Дата выгрузки:  2026-02-18 23:36:00
Всего писем:    4

======================================================================
  Письмо 1 из 4
======================================================================
  Тема:       Тема письма
  От:         sender@example.com
  Кому:       user@outlook.com
  Дата:       Thu, 12 Feb 2026 15:47:40 -0800
  Вложения:   document.pdf (1.2 MB)
----------------------------------------------------------------------
Текст письма...
```

## Структура проекта

```
.
├── imap_outlook.py       # Основной скрипт
├── config.json           # Конфигурация (не в git)
├── config.example.json   # Пример конфигурации
├── requirements.txt      # Зависимости
├── emails/               # Выгруженные письма (не в git)
│   └── user@outlook.com.txt
└── README.md
```
