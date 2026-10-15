# UniversalReport — макрос для Excel (АМД / Часы / УК)

![Version](https://img.shields.io/badge/version-v1.0-green) ![License](https://img.shields.io/badge/license-MIT-blue) ![Excel](https://img.shields.io/badge/Excel-2016%2B-brightgreen)

Один макрос для трех типов отчетов. Нажал `Alt+F8` → `UniversalReport` → готово.

> **v1.0 — первый стабильный релиз (15.10.2026)**

**Что делает:**
* **АМД** — сводная по `отчет`. `Результат вызова` / `ПРОЕКТ ТМ`, `% АО` и `% от общего`
* **Часы** — мульти по ТЛ, чистое время
* **УК** — мульти по ТЛ, только `ФИО` + `Итого уникальных`

---
# Как пользоваться
## Скачать актуальный релиз

https://github.com/user-attachments/assets/b4b6f345-4eb0-4ab3-b290-8bf135ba6971

---

### Шаг 0 — Включить Разработчика

Файл → Параметры → Настройка ленты → Разработчик

https://github.com/user-attachments/assets/a0ed15f4-6957-4500-b9a0-caa73ca0ab11

### Шаг 1 — Установка (3 варианта)

**Вариант 1 — Через копирование личной книги PERSONAL.XLSB (основной)**

Win+R → `%APPDATA%\Microsoft\Excel\XLSTART` → скопировать `PERSONAL.XLSB`

https://github.com/user-attachments/assets/61f724c4-448d-4c47-875d-155cfaa546bb

## для вариантов 2 и 3 нужно проверить наличие личной книги PERSONAL.XLSB в Visual Basic
если в Visual Basic нет такой книги то нужно начать запись макроса обязательно указав что сохранять его нужно в личную книгу макросов, остановить запись, после этого книга появится в Visual Basic.

https://github.com/user-attachments/assets/2df7900a-0903-4a0c-95f2-6f3123860f6d

**Вариант 2 — Импорт .bas**

File → Import File → `UniversalReport.bas`

https://github.com/user-attachments/assets/ecf0fe77-ec33-46ac-b7d3-f18ea8c96f68

**Вариант 3 — Копировать код из .bas**

Insert → Module → Ctrl+V

https://github.com/user-attachments/assets/d8a391c6-409a-456c-b822-2719739f542f

### Шаг 2 — Фильтр по ТЛ (если нужно не по всем)

Вид → Показать → PERSONAL.XLSB → переименовать лист в TL → вписать ТЛ → Вид → Скрыть

https://github.com/user-attachments/assets/770dabfa-b3f2-42a0-8ebb-5ed65a76723c

### Шаг 3 — Запуск

вид → макрос (или Alt+F8) → UniversalReport → кнопка скопировать изображение

https://github.com/user-attachments/assets/586456bf-e384-4374-bed7-f02f41d2e7ea

---

## Лицензия

MIT
