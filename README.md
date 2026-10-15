# UniversalReport — макрос для Excel (АМД / Часы / УК)

![Version](https://img.shields.io/badge/version-v1.0-green) ![License](https://img.shields.io/badge/license-MIT-blue) ![Excel](https://img.shields.io/badge/Excel-2016%2B-brightgreen)

Один макрос для трех типов отчетов. Нажал `Alt+F8` → `UniversalReport` → готово.

> **v1.0 — первый стабильный релиз (15.10.2026)**

**Что делает:**
* **АМД** — сводная по `отчет`. `Результат вызова` / `ПРОЕКТ ТМ`, `% АО` и `% от общего`
* **Часы** — мульти по ТЛ, чистое время
* **УК** — мульти по ТЛ, только `ФИО` + `Итого уникальных`

---

## Скачать

Скачай актуальный релиз

---

## Видео-гайды (7 шт)

### Шаг 0 — Включить Разработчика

Файл → Параметры → Настройка ленты → Разработчик

https://github.com/user-attachments/assets/09c3cf6b-6532-4ffd-9230-40c7c936a967

### Шаг 1 — Проверить наличие личной книги

https://github.com/user-attachments/assets/30f2213d-0b4b-4940-b02f-daaedf4d2a23

### Шаг 2 — Установка (3 варианта)

**Вариант 1 — Через PERSONAL.XLSB (основной)**

Win+R → `%APPDATA%\Microsoft\Excel\XLSTART` → скопировать `PERSONAL.XLSB`

https://github.com/user-attachments/assets/e3c40bbd-804e-4b4d-9689-5675e8a5c188

**Вариант 2 — Импорт .bas**

File → Import File → `UniversalReport.bas`

https://github.com/user-attachments/assets/197c48a5-b594-4790-b3e0-85d24cea3056

**Вариант 3 — Копипаст**

Insert → Module → Ctrl+V

https://github.com/user-attachments/assets/5efd2c88-a7c9-4880-8c55-9d96d9ea06ed

### Шаг 3 — Фильтр по ТЛ (если нужно не по всем)

Вид → Показать → PERSONAL.XLSB → переименовать лист в TL → вписать ТЛ → Вид → Скрыть

https://github.com/user-attachments/assets/e79c2b39-c256-4ba6-854c-3622bcf9df9a

### Шаг 4 — Запуск

вид → макрос (или Alt+F8) → UniversalReport → кнопка скопировать изображение

https://github.com/user-attachments/assets/09bfa352-4c35-482d-bb8c-3225818b91d4

---

## Лицензия

MIT
