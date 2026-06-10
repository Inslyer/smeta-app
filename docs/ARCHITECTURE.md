# Архитектура приложения «Генрих»

## Что это
Streamlit-приложение для электромонтажной компании ГРАФТИО. Сметчик/закупщик подгружает три источника данных, AI-ассистент «Генрих» сопоставляет позиции, на выходе — Excel со сводной сметой и расчётом маржи.

## Стек
- **UI**: Streamlit 1.40+, Pandas, кастомный CSS (тёмная тема, Inter)
- **AI**: Anthropic API (Claude Sonnet 4.x), tool_use для structured output, нативный PDF input для счетов
- **БД**: SQLAlchemy 2.0 Core. Postgres (Neon) на проде, SQLite локально
- **Excel**: openpyxl (генерация + копирование исходников)
- **Парсинг входных Excel**: openpyxl, pdfplumber для PDF
- **Деплой**: GitHub → Streamlit Community Cloud, авто-rebuild на push в main

## Структура папок
```
smeta-app/
  .streamlit/config.toml      # тёмная тема, palette, шрифты
  assets/
    graftio_symbol.svg        # «Сатурн» — только знак (page icon, sidebar)
    graftio_planet.svg        # полный горизонтальный логотип ГРАФТИО
  core/
    models.py                 # dataclass: EstimateItem, Estimate, Match,
                              #   MaterialMatch, ContractorExtra,
                              #   SupplierInvoice, SupplierInvoiceItem
    db.py                     # SQLAlchemy engine, схема, CRUD
    parser_client.py          # Excel сметы Заказчика → Estimate
    parser_contractor.py      # Excel сметы подрядчика → Estimate
    parser_supplier_invoice.py # PDF/XLSX счёта поставщика → SupplierInvoice
    matcher.py                # три Claude-функции + дисковый кэш
    exporter.py               # build_workbook → Excel
  app_with_ai.py              # главный Streamlit-файл (всё UI)
  app.py, app_min.py          # legacy / минимальный демо — не использовать
  requirements.txt
  start.sh, start.bat         # локальный запуск через venv
  .env.example                # ключи и DATABASE_URL
```

## Потоки данных

### Поток 1: загрузка и парсинг
```
UI uploader → _save_upload(tmp) → parser_*.py → Estimate / SupplierInvoice
                                      │
                             session_state["client_estimate"] / etc
```
Bytes исходных XLSX сохраняются в session_state, чтобы потом вставить как-есть в финальный Excel.

### Поток 2: AI-сопоставление (трижды)
```
1. match_estimates(client, contractor)
   → list[Match]: какая работа клиента ↔ какая работа подрядчика

2. match_materials(client, all_invoice_items_of_project)
   → list[MaterialMatch]: материал клиента ↔ позиция счёта поставщика
   (приоритет — по артикулу, иначе семантика)

3. classify_contractor_extras(client, contractor, matches_from_step_1)
   → list[ContractorExtra]: каждая неприматченная позиция подрядчика —
   либо "included" (компонент крупной позиции Заказчика), либо "extra"
   (вне сметы)
```
Каждый вызов через `tool_use` со схемой инструмента. Дисковый кэш `core/.cache/matches/<sha>.json` чтобы не платить за повторы.

### Поток 3: сборка Excel
`build_workbook(client, contractor, matches, material_matches, supplier_rows, client_xlsx_bytes, contractor_xlsx_bytes, invoices_with_items, contractor_extras)` собирает книгу:

1. **Сводная смета** (`_build_summary_sheet`) — блок Работы → блок Материалы → блок Доп. работ вне сметы (если есть) → ВСЕГО ПО ПРОЕКТУ. Все суммы — формулы.
2. **Аналитика** (`_build_analytics_sheet`) — метрики проекта, маржа по разделам работ, маржа по категориям материалов. Все ссылки на «Сводную смету» формулами.
3. **Исходник Заказчика: <лист>** — копии листов из загруженного XLSX как-есть (`_copy_external_sheets`).
4. **Исходник подрядчика: <лист>** — то же.
5. **Счёт <Поставщик> №<номер>** — по одной вкладке на каждый счёт проекта.

## БД (SQLAlchemy Core, не ORM)

| таблица | назначение |
|---------|------------|
| `suppliers` | Поставщики (kind=material/work, price_policy=fixed/volatile) |
| `projects` | Проекты-объекты, к которым привязываются счета |
| `invoices` | Шапки счетов поставщиков |
| `invoice_items` | Позиции счетов с артикулами |
| `items` + `prices` | Старая модель «база цен» — пока живёт, но не используется напрямую через UI |

Подключение определяется в `_resolve_database_url()`:
1. `DATABASE_URL` env (`postgres://...` автоматически переписывается в `postgresql+psycopg://...`)
2. иначе SQLite в `data/prices.db`

## Бренд
Палитра, шрифт, лого: см. `docs/BRAND.md`.

## Конвенции имён
В коде и моделях — `client` / `contractor` / `supplier`. В UI и в Excel — **Заказчик** / **Подрядчик** / **Поставщик**. Инструмент — **Генрих** (никогда «Claude» в UI-строках).
