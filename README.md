# Excel GTIN/EAN Finder

> Public experimental source: see [PUBLICATION.md](PUBLICATION.md) for actual checks, known defects, restored modules and operational limits. Earlier audit decisions are historical.

Desktop-утилита на Python для анализа Excel-файлов со штрихкодами и номенклатурой. Скрипт ищет позиции, у которых найден только один уникальный штрихкод, и сохраняет результат в новый Excel-файл с отдельным проблемным листом.

Проект рассчитан на локальную офисную обработку таблиц без внешних сервисов и API.

## Возможности

- выбор Excel-файла через графический диалог;
- автоматический поиск колонок по распространённым названиям;
- безопасная обработка штрихкодов как строк, чтобы не терять ведущие нули;
- подсчёт числа уникальных штрихкодов для каждой номенклатуры;
- экспорт результата в новый Excel с двумя листами.

## Стек

- Python;
- Tkinter;
- pandas;
- openpyxl.

## Как использовать

```bash
python analyze_barcodes.py
```

После выбора Excel-файла результат сохраняется рядом с исходником с суффиксом `_analyzed`.

## Статус аудита

Небольшой public utility candidate после full-history scan и добавления sample workbook без реальных товарных данных. Имя основного файла уже исправлено на `analyze_barcodes.py`.

## License recommendation

MIT.


## Verified XLSX behavior — 2026-10-29

Install: `python -m pip install -r requirements.txt`, then run `python analyze_barcodes.py`.
The entrypoint filename has been corrected from the old double `.py.py` suffix.

Barcodes stored as Excel **text** retain leading zeros. Distinct text codes remain distinct when
counting per product; additional numeric columns keep numeric values. Zeros shown only through
the number format of an already numeric cell cannot be reconstructed reliably and are not guessed.

`python -m unittest discover -s tests -v` runs a synthetic XLSX roundtrip and checks exact barcode
text, unique counts, the problem sheet, numeric prices and unchanged input bytes. Only XLSX was
verified; the older XLS path and GUI interaction were not tested.
