# LengthHelper

Планировщик производства многожильного кабеля.
Автоматически рассчитывает: изолирование жил → скрутка → намотка на выходные барабаны.

---

## Структура проекта

```
LengthHelper/
├── new/                        ← АКТУАЛЬНАЯ ВЕРСИЯ (v5)
│   ├── models.py               Модели данных
│   ├── parser.py               Парсер (форматы v5 и старый)
│   ├── planner.py              Алгоритм планирования (OR-Tools CP-SAT)
│   ├── exporter.py             Экспорт результата в Excel
│   ├── main.py                 Точка входа (CLI)
│   └── ANALYSIS.md             Архитектурный анализ и план развития
│
├── test_temp_files/            ← Исходники v4 (для справки)
│   ├── claude_insulation_v4.ipynb   Ноутбук-прототип с OR-Tools
│   ├── input_v5.xlsx                Шаблон входного файла v5
│   ├── make_input_v5.py             Генератор шаблона v5
│   └── result_old/                  Архив сгенерированных инструкций
│
├── archive/                    ← Старые версии (хранятся для справки)
│   ├── models.py / parser.py / planner.py / exporter.py   Трек A (greedy)
│   ├── length_helper.ipynb     Первый монолитный ноутбук
│   ├── DS_insulation.ipynb     Исследовательский ноутбук
│   └── DS_ortools.ipynb        OR-Tools эксперименты
│
├── input.xlsx                  Входной файл (старый формат, для справки)
├── ТЗ.md                       Техническое задание
└── Замечания.docx              Замечания и правки
```

---

## Быстрый старт (v5)

```bash
pip install ortools openpyxl
cd new/
python main.py ../test_temp_files/input_v5.xlsx
```

Результат сохраняется в `result_<timestamp>.xlsx`.

### Программный вызов

```python
from new.parser import parse_input_v5
from new.planner import plan
from new.exporter import export

orders, raw_wires, insulated_cores, cable_stock, \
    core_drum_caps, cable_drum_caps, params = parse_input_v5('input_v5.xlsx')

result = plan(orders, raw_wires, insulated_cores, cable_stock,
              core_drum_caps, cable_drum_caps, params)

export('result.xlsx', result)
```

---

## Ключевое улучшение v5 vs v4

| | v4 (test_temp_files) | v5 (new/) |
|--|--|--|
| Алгоритм распределения ТПЖ | OR-Tools CP-SAT | OR-Tools CP-SAT |
| Полный конвейер (5 шагов) | Нет | **Да** |
| Склад изолированных жил | Нет | **Да** |
| Склад готового кабеля | Нет | **Да** |
| Выходные барабаны | Нет | **Да** |
| Чтение из Excel | Нет (ручной ввод) | **Да** |
| Гибкие кабели (flexible) | Да | **Да** |
| Группировка по составу жил | Да | **Да** |

---

## Зависимости

```
python >= 3.10
ortools >= 9.x    (pip install ortools)
openpyxl >= 3.1   (pip install openpyxl)
```
