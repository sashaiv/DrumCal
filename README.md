# LengthHelper

Планировщик производства кабеля: изолирование → скрутка → намотка на барабаны.

## Файлы

| Файл | Назначение |
|---|---|
| `input.xlsx` | Входные данные (заказы, склад, барабаны, параметры) |
| `output.xlsx` | Результат планирования |
| `length_helper.ipynb` | Основной ноутбук — запускать здесь |
| `create_input_template.py` | Генерация шаблона `input.xlsx` |
| `models.py` | Модели данных |
| `parser.py` | Парсер входного Excel |
| `planner.py` | Алгоритм планирования |
| `exporter.py` | Экспорт результата в Excel |
| `test_multicolor.py` | Тест многожильного сценария |
| `ТЗ.md` | Техническое задание |

## Быстрый старт

1. Заполни `input.xlsx` (или запусти `create_input_template.py` для шаблона).
2. Открой `length_helper.ipynb` в Jupyter / VS Code.
3. Kernel → Restart & Run All.
4. Результат — в `output.xlsx`.

