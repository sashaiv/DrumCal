"""
Точка входа для LengthHelper v5.

Использование:
  python main.py input_v5.xlsx             # v5 формат (4 листа)
  python main.py input.xlsx --old          # старый формат (5 листов)
  python main.py input_v5.xlsx --output result.xlsx

Пример из Python:
    from parser import parse_input_v5
    from planner import plan
    from exporter import export

    orders, raw_wires, insulated_cores, cable_stock, \\
        core_drum_caps, cable_drum_caps, params = parse_input_v5('input_v5.xlsx')

    result = plan(orders, raw_wires, insulated_cores, cable_stock,
                  core_drum_caps, cable_drum_caps, params)

    if result.errors:
        print('ОШИБКИ:')
        for e in result.errors:
            print(' ', e)

    if result.warnings:
        print('ПРЕДУПРЕЖДЕНИЯ:')
        for w in result.warnings:
            print(' ', w)

    export('result.xlsx', result)
"""
import sys
import os
import argparse
from datetime import datetime


def main():
    parser_arg = argparse.ArgumentParser(
        description='LengthHelper v5 — планировщик производства кабеля'
    )
    parser_arg.add_argument('input', help='Путь к входному Excel-файлу')
    parser_arg.add_argument('--old', action='store_true',
                            help='Использовать старый формат (5 листов)')
    parser_arg.add_argument('--output', default=None,
                            help='Путь к выходному Excel-файлу (по умолчанию auto)')
    args = parser_arg.parse_args()

    # Добавляем текущую папку в sys.path чтобы найти модули
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

    from parser import parse_input_v5, parse_input
    from planner import plan
    from exporter import export

    # Чтение входного файла
    print(f'\n{"="*60}')
    print(f'  LengthHelper v5 — OR-Tools CP-SAT оптимизатор')
    print(f'{"="*60}')
    print(f'  Входной файл: {args.input}')
    print(f'  Формат: {"старый (5 листов)" if args.old else "v5 (4 листа)"}')
    print()

    try:
        if args.old:
            result_tuple = parse_input(args.input)
        else:
            result_tuple = parse_input_v5(args.input)
    except KeyError as e:
        print(f'❌ Ошибка чтения файла: {e}')
        sys.exit(1)
    except Exception as e:
        print(f'❌ Ошибка при парсинге: {e}')
        raise

    orders, raw_wires, insulated_cores, cable_stock, \
        core_drum_caps, cable_drum_caps, params = result_tuple

    print(f'  Заказов: {len(orders)}')
    print(f'  Барабанов ТПЖ: {len(raw_wires)}')
    print(f'  Склад жил: {len(insulated_cores)}')
    print(f'  Склад кабеля: {len(cable_stock)}')
    print()

    # Планирование
    print('  Запуск планировщика...')
    result = plan(orders, raw_wires, insulated_cores, cable_stock,
                  core_drum_caps, cable_drum_caps, params)

    # Вывод результата
    print(f'\n{"="*60}')
    print(f'  РЕЗУЛЬТАТ')
    print(f'{"="*60}')
    print(f'  Партий скрутки:        {len(result.batches)}')
    print(f'  Прогонов изолирования: {len(result.insulation_runs)}')
    print(f'  Жил со склада:         {len(result.insulated_core_uses)}')
    print(f'  Выходных барабанов:    {len(result.drum_assignments)}')

    if result.errors:
        print(f'\n  ❌ ОШИБКИ ({len(result.errors)}):')
        for e in result.errors:
            print(f'    {e}')

    if result.warnings:
        print(f'\n  ⚠ ПРЕДУПРЕЖДЕНИЯ ({len(result.warnings)}):')
        for w in result.warnings:
            print(f'    {w}')

    # Экспорт
    if args.output:
        out_path = args.output
    else:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_path = f'result_{ts}.xlsx'

    try:
        export(out_path, result)
        print(f'\n✓ Результат сохранён: {os.path.abspath(out_path)}')
    except PermissionError:
        print(f'\n⚠ Файл {out_path} занят — закройте Excel и повторите.')
        sys.exit(1)


if __name__ == '__main__':
    main()
