"""
Тест: скрутка только из склада изолированных жил, несколько катушек на цвет.
Данные из примера пользователя.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from models import (
    InsulatedCore, CableOrder, CoreDrumCapacity, CableDrumCapacity,
    ProcessParams, DrumType,
)
from planner import plan
from exporter import export

# ─── Параметры кабеля ──────────────────────────────────────────────────
CABLE_MARK = 'ВВГ 5х2,5ок'
CS         = '2,5'
WT         = 'ок'
INS_MAT    = 'LS'
COLORS     = ['Натуральная', 'Синяя', 'Желто-зеленая', 'Коричневая', 'Черная']


def _core(id_, color, length):
    return InsulatedCore(
        id=id_,
        name=f'{color} {CS}{WT} {INS_MAT}',
        color=color,
        cross_section=CS,
        wire_type=WT,
        insulation_material=INS_MAT,
        fire_resistant='',
        length=length,
    )


# ─── Склад изолированных жил ───────────────────────────────────────────
insulated_cores = [
    # Натуральная: 3245 + 371 + 439 = 4055 м
    _core('НАТ-1', 'Натуральная',  3245),
    _core('НАТ-2', 'Натуральная',   371),
    _core('НАТ-3', 'Натуральная',   439),
    # Синяя: 2839 + 829 + 439 = 4107 м
    _core('СИН-1', 'Синяя',        2839),
    _core('СИН-2', 'Синяя',         829),
    _core('СИН-3', 'Синяя',         439),
    # Желто-зеленая: 3658 + 381 = 4039 м
    _core('ЖЗ-1',  'Желто-зеленая', 3658),
    _core('ЖЗ-2',  'Желто-зеленая',  381),
    # Коричневая: 3658 + 439 = 4097 м
    _core('КОР-1', 'Коричневая',   3658),
    _core('КОР-2', 'Коричневая',    439),
    # Черная: 670 + 711 + 1087 + 1328 = 3796 м
    _core('ЧЕР-1', 'Черная',        670),
    _core('ЧЕР-2', 'Черная',        711),
    _core('ЧЕР-3', 'Черная',       1087),
    _core('ЧЕР-4', 'Черная',       1328),
]

# ─── Кабельный журнал ──────────────────────────────────────────────────
# Отрезки выбраны так, что каждый ≤ любой одиночной катушки Черной
# (чтобы проверить, правильно ли плановик ищет одну катушку на партию).
JOURNAL = [670, 711, 1087, 1328]   # итого 3796 м = весь запас Черной

order = CableOrder(
    mark=CABLE_MARK,
    total_length=sum(JOURNAL),
    journal=JOURNAL,
    colors=COLORS,
    cross_section=CS,
    wire_type=WT,
    insulation_material=INS_MAT,
    fire_resistant='',
)

# ─── Ёмкости барабанов ────────────────────────────────────────────────
core_drum_caps = [
    CoreDrumCapacity(
        wire_key=f'{CS}{WT}',
        drum_types=[
            DrumType('К-500',   500),
            DrumType('К-1000', 1000),
            DrumType('Б-630',  2000),
            DrumType('Б-800',  3000),
            DrumType('Б-1000', 4500),
        ],
    ),
]

cable_drum_caps = [
    CableDrumCapacity(
        cable_mark=CABLE_MARK,
        drum_types=[
            DrumType('№12',   1500),
            DrumType('Б-630', 3000),
            DrumType('Б-800', 5000),
        ],
    ),
]

params = ProcessParams(
    max_insulation_run=4500.0,
    max_twisting_run=2100.0,
    allow_multi_segment_drum=True,
)

# ─── Планирование ─────────────────────────────────────────────────────
result = plan(
    orders=[order],
    raw_wires=[],
    insulated_cores=insulated_cores,
    cable_stock=[],
    core_drum_caps=core_drum_caps,
    cable_drum_caps=cable_drum_caps,
    params=params,
)

# ─── Вывод результатов ────────────────────────────────────────────────
print('=' * 60)
print('REZULTAT PLANIROVANIYA')
print('=' * 60)

print(f'\nПартий скрутки: {len(result.batches)}')
for b in result.batches:
    segs = ' + '.join(str(int(s)) for s in b.segments)
    print(f'  {b.id}  [{segs}] = {int(b.total_length)} м')

print(f'\nИсп. жил со склада: {len(result.insulated_core_uses)}')
for u in result.insulated_core_uses:
    print(f'  {u.for_batch_id} / {u.color:15s}: берём {int(u.length):5d} м '
          f'из {u.source_id} (ост. {int(u.remainder)} м)')

def _safe(s): return s.encode('cp1251', errors='replace').decode('cp1251')

print(f'\nОшибок: {len(result.errors)}')
for e in result.errors:
    print(f'  {_safe(e)}')

print(f'\nПредупреждений: {len(result.warnings)}')
for w in result.warnings:
    print(f'  {_safe(w)}')

print()
export('output_test.xlsx', result)
