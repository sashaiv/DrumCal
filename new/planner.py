"""
Алгоритм производственного планирования (v5 — гибрид Трека A + OR-Tools из Трека B).

Шаги:
  1. Заполнить кабельный журнал (если не задан).
  2. Покрыть отрезки журнала из готового кабеля на складе.
  3. Упаковать оставшиеся отрезки в партии скрутки (FFD bin-packing),
     РАЗДЕЛЬНО для каждой группы кабелей по составу жил (из claude_insulation_v4).
  4. Для каждой партии × цвет — выделить жилу:
     4a. Склад изолированных жил (multi-spool, из root/planner.py).
     4b. ТПЖ через CP-SAT оптимизатор (из claude_insulation_v4) — оптимальное назначение.
  5. Упаковать ВСЕ отрезки журнала на выходные барабаны (greedy bin-packing, из root).

Ключевое улучшение: Step 4b заменяет жадный выбор max-drum из root/planner.py на
CP-SAT Assignment Problem — гарантирует оптимальное использование барабанов ТПЖ.
"""
from __future__ import annotations
import math
from copy import deepcopy
from typing import List, Tuple, Dict, Optional
from collections import defaultdict

from models import (
    RawWire, InsulatedCore, CableStock, CableOrder,
    CoreDrumCapacity, CableDrumCapacity, ProcessParams,
    TwistingBatch, InsulationRun, InsulatedCoreUse,
    CableStockUse, DrumAssignment, PlanResult, DrumType,
)

# ── CP-SAT опционально (если ortools не установлен — fallback на жадный) ────
try:
    from ortools.sat.python import cp_model as _cp_model
    _ORTOOLS_AVAILABLE = True
except ImportError:
    _ORTOOLS_AVAILABLE = False
    _cp_model = None

_RUN_CTR: Dict[str, int] = {}


def _uid(prefix: str) -> str:
    _RUN_CTR[prefix] = _RUN_CTR.get(prefix, 0) + 1
    return f'{prefix}-{_RUN_CTR[prefix]:03d}'


# ═══════════════════════════════════════════════════════════════════════
# Шаг 1 — кабельный журнал
# ═══════════════════════════════════════════════════════════════════════

def _fill_journal(order: CableOrder, params: ProcessParams) -> List[float]:
    """
    Если журнал задан — возвращает его.
    Иначе делит total_length на равные отрезки длиной ≤ max_twisting_run.
    """
    if order.has_journal:
        return list(order.journal)

    total = order.total_length
    n = max(1, math.ceil(total / params.max_twisting_run))
    base = total / n
    segments = [round(base) for _ in range(n - 1)]
    last = round(total - sum(segments))
    segments.append(last)
    return segments


# ═══════════════════════════════════════════════════════════════════════
# Шаг 2 — покрытие из склада готового кабеля
# ═══════════════════════════════════════════════════════════════════════

def _allocate_cable_stock(
    order: CableOrder,
    segments: List[float],
    cable_stock: List[CableStock],
) -> Tuple[List[Optional[CableStockUse]], List[float]]:
    """
    Для каждого отрезка журнала пробует найти подходящий кабель на складе.
    Возвращает (stock_uses, remaining).
    """
    stock_uses: List[Optional[CableStockUse]] = []
    remaining: List[float] = []

    for seg in segments:
        candidates = [
            s for s in cable_stock
            if s.cable_mark == order.mark and s.available >= seg
        ]
        if candidates:
            best = min(candidates, key=lambda s: s.available)   # наименьший подходящий
            remainder = round(best.available - seg, 6)
            best.used += seg
            stock_uses.append(CableStockUse(
                id=_uid('СК'),
                cable_mark=order.mark,
                source_id=best.id,
                segment_length=seg,
                remainder=remainder,
            ))
        else:
            stock_uses.append(None)
            remaining.append(seg)

    return stock_uses, remaining


# ═══════════════════════════════════════════════════════════════════════
# Шаг 3 — упаковка отрезков в партии скрутки (FFD) с группировкой
# ═══════════════════════════════════════════════════════════════════════

def _colors_group_label(colors: List[str], existing_labels: Dict[frozenset, str]) -> str:
    """
    Формирует метку группы по набору цветов.
    Логика из claude_insulation_v4: '5ж', '4ж', '3ж', '2ж'.
    При коллизии числа жил с разным составом → суффиксы 'А', 'Б', 'В', …
    """
    key = frozenset(colors)
    if key in existing_labels:
        return existing_labels[key]

    n = len(colors)
    base = f'{n}ж'
    # Проверяем коллизию: то же число жил, другой состав
    used_labels = set(existing_labels.values())
    if base not in used_labels:
        existing_labels[key] = base
        return base

    # Ищем первый свободный суффикс А, Б, В, …
    for suffix in 'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩ':
        label = f'{base}-{suffix}'
        if label not in used_labels:
            existing_labels[key] = label
            return label

    existing_labels[key] = f'{base}-?'
    return existing_labels[key]


def _max_batch_size(
    wire_key: str,
    core_drum_caps: List[CoreDrumCapacity],
    params: ProcessParams,
) -> float:
    cap_obj = next((c for c in core_drum_caps if c.wire_key == wire_key), None)
    max_drum = cap_obj.max_capacity if cap_obj else params.max_insulation_run
    return min(params.max_twisting_run, max_drum)


def _pack_batches(
    order: CableOrder,
    segments: List[float],
    core_drum_caps: List[CoreDrumCapacity],
    params: ProcessParams,
    group_label: str,
) -> Tuple[List[TwistingBatch], List[str]]:
    """
    Bin-packing отрезков в партии скрутки.
    keep_journal_order=False (по умолчанию): First Fit Decreasing.
    keep_journal_order=True (П5): в порядке журнала без сортировки.
    """
    errors: List[str] = []
    batches: List[TwistingBatch] = []
    max_batch = _max_batch_size(order.wire_key, core_drum_caps, params)

    ordered_segs = segments if params.keep_journal_order else sorted(segments, reverse=True)
    for seg in ordered_segs:
        if seg > max_batch:
            errors.append(
                f'❌ {order.mark}: отрезок {seg:.0f} м превышает макс. партию скрутки '
                f'{max_batch:.0f} м — невозможно произвести без спайки.'
            )
            continue

        placed = False
        for batch in batches:
            if batch.total_length + seg <= max_batch + 1e-6:
                batch.segments.append(seg)
                placed = True
                break

        if not placed:
            batches.append(TwistingBatch(
                id=_uid('ПС'),
                cable_mark=order.mark,
                segments=[seg],
                wire_key=order.wire_key,
                colors=order.colors,
                insulation_material=order.insulation_material,
                fire_resistant=order.fire_resistant,
                group_label=group_label,
            ))

    return batches, errors


# ═══════════════════════════════════════════════════════════════════════
# Шаг 4a — выделение жил со склада изолированных (multi-spool)
# (из root/planner.py, без изменений)
# ═══════════════════════════════════════════════════════════════════════

def _commit_spool(
    batch: TwistingBatch,
    color: str,
    wire_key: str,
    spool,
    seg_indices: list,
    total: float,
    spool_num: Dict[str, int],
    waste_thresh: float,
    warnings: List[str],
    all_ins_uses: List[InsulatedCoreUse],
):
    rem = round(spool.available - total, 6)
    spool.used += total
    spool_num[color] += 1
    if rem < waste_thresh:
        warnings.append(
            f'⚠ {batch.cable_mark} / {color}: катушка «{spool.name}» '
            f'после использования — остаток {rem:.0f} м '
            f'(< {waste_thresh:.0f} м), вероятно в отход.'
        )
    all_ins_uses.append(InsulatedCoreUse(
        id=_uid('ГЖ'), color=color, wire_key=wire_key,
        source_id=spool.id, source_name=spool.name,
        length=round(total, 6), remainder=rem,
        for_batch_id=batch.id,
        covered_segments=list(seg_indices),
        spool_index=spool_num[color],
    ))


def _allocate_batch_all_colors(
    batch: TwistingBatch,
    colors: List[str],
    insulated_cores: List[InsulatedCore],
    params: ProcessParams,
) -> Tuple[List[float], List[InsulatedCoreUse], Dict[str, bool], List[str], List[str]]:
    """
    Распределяет СКЛАДСКИЕ жилы для партии по всем цветам одновременно.

    Возвращает:
        actual_segs     — фактические длины сегментов (могут быть уменьшены)
        ins_uses        — использования со склада
        needs_tpzh      — {color: True} для цветов, требующих изолирования из ТПЖ
        errors, warnings
    """
    min_len  = params.min_construction_length
    ins_mat  = batch.insulation_material
    fr_flag  = batch.fire_resistant
    wire_key = batch.wire_key

    # Классификация цветов: есть ли склад?
    stock_cands: Dict[str, List[InsulatedCore]] = {}
    from_stock: Dict[str, bool] = {}
    for color in colors:
        cands = [
            c for c in insulated_cores
            if c.color == color
            and c.wire_key == wire_key
            and c.insulation_material == ins_mat
            and c.fire_resistant == fr_flag
        ]
        stock_cands[color] = sorted(cands, key=lambda c: c.available)
        from_stock[color] = sum(c.available for c in cands) >= min_len - 1e-6

    needs_tpzh = {color: not from_stock[color] for color in colors}

    # Если все цвета — только ТПЖ, фактические длины = плановые
    if all(needs_tpzh.values()):
        return list(batch.segments), [], needs_tpzh, [], []

    # Состояние жадного алгоритма на катушку
    curr: Dict[str, tuple] = {c: (None, 0.0) for c in colors if from_stock[c]}
    pending: Dict[tuple, dict] = {}
    spool_num: Dict[str, int] = {c: 0 for c in colors}

    actual_segs: List[float] = []
    all_ins_uses: List[InsulatedCoreUse] = []
    warnings: List[str] = []
    errors: List[str] = []
    waste_thresh = params.waste_warning_threshold_m

    for planned_seg in batch.segments:
        color_max: Dict[str, float] = {}
        color_spool: Dict[str, tuple] = {}

        for color in colors:
            if not from_stock[color]:
                continue
            cs, cb = curr[color]

            if cs is not None and cb >= planned_seg - 1e-6:
                color_max[color] = planned_seg
                color_spool[color] = (cs, False)
            else:
                fitting = sorted(
                    [c for c in stock_cands[color]
                     if c is not cs and c.available >= planned_seg - 1e-6],
                    key=lambda c: c.available,
                )
                if fitting:
                    color_max[color] = planned_seg
                    color_spool[color] = (fitting[0], True)
                else:
                    options = [(c.available, c) for c in stock_cands[color]
                               if c is not cs and c.available >= min_len - 1e-6]
                    if cs is not None and cb >= min_len - 1e-6:
                        options.append((cb, cs))
                    if options:
                        best_avail, best_spool = max(options, key=lambda x: x[0])
                        color_max[color] = best_avail
                        color_spool[color] = (best_spool, best_spool is not cs)
                    else:
                        color_max[color] = 0.0
                        color_spool[color] = (None, False)

        if color_max:
            actual = min(color_max.values())
        else:
            actual = planned_seg

        if actual < min_len - 1e-6:
            warnings.append(
                f'⚠ {batch.cable_mark}: отрезок {planned_seg:.0f}м пропущен '
                f'(максимально возможно {actual:.0f}м < мин.{min_len:.0f}м).'
            )
            continue

        if actual < planned_seg - 0.5:
            warnings.append(
                f'⚠ {batch.cable_mark}: отрезок {planned_seg:.0f}м → {actual:.0f}м '
                f'(ограничение склада жил).'
            )

        ai = len(actual_segs)
        actual_segs.append(round(actual, 3))

        for color in colors:
            if not from_stock[color]:
                continue
            cs, cb = curr[color]
            spool, is_new = color_spool[color]
            if spool is None:
                continue

            if is_new:
                if cs is not None:
                    k = (color, cs.id)
                    if k in pending:
                        pu = pending.pop(k)
                        _commit_spool(batch, color, wire_key, pu['spool'],
                                      pu['segs'], pu['total'], spool_num,
                                      waste_thresh, warnings, all_ins_uses)
                curr[color] = (spool, spool.available - actual)
                pending[(color, spool.id)] = {
                    'spool': spool, 'segs': [ai], 'total': actual}
            else:
                k = (color, spool.id)
                if k in pending:
                    pending[k]['segs'].append(ai)
                    pending[k]['total'] += actual
                else:
                    pending[k] = {'spool': spool, 'segs': [ai], 'total': actual}
                curr[color] = (spool, cb - actual)

    for (color, _), pu in list(pending.items()):
        _commit_spool(batch, color, wire_key, pu['spool'], pu['segs'],
                      pu['total'], spool_num, waste_thresh, warnings, all_ins_uses)

    return actual_segs, all_ins_uses, needs_tpzh, errors, warnings


# ═══════════════════════════════════════════════════════════════════════
# Шаг 4b — CP-SAT: оптимальное назначение ТПЖ-барабанов
# (ключевое улучшение относительно root/planner.py)
# ═══════════════════════════════════════════════════════════════════════

def _core_drum_for(
    length: float,
    wire_key: str,
    core_drum_caps: List[CoreDrumCapacity],
    params: ProcessParams,
) -> Tuple[str, float]:
    cap_obj = next((c for c in core_drum_caps if c.wire_key == wire_key), None)
    if cap_obj:
        drum = cap_obj.smallest_fitting(length)
        if drum:
            return drum.name, drum.capacity
    return '—', params.max_insulation_run


def _allocate_tpzh_cpsat(
    tasks: List[Tuple[str, str, float, float]],
    # (batch_id, color, run_length, raw_needed) — run_length = что получим;
    # raw_needed = run_length + startup + tolerance = сколько снять с барабана.
    drums: List[RawWire],
    params: ProcessParams,
) -> Tuple[Dict[Tuple[str, str], str], List[str]]:
    """
    CP-SAT Assignment Problem: каждой задаче (batch_id, color) назначить один барабан ТПЖ.

    Почему CP-SAT лучше жадного:
      Жадный берёт наибольший доступный барабан для каждой задачи по очереди.
      Это может оставить большую задачу без подходящего барабана, хотя оптимальное
      решение существует.

    Возвращает:
      assignment  — {(batch_id, color): drum_id}
      errors      — список ошибок

    Fallback: если ortools не установлен, использует жадный алгоритм.
    """
    errors: List[str] = []

    if not tasks:
        return {}, []

    # Проверяем доступность ortools
    if not _ORTOOLS_AVAILABLE:
        return _allocate_tpzh_greedy(tasks, drums, errors), errors

    model = _cp_model.CpModel()

    # Масштаб для целочисленной оптимизации: работаем в дм (×10)
    SCALE = 10
    tasks_list = list(tasks)

    # Переменные: assign[i][j] ∈ {0, 1}
    # assign[i][j] = 1 → задача i назначена на барабан j
    assign = {}
    for i, (bid, color, run_len, raw_needed) in enumerate(tasks_list):
        for j, drum in enumerate(drums):
            assign[(i, j)] = model.new_bool_var(f'a_{i}_{j}')

    # Ограничение 1: каждая задача назначена ровно на один барабан
    for i in range(len(tasks_list)):
        model.add(sum(assign[(i, j)] for j in range(len(drums))) == 1)

    # Ограничение 2: ёмкость каждого барабана не превышена
    for j, drum in enumerate(drums):
        avail_scaled = int(drum.available * SCALE)
        model.add(
            sum(
                assign[(i, j)] * int(tasks_list[i][3] * SCALE)  # raw_needed
                for i in range(len(tasks_list))
            ) <= avail_scaled
        )

    # Ограничение 3: задача может быть назначена только на барабан с достаточным запасом
    for i, (bid, color, run_len, raw_needed) in enumerate(tasks_list):
        for j, drum in enumerate(drums):
            if drum.available < raw_needed - 1e-6:
                model.add(assign[(i, j)] == 0)

    # Цель: минимизировать отходы (эквивалентно максимизации использования)
    # Отход барабана j = drum.available - Σ(assign[i,j] * raw_needed[i])
    # Минимизируем суммарный отход.
    total_waste_scaled = sum(
        int(drum.available * SCALE) -
        sum(assign[(i, j)] * int(tasks_list[i][3] * SCALE) for i in range(len(tasks_list)))
        for j, drum in enumerate(drums)
    )
    model.minimize(total_waste_scaled)

    # Решаем
    solver = _cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = params.cpsat_time_limit
    status = solver.solve(model)

    if status in (_cp_model.OPTIMAL, _cp_model.FEASIBLE):
        assignment: Dict[Tuple[str, str], str] = {}
        for i, (bid, color, run_len, raw_needed) in enumerate(tasks_list):
            for j, drum in enumerate(drums):
                if solver.value(assign[(i, j)]) == 1:
                    assignment[(bid, color)] = drum.id
                    break
        return assignment, errors
    else:
        # CP-SAT не нашёл решения — попробуем жадный для диагностики
        errors.append(
            f'⚠ CP-SAT не нашёл оптимального решения (статус: {solver.status_name(status)}). '
            f'Используется жадный алгоритм.'
        )
        return _allocate_tpzh_greedy(tasks, drums, errors), errors


def _allocate_tpzh_greedy(
    tasks: List[Tuple[str, str, float, float]],
    drums: List[RawWire],
    errors: List[str],
) -> Dict[Tuple[str, str], str]:
    """
    Fallback: жадный алгоритм — берёт наибольший доступный барабан.
    Используется когда ortools не установлен или CP-SAT не нашёл решения.

    Внимание: может дать неоптимальный результат (см. ANALYSIS.md Проблема 1).
    """
    # Работаем с копиями балансов
    balances = {d.id: d.available for d in drums}
    drum_by_id = {d.id: d for d in drums}
    assignment: Dict[Tuple[str, str], str] = {}

    for bid, color, run_len, raw_needed in tasks:
        candidates = [
            d for d in drums
            if balances[d.id] >= raw_needed - 1e-6
        ]
        if not candidates:
            errors.append(
                f'❌ {bid} / {color}: нет барабана ТПЖ с достаточным остатком '
                f'(нужно {raw_needed:.0f} м, ни один не подходит).'
            )
            continue
        best = max(candidates, key=lambda d: balances[d.id])
        assignment[(bid, color)] = best.id
        balances[best.id] -= raw_needed

    return assignment


def _build_insulation_runs(
    batch: TwistingBatch,
    actual_segs: List[float],
    needs_tpzh: Dict[str, bool],
    assignment: Dict[Tuple[str, str], str],
    raw_wires: List[RawWire],
    core_drum_caps: List[CoreDrumCapacity],
    params: ProcessParams,
) -> Tuple[List[InsulationRun], List[str]]:
    """
    Создаёт InsulationRun объекты по результатам CP-SAT назначения.
    Обновляет RawWire.used для назначенных барабанов.
    """
    runs: List[InsulationRun] = []
    errors: List[str] = []
    n_segs = len(actual_segs)
    total_actual = round(sum(actual_segs), 6)

    if total_actual <= 0:
        return runs, errors

    tol_total = n_segs * params.length_tolerance_m
    startup = params.insulation_startup_loss_m
    raw_needed = total_actual + tol_total + startup

    for color in batch.colors:
        if not needs_tpzh.get(color, False):
            continue

        drum_id = assignment.get((batch.id, color))
        if drum_id is None:
            # Ошибка уже добавлена в _allocate_tpzh_cpsat
            continue

        rw = next((r for r in raw_wires if r.id == drum_id), None)
        if rw is None:
            errors.append(f'❌ {batch.cable_mark}/{color}: барабан {drum_id} не найден.')
            continue

        drum_name, drum_cap = _core_drum_for(total_actual, batch.wire_key, core_drum_caps, params)
        max_run = min(params.max_insulation_run, drum_cap)

        if total_actual > max_run:
            ins_mat = batch.insulation_material
            fr_flag = batch.fire_resistant
            mat_info = f' [{ins_mat}{"+" + fr_flag if fr_flag else ""}]'
            errors.append(
                f'❌ {batch.cable_mark} / {color}{mat_info}: нужно {total_actual:.0f} м, '
                f'но макс. прогон изолирования {max_run:.0f} м — невозможно без спайки.'
            )
            continue

        rw.used += raw_needed
        runs.append(InsulationRun(
            id=_uid('ИЗ'),
            color=color,
            wire_key=batch.wire_key,
            source_id=rw.id,
            source_name=rw.name,
            length=total_actual,
            drum_type=drum_name,
            for_batch_id=batch.id,
            insulation_material=batch.insulation_material,
            fire_resistant=batch.fire_resistant,
            raw_wire_consumed=round(raw_needed, 6),
        ))

    return runs, errors


# ═══════════════════════════════════════════════════════════════════════
# Шаг 5 — упаковка отрезков на выходные барабаны
# ═══════════════════════════════════════════════════════════════════════

def _assign_drums(
    cable_mark: str,
    segments: List[float],
    sources: List[str],
    cable_drum_caps: List[CableDrumCapacity],
    allow_multi_segment: bool = True,
) -> Tuple[List[DrumAssignment], List[str], List[str]]:
    """
    Распределение отрезков журнала по выходным барабанам.
    Правила:
      • Складской отрезок (src='склад') — всегда отдельный барабан.
      • Производственные отрезки из одной партии (src=batch_id) упаковываются вместе.
      • Производственные отрезки из разных партий — всегда на разных барабанах.
    """
    errors:   List[str] = []
    warnings: List[str] = []
    assignments: List[DrumAssignment] = []

    cap_obj = next((c for c in cable_drum_caps if c.cable_mark == cable_mark), None)
    if not cap_obj:
        errors.append(
            f'⚠ Не найдена таблица ёмкости барабанов для марки "{cable_mark}". '
            f'Отрезки не распределены по барабанам.'
        )
        return assignments, errors, warnings

    largest_drum = sorted(cap_obj.drum_types, key=lambda d: d.capacity)[-1]
    open_batch_drums: Dict[str, List] = {}

    for seg, src in zip(segments, sources):
        is_stock = (src == 'склад')

        if seg > largest_drum.capacity:
            warnings.append(
                f'⚠ {cable_mark}: отрезок {seg:.0f} м > макс. ёмкость барабана '
                f'{largest_drum.capacity:.0f} м ({largest_drum.name}). '
                f'Отрезок будет намотан на {math.ceil(seg / largest_drum.capacity)} барабана(ов).'
            )
            remaining_seg = seg
            while remaining_seg > 1e-6:
                take = min(remaining_seg, largest_drum.capacity)
                assignments.append(DrumAssignment(
                    id=_uid('БК'), cable_mark=cable_mark,
                    drum_type=largest_drum.name, drum_capacity=largest_drum.capacity,
                    segments=[take],
                    source='склад' if is_stock else 'партия',
                    batch_id='' if is_stock else src,
                ))
                remaining_seg -= take
            continue

        if is_stock:
            drum = cap_obj.smallest_fitting(seg)
            if drum is None:
                errors.append(
                    f'❌ {cable_mark}: складской отрезок {seg:.0f} м не помещается '
                    f'ни в один тип барабана (макс. {cap_obj.max_capacity:.0f} м).'
                )
                continue
            assignments.append(DrumAssignment(
                id=_uid('БК'), cable_mark=cable_mark,
                drum_type=drum.name, drum_capacity=drum.capacity,
                segments=[seg], source='склад', batch_id='',
            ))
            continue

        batch_id = src
        placed = False

        if allow_multi_segment:
            pool = open_batch_drums.get(batch_id, [])
            best_open = None
            best_remaining = float('inf')
            for drum_type, remaining, da in pool:
                if remaining >= seg and (remaining - seg) < best_remaining:
                    best_remaining = remaining - seg
                    best_open = (drum_type, remaining, da)

            if best_open is not None:
                drum_type, remaining, da = best_open
                idx = pool.index(best_open)
                da.segments.append(seg)
                pool[idx] = (drum_type, remaining - seg, da)
                open_batch_drums[batch_id] = pool
                placed = True

        if not placed:
            drum = cap_obj.smallest_fitting(seg)
            if drum is None:
                errors.append(
                    f'❌ {cable_mark}: производственный отрезок {seg:.0f} м не помещается '
                    f'ни в один тип барабана (макс. {cap_obj.max_capacity:.0f} м).'
                )
                continue
            da = DrumAssignment(
                id=_uid('БК'), cable_mark=cable_mark,
                drum_type=drum.name, drum_capacity=drum.capacity,
                segments=[seg], source='партия', batch_id=batch_id,
            )
            assignments.append(da)
            pool = open_batch_drums.setdefault(batch_id, [])
            pool.append((drum, drum.capacity - seg, da))

    return assignments, errors, warnings


# ═══════════════════════════════════════════════════════════════════════
# Главная функция
# ═══════════════════════════════════════════════════════════════════════

def plan(
    orders: List[CableOrder],
    raw_wires: List[RawWire],
    insulated_cores: List[InsulatedCore],
    cable_stock: List[CableStock],
    core_drum_caps: List[CoreDrumCapacity],
    cable_drum_caps: List[CableDrumCapacity],
    params: ProcessParams,
) -> PlanResult:
    """
    Планирует производство по всем заказам.
    Все входные списки — рабочие копии (deepcopy выполняется здесь).
    """
    _RUN_CTR.clear()

    raw_wires       = deepcopy(raw_wires)
    insulated_cores = deepcopy(insulated_cores)
    cable_stock     = deepcopy(cable_stock)

    all_batches:      List[TwistingBatch]    = []
    all_ins_runs:     List[InsulationRun]    = []
    all_ins_uses:     List[InsulatedCoreUse] = []
    all_stock_uses:   List[CableStockUse]    = []
    all_drum_assigns: List[DrumAssignment]   = []
    errors:   List[str] = []
    warnings: List[str] = []

    # Трекер меток групп (5ж, 4ж, …) для всего планирования
    group_labels: Dict[frozenset, str] = {}

    for order in orders:

        # ── Валидация ─────────────────────────────────────────────────
        if not order.colors:
            errors.append(f'❌ "{order.mark}": состав (цвета жил) не определён — пропускаем.')
            continue
        if not order.cross_section:
            errors.append(f'❌ "{order.mark}": сечение жил не определено — пропускаем.')
            continue
        if not order.wire_type:
            warnings.append(
                f'⚠ "{order.mark}": индекс ТПЖ не указан (wire_key="{order.cross_section}").'
            )
        if not order.insulation_material:
            warnings.append(f'⚠ "{order.mark}": материал изоляции не указан.')

        # ── Шаг 1: журнал ────────────────────────────────────────────
        journal = _fill_journal(order, params)

        journal_sum = round(sum(journal), 3)
        total = round(order.total_length, 3)
        if abs(journal_sum - total) > 0.5:
            warnings.append(
                f'⚠ "{order.mark}": сумма журнала {journal_sum} м ≠ длина заказа {total} м.'
            )

        # ── Шаг 2: склад кабеля ──────────────────────────────────────
        stock_uses, remaining_segs = _allocate_cable_stock(order, journal, cable_stock)
        all_stock_uses.extend(u for u in stock_uses if u is not None)

        # ── Шаг 3: партии скрутки (FFD с метками групп) ──────────────
        group_label = _colors_group_label(order.colors, group_labels)
        batches, batch_errors = _pack_batches(
            order, remaining_segs, core_drum_caps, params, group_label)
        errors.extend(batch_errors)
        all_batches.extend(batches)

        # ── Шаг 4: выделение жил ─────────────────────────────────────
        # 4a. Склад изолированных жил (multi-spool, посегментно)
        # 4b. ТПЖ через CP-SAT для всех партий этого заказа сразу

        # Сначала обрабатываем все партии через 4a, собираем задачи для CP-SAT
        batch_actual: Dict[str, List[float]] = {}        # batch_id → actual_segs
        batch_needs_tpzh: Dict[str, Dict[str, bool]] = {}

        for batch in batches:
            actual_segs, ins_uses, needs_tpzh, errs, warns = \
                _allocate_batch_all_colors(batch, order.colors, insulated_cores, params)
            errors.extend(errs)
            warnings.extend(warns)
            all_ins_uses.extend(ins_uses)
            batch.segments = actual_segs
            batch_actual[batch.id] = actual_segs
            batch_needs_tpzh[batch.id] = needs_tpzh

        # Собираем задачи для CP-SAT: (batch_id, color, run_length, raw_needed)
        wire_key = order.wire_key
        tpzh_tasks: List[Tuple[str, str, float, float]] = []
        for batch in batches:
            actual_segs = batch_actual[batch.id]
            if not actual_segs:
                continue
            n_segs = len(actual_segs)
            total_actual = round(sum(actual_segs), 6)
            tol_total = n_segs * params.length_tolerance_m
            startup = params.insulation_startup_loss_m
            raw_needed = total_actual + tol_total + startup

            for color in order.colors:
                if batch_needs_tpzh[batch.id].get(color, False):
                    tpzh_tasks.append((batch.id, color, total_actual, raw_needed))

        # 4b. CP-SAT (или greedy fallback)
        if tpzh_tasks:
            tpzh_drums = [r for r in raw_wires if r.wire_key == wire_key]
            assignment, cpsat_errors = _allocate_tpzh_cpsat(tpzh_tasks, tpzh_drums, params)
            errors.extend(cpsat_errors)

            # Создаём InsulationRun по результатам назначения
            for batch in batches:
                actual_segs = batch_actual[batch.id]
                if not actual_segs:
                    continue
                ins_runs, run_errors = _build_insulation_runs(
                    batch, actual_segs, batch_needs_tpzh[batch.id],
                    assignment, raw_wires, core_drum_caps, params,
                )
                errors.extend(run_errors)
                all_ins_runs.extend(ins_runs)

        # ── Шаг 5: выходные барабаны ─────────────────────────────────
        final_segs:    List[float] = []
        final_sources: List[str] = []

        # Складские отрезки
        for seg, su in zip(journal, stock_uses):
            if su is not None:
                final_segs.append(seg)
                final_sources.append('склад')

        # Производственные отрезки из партий
        for batch in batches:
            for seg in batch.segments:
                final_segs.append(seg)
                final_sources.append(batch.id)

        drum_assigns, drum_errors, drum_warns = _assign_drums(
            order.mark, final_segs, final_sources, cable_drum_caps,
            allow_multi_segment=params.allow_multi_segment_drum,
        )
        errors.extend(drum_errors)
        warnings.extend(drum_warns)
        all_drum_assigns.extend(drum_assigns)

    # ── Остатки ──────────────────────────────────────────────────────
    remaining_raw = [r for r in raw_wires if r.available > 0.5]
    remaining_ins = [c for c in insulated_cores if c.available > 0.5]
    remaining_cab = [s for s in cable_stock if s.available > 0.5]

    return PlanResult(
        orders=orders,
        batches=all_batches,
        insulation_runs=all_ins_runs,
        insulated_core_uses=all_ins_uses,
        cable_stock_uses=all_stock_uses,
        drum_assignments=all_drum_assigns,
        remaining_raw_wires=remaining_raw,
        remaining_insulated=remaining_ins,
        remaining_cable_stock=remaining_cab,
        errors=errors,
        warnings=warnings,
    )
