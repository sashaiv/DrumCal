"""
Алгоритм производственного планирования.

Шаги:
  1. Заполнить кабельный журнал (если не задан).
  2. Покрыть отрезки журнала из готового кабеля на складе.
  3. Упаковать оставшиеся отрезки в партии скрутки (FFD bin-packing ≤ max_twisting_run).
  4. Для каждой партии × цвет — выделить жилу (склад изолир. → ТПЖ с планом изолирования).
     Жилы с разными материалами изоляции/огнестойкостью изолируются отдельными прогонами.
  5. Упаковать ВСЕ отрезки журнала на выходные барабаны (min drum bin-packing).
"""
from __future__ import annotations
import math
from copy import deepcopy
from typing import List, Tuple, Dict, Optional

from models import (
    RawWire, InsulatedCore, CableStock, CableOrder,
    CoreDrumCapacity, CableDrumCapacity, ProcessParams,
    TwistingBatch, InsulationRun, InsulatedCoreUse,
    CableStockUse, DrumAssignment, PlanResult, DrumType,
)

_RUN_CTR: Dict[str, int] = {}   # счётчики id внутри одного вызова plan()


def _uid(prefix: str) -> str:
    _RUN_CTR[prefix] = _RUN_CTR.get(prefix, 0) + 1
    return f'{prefix}-{_RUN_CTR[prefix]:03d}'


# ═══════════════════════════════════════════════════════════════════════
# Шаг 1 — кабельный журнал
# ═══════════════════════════════════════════════════════════════════════

def _fill_journal(order: CableOrder, params: ProcessParams) -> List[float] | List[int]:
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
    segments: List[float] | List[int],
    cable_stock: List[CableStock],
) -> Tuple[List[Optional[CableStockUse]], List[float]]:
    """
    Для каждого отрезка журнала пробует найти подходящий кабель на складе.
    Возвращает:
        stock_uses  — список CableStockUse (None для отрезков без склада)
        remaining   — отрезки, не покрытые со склада (идут в производство)
    """
    stock_uses: List[Optional[CableStockUse]] = []
    remaining: List[float] = []

    for seg in segments:
        candidates = [
            s for s in cable_stock
            if s.cable_mark == order.mark and s.available >= seg
        ]
        if candidates:
            best = min(candidates, key=lambda s: s.available)
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
# Шаг 3 — упаковка отрезков в партии скрутки (FFD)
# ═══════════════════════════════════════════════════════════════════════

def _max_batch_size(
    wire_key: str,
    core_drum_caps: List[CoreDrumCapacity],
    params: ProcessParams,
) -> float:
    """Максимальная длина одной партии скрутки с учётом ёмкости катушек жил."""
    cap_obj = next((c for c in core_drum_caps if c.wire_key == wire_key), None)
    max_drum = cap_obj.max_capacity if cap_obj else params.max_insulation_run
    return min(params.max_twisting_run, max_drum)


def _pack_batches(
    order: CableOrder,
    segments: List[float],
    core_drum_caps: List[CoreDrumCapacity],
    params: ProcessParams,
) -> Tuple[List[TwistingBatch], List[str]]:
    """
    Bin-packing отрезков в партии скрутки.

    Если params.keep_journal_order=False (по умолчанию): First Fit Decreasing —
    сортировка по убыванию для лучшего заполнения партий.
    Если params.keep_journal_order=True (П5): отрезки берутся в порядке журнала
    без сортировки — физическая последовательность намотки сохраняется.
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
            ))

    return batches, errors


# ═══════════════════════════════════════════════════════════════════════
# Шаг 4 — выделение жил для партии
# ═══════════════════════════════════════════════════════════════════════
#
# АРХИТЕКТУРА:
#   _allocate_batch_all_colors  — основная функция; обрабатывает все цвета
#     вместе посегментно. Если ни одна катушка не покрывает сегмент,
#     уменьшает длину до максимально возможного (не менее min_construction_length).
#     Не использует спайку.
#   _allocate_cores_for_batch  — вспомогательная; один цвет, один прогон
#     (используется внутри _allocate_batch_all_colors для прогонов изолирования).
# ═══════════════════════════════════════════════════════════════════════

def _core_drum_for(
    length: float,
    wire_key: str,
    core_drum_caps: List[CoreDrumCapacity],
    params: ProcessParams,
) -> Tuple[str, float]:
    """Возвращает (имя_барабана, ёмкость) — наименьший подходящий."""
    cap_obj = next((c for c in core_drum_caps if c.wire_key == wire_key), None)
    if cap_obj:
        drum = cap_obj.smallest_fitting(length)
        if drum:
            return drum.name, drum.capacity
    # Нет таблицы — используем параметр
    return '—', params.max_insulation_run


def _multispool_from_stock(
    batch: TwistingBatch,
    color: str,
    wire_key: str,
    candidates: List[InsulatedCore],
    waste_threshold: float = 50.0,
) -> Tuple[List[InsulatedCoreUse], Optional[str], List[str]]:
    """
    Покрывает все сегменты партии из нескольких катушек склада.

    Алгоритм: итерируем по сегментам партии по порядку.
    Если текущая катушка имеет достаточный остаток — берём из неё.
    Иначе — ищем следующую катушку (наименьшую подходящую) и переходим на неё.
    Смена катушки происходит только между сегментами.

    Возвращает (список InsulatedCoreUse, строка-ошибка или None, список предупреждений).
    """
    segs = batch.segments
    uses: List[InsulatedCoreUse] = []
    warn: List[str] = []
    spool_n = 0

    curr_spool: Optional[InsulatedCore] = None
    curr_balance = 0.0          # остаток текущей катушки
    curr_seg_indices: List[int] = []

    def _commit():
        """Фиксирует накопленные сегменты на curr_spool."""
        nonlocal curr_spool, curr_balance, curr_seg_indices, spool_n
        if curr_spool is None or not curr_seg_indices:
            return
        taken = round(sum(segs[j] for j in curr_seg_indices), 6)
        rem   = round(curr_spool.available - taken, 6)
        curr_spool.used += taken
        spool_n += 1
        # П6: предупреждение об остатке→отход
        if rem < waste_threshold:
            warn.append(
                f'⚠ {batch.cable_mark} / {color}: катушка «{curr_spool.name}» '
                f'после использования — остаток {rem:.0f} м (< {waste_threshold:.0f} м), '
                f'вероятно в отход.'
            )
        uses.append(InsulatedCoreUse(
            id=_uid('ГЖ'),
            color=color,
            wire_key=wire_key,
            source_id=curr_spool.id,
            source_name=curr_spool.name,
            length=taken,
            remainder=rem,
            for_batch_id=batch.id,
            covered_segments=list(curr_seg_indices),
            spool_index=spool_n,
        ))
        curr_spool = None
        curr_balance = 0.0
        curr_seg_indices = []

    for seg_i, seg in enumerate(segs):
        if curr_spool is not None and curr_balance >= seg - 1e-6:
            # Продолжаем на текущей катушке
            curr_seg_indices.append(seg_i)
            curr_balance -= seg
        else:
            # Фиксируем текущую катушку, ищем следующую
            _commit()
            fitting = sorted(
                [c for c in candidates if c.available >= seg - 1e-6],
                key=lambda c: c.available,   # наименьшая подходящая
            )
            if not fitting:
                return [], (
                    f'❌ {batch.cable_mark} / {color}: '
                    f'нет катушки ≥{seg:.0f} м для сегмента {seg_i + 1}.'
                ), warn
            curr_spool   = fitting[0]
            curr_balance = curr_spool.available - seg
            curr_seg_indices = [seg_i]

    _commit()
    return uses, None, warn


def _allocate_cores_for_batch(
    batch: TwistingBatch,
    color: str,
    insulated_cores: List[InsulatedCore],
    raw_wires: List[RawWire],
    core_drum_caps: List[CoreDrumCapacity],
    params: ProcessParams,
) -> Tuple[List[InsulationRun], List[InsulatedCoreUse], List[str]]:
    """
    Выделяет жилы для одной партии скрутки — с поддержкой нескольких катушек.

    Приоритет 1 — склад изолированных жил:
      • Если одна катушка покрывает всю партию → 1 InsulatedCoreUse.
      • Если нужна смена катушки между сегментами → несколько InsulatedCoreUse.
    Приоритет 2 — прогон изолирования с ТПЖ (один прогон = вся партия).

    Возвращает (runs, uses, errors).
    """
    needed   = batch.total_length
    wire_key = batch.wire_key
    ins_mat  = batch.insulation_material
    fr_flag  = batch.fire_resistant

    # Кандидаты со склада: совпадение по всем параметрам
    candidates = [
        c for c in insulated_cores
        if c.color           == color
        and c.wire_key       == wire_key
        and c.insulation_material == ins_mat
        and c.fire_resistant == fr_flag
    ]
    total_stock = sum(c.available for c in candidates)

    # ── Приоритет 1: склад жил ───────────────────────────────────────
    if total_stock >= needed - 1e-6:
        # Быстрый путь: одна катушка покрывает всю партию
        single = sorted(
            [c for c in candidates if c.available >= needed - 1e-6],
            key=lambda c: c.available,
        )
        if single:
            spool = single[0]
            rem   = round(spool.available - needed, 6)
            spool.used += needed
            warns = []
            if rem < params.waste_warning_threshold_m:
                warns.append(
                    f'⚠ {batch.cable_mark} / {color}: катушка «{spool.name}» '
                    f'после использования — остаток {rem:.0f} м '
                    f'(< {params.waste_warning_threshold_m:.0f} м), вероятно в отход.'
                )
            return [], [InsulatedCoreUse(
                id=_uid('ГЖ'),
                color=color,
                wire_key=wire_key,
                source_id=spool.id,
                source_name=spool.name,
                length=needed,
                remainder=rem,
                for_batch_id=batch.id,
                covered_segments=list(range(len(batch.segments))),
                spool_index=1,
            )], warns

        # Несколько катушек
        uses, err, multi_warns = _multispool_from_stock(
            batch, color, wire_key, candidates,
            waste_threshold=params.waste_warning_threshold_m,
        )
        if not err:
            return [], uses, multi_warns
        # Ошибка внутри multi-spool — передаём в errors
        return [], [], [err]

    # ── Приоритет 2: прогон изолирования с ТПЖ ───────────────────────
    # П4: запас на обрезку торцов: каждый сегмент партии требует доп. length_tolerance_m.
    # П1: потери на заправку изолировочной линии (один раз на весь прогон).
    n_segs      = len(batch.segments)
    tol_total   = n_segs * params.length_tolerance_m        # отходы на обрезку торцов
    startup     = params.insulation_startup_loss_m          # потери на заправку линии
    needed_raw  = needed + tol_total + startup              # сколько снять с барабана ТПЖ

    drum_name, drum_cap = _core_drum_for(needed, wire_key, core_drum_caps, params)
    max_run = min(params.max_insulation_run, drum_cap)

    if needed > max_run:
        mat_info = f' [{ins_mat}{"+" + fr_flag if fr_flag else ""}]'
        return [], [], [
            f'❌ {batch.cable_mark} / {color}{mat_info}: нужно {needed:.0f} м, '
            f'но макс. прогон изолирования {max_run:.0f} м — невозможно без спайки.'
        ]

    rw_candidates = [r for r in raw_wires if r.wire_key == wire_key and r.available >= needed_raw]
    if not rw_candidates:
        # П1: проверим, хватает ли ТПЖ без учёта потерь (чтобы дать точный диагноз)
        total_avail = sum(r.available for r in raw_wires if r.wire_key == wire_key)
        mat_info = f' [{ins_mat}{"+" + fr_flag if fr_flag else ""}]'
        if total_avail >= needed - 1e-6:
            return [], [], [
                f'❌ {batch.cable_mark} / {color}{mat_info}: ТПЖ {wire_key} хватит на прогон '
                f'({needed:.0f} м), но с учётом потерь на заправку ({startup:.0f} м) и '
                f'допуска на обрезку ({tol_total:.0f} м) нужно {needed_raw:.0f} м — '
                f'доступно {total_avail:.0f} м.'
            ]
        return [], [], [
            f'❌ {batch.cable_mark} / {color}{mat_info}: нужно {needed_raw:.0f} м ТПЖ {wire_key} '
            f'(прогон {needed:.0f} м + заправка {startup:.0f} м + допуск {tol_total:.0f} м), '
            f'доступно суммарно {total_avail:.0f} м (ни одного куска достаточной длины).'
        ]

    best_rw = max(rw_candidates, key=lambda r: r.available)
    best_rw.used += needed_raw

    return [InsulationRun(
        id=_uid('ИЗ'),
        color=color,
        wire_key=wire_key,
        source_id=best_rw.id,
        source_name=best_rw.name,
        length=needed,
        drum_type=drum_name,
        for_batch_id=batch.id,
        insulation_material=ins_mat,
        fire_resistant=fr_flag,
        raw_wire_consumed=round(needed_raw, 6),
    )], [], []


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
      • Складской отрезок (src='склад') — всегда отдельный барабан,
        НИКОГДА не смешивается с производственными или другими складскими.
      • Производственные отрезки из одной партии (src=batch_id) упаковываются
        вместе на барабан (greedy bin-packing), если allow_multi_segment=True.
      • Производственные отрезки из разных партий — всегда на разных барабанах.
    """
    errors:   List[str] = []
    warnings: List[str] = []
    assignments: List[DrumAssignment] = []

    cap_obj = next(
        (c for c in cable_drum_caps if c.cable_mark == cable_mark), None)
    if not cap_obj:
        errors.append(
            f'⚠ Не найдена таблица ёмкости барабанов для марки "{cable_mark}". '
            f'Отрезки не распределены по барабанам.'
        )
        return assignments, errors, warnings

    largest_drum = sorted(cap_obj.drum_types, key=lambda d: d.capacity)[-1]

    # Открытые производственные барабаны, раздельно по batch_id:
    # {batch_id: [(DrumType, remaining_capacity, DrumAssignment), ...]}
    open_batch_drums: Dict[str, List] = {}

    for seg, src in zip(segments, sources):
        is_stock = (src == 'склад')

        # ── Негабаритный отрезок ─────────────────────────────────────
        if seg > largest_drum.capacity:
            warnings.append(
                f'⚠ {cable_mark}: отрезок {seg:.0f} м > макс. ёмкость барабана '
                f'{largest_drum.capacity:.0f} м ({largest_drum.name}). '
                f'Отрезок будет намотан на {math.ceil(seg / largest_drum.capacity)} барабана(ов).'
            )
            remaining_seg = seg
            while remaining_seg > 1e-6:
                take = min(remaining_seg, largest_drum.capacity)
                da = DrumAssignment(
                    id=_uid('БК'),
                    cable_mark=cable_mark,
                    drum_type=largest_drum.name,
                    drum_capacity=largest_drum.capacity,
                    segments=[take],
                    source='склад' if is_stock else 'партия',
                    batch_id='' if is_stock else src,
                )
                assignments.append(da)
                remaining_seg -= take
            continue

        # ── Складской отрезок: всегда отдельный барабан ───────────────
        if is_stock:
            drum = cap_obj.smallest_fitting(seg)
            if drum is None:
                errors.append(
                    f'❌ {cable_mark}: складской отрезок {seg:.0f} м не помещается '
                    f'ни в один тип барабана (макс. {cap_obj.max_capacity:.0f} м).'
                )
                continue
            da = DrumAssignment(
                id=_uid('БК'),
                cable_mark=cable_mark,
                drum_type=drum.name,
                drum_capacity=drum.capacity,
                segments=[seg],
                source='склад',
                batch_id='',
            )
            assignments.append(da)
            # Не добавляем в open_batch_drums — склад не смешивается с производством
            continue

        # ── Производственный отрезок ──────────────────────────────────
        batch_id = src
        placed = False

        if allow_multi_segment:
            # Ищем открытый барабан для этой же партии
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
                id=_uid('БК'),
                cable_mark=cable_mark,
                drum_type=drum.name,
                drum_capacity=drum.capacity,
                segments=[seg],
                source='партия',
                batch_id=batch_id,
            )
            assignments.append(da)
            pool = open_batch_drums.setdefault(batch_id, [])
            pool.append((drum, drum.capacity - seg, da))

    return assignments, errors, warnings


# ═══════════════════════════════════════════════════════════════════════
# Главная функция
# ═══════════════════════════════════════════════════════════════════════
# Шаг 4b — единый многоцветный аллокатор с обрезкой сегментов
# ═══════════════════════════════════════════════════════════════════════

def _allocate_batch_all_colors(
    batch: TwistingBatch,
    colors: List[str],
    insulated_cores: List[InsulatedCore],
    raw_wires: List[RawWire],
    core_drum_caps: List[CoreDrumCapacity],
    params: ProcessParams,
) -> Tuple[List[float], List[InsulationRun], List[InsulatedCoreUse], List[str], List[str]]:
    """
    Распределяет жилы для партии по ВСЕМ цветам одновременно,
    обрабатывая сегменты один за другим.

    Если ни одна отдельная катушка не покрывает плановый сегмент:
      • вычисляется фактически возможная длина = min по цветам от max доступной катушки;
      • если фактическая длина ≥ min_construction_length → сегмент сокращается, предупреждение;
      • если < min_construction_length → сегмент пропускается, предупреждение.
    Спайки (stыковки) не используются.

    Возвращает (actual_segs, ins_runs, ins_uses, errors, warnings).
    """
    min_len  = params.min_construction_length
    ins_mat  = batch.insulation_material
    fr_flag  = batch.fire_resistant
    wire_key = batch.wire_key

    # ── Классификация цветов: склад vs изолирование ─────────────────
    stock_cands: Dict[str, List[InsulatedCore]] = {}
    from_stock:  Dict[str, bool] = {}
    for color in colors:
        cands = [
            c for c in insulated_cores
            if c.color == color
            and c.wire_key == wire_key
            and c.insulation_material == ins_mat
            and c.fire_resistant == fr_flag
        ]
        stock_cands[color] = sorted(cands, key=lambda c: c.available)
        from_stock[color]  = sum(c.available for c in cands) >= min_len - 1e-6

    # ── Состояние жадного алгоритма на катушку ──────────────────────
    # curr[color] = (InsulatedCore | None, remaining_balance)
    curr: Dict[str, tuple] = {c: (None, 0.0) for c in colors if from_stock[c]}

    # Накопленные использования: (color, spool_id) → dict
    pending: Dict[tuple, dict] = {}
    spool_num: Dict[str, int]  = {c: 0 for c in colors}

    actual_segs: List[float] = []
    all_ins_uses: List[InsulatedCoreUse] = []
    warnings:   List[str] = []
    errors:     List[str] = []

    waste_thresh = params.waste_warning_threshold_m

    def _commit(color: str, spool: InsulatedCore, seg_indices: list, total: float):
        rem = round(spool.available - total, 6)
        spool.used += total
        spool_num[color] += 1
        # П6: предупреждение об остатке→отход
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

    for planned_seg in batch.segments:

        # Для каждого складского цвета — определяем максимально возможную длину
        color_max: Dict[str, float]  = {}
        color_spool: Dict[str, tuple] = {}  # color → (spool, is_new_spool)

        for color in colors:
            if not from_stock[color]:
                continue
            cs, cb = curr[color]

            if cs is not None and cb >= planned_seg - 1e-6:
                color_max[color]   = planned_seg
                color_spool[color] = (cs, False)
            else:
                # Текущая катушка не тянет — ищем ДРУГУЮ (c is not cs).
                # Нельзя включать cs в поиск: её реальный остаток (cb) уже учтён
                # через curr[color], но spool.available ещё не обновлён (lazy commit).
                fitting = sorted(
                    [c for c in stock_cands[color]
                     if c is not cs and c.available >= planned_seg - 1e-6],
                    key=lambda c: c.available,
                )
                if fitting:
                    color_max[color]   = planned_seg
                    color_spool[color] = (fitting[0], True)
                else:
                    # Нет новой катушки >= сегменту; ищем наибольшую >= min_len
                    # (можно использовать остаток cs, если он >= min_len)
                    options = [(c.available, c) for c in stock_cands[color]
                               if c is not cs and c.available >= min_len - 1e-6]
                    if cs is not None and cb >= min_len - 1e-6:
                        options.append((cb, cs))
                    if options:
                        best_avail, best_spool = max(options, key=lambda x: x[0])
                        color_max[color]   = best_avail
                        color_spool[color] = (best_spool, best_spool is not cs)
                    else:
                        color_max[color]   = 0.0
                        color_spool[color] = (None, False)

        # Фактическая длина = минимум по всем складским цветам
        if color_max:
            actual = min(color_max.values())
        else:
            actual = planned_seg  # все цвета — изолирование, ограничений нет

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

        ai = len(actual_segs)   # индекс этого сегмента в actual_segs
        actual_segs.append(round(actual, 3))

        # Выделяем actual метров из каждого складского цвета
        for color in colors:
            if not from_stock[color]:
                continue
            cs, cb = curr[color]
            spool, is_new = color_spool[color]

            if spool is None:
                continue

            if is_new:
                # Фиксируем предыдущую катушку
                if cs is not None:
                    k = (color, cs.id)
                    if k in pending:
                        pu = pending.pop(k)
                        _commit(color, pu['spool'], pu['segs'], pu['total'])
                # Начинаем новую
                curr[color] = (spool, spool.available - actual)
                pending[(color, spool.id)] = {
                    'spool': spool, 'segs': [ai], 'total': actual}
            else:
                # Продолжаем на текущей катушке
                k = (color, spool.id)
                if k in pending:
                    pending[k]['segs'].append(ai)
                    pending[k]['total'] += actual
                else:
                    pending[k] = {'spool': spool, 'segs': [ai], 'total': actual}
                curr[color] = (spool, cb - actual)

    # Фиксируем оставшиеся накопленные использования
    for (color, _), pu in list(pending.items()):
        _commit(color, pu['spool'], pu['segs'], pu['total'])

    # ── Прогоны изолирования для цветов без склада ──────────────────
    all_ins_runs: List[InsulationRun] = []
    if actual_segs:
        total_actual = round(sum(actual_segs), 6)
        for color in colors:
            if from_stock[color]:
                continue
            # Создаём временный batch с actual_segs для поиска ТПЖ
            tmp_batch = TwistingBatch(
                id=batch.id, cable_mark=batch.cable_mark,
                segments=actual_segs, wire_key=wire_key,
                colors=colors, insulation_material=ins_mat, fire_resistant=fr_flag,
            )
            runs, _, errs = _allocate_cores_for_batch(
                tmp_batch, color, insulated_cores, raw_wires, core_drum_caps, params)
            errors.extend(errs)
            all_ins_runs.extend(runs)

    return actual_segs, all_ins_runs, all_ins_uses, errors, warnings


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

    raw_wires = deepcopy(raw_wires)
    insulated_cores = deepcopy(insulated_cores)
    cable_stock = deepcopy(cable_stock)

    all_batches:        List[TwistingBatch] = []
    all_ins_runs:       List[InsulationRun] = []
    all_ins_uses:       List[InsulatedCoreUse] = []
    all_stock_uses:     List[CableStockUse] = []
    all_drum_assigns:   List[DrumAssignment] = []
    errors:   List[str] = []
    warnings: List[str] = []

    for order in orders:

        # ── Проверка наличия состава кабеля ──────────────────────────
        if not order.colors:
            errors.append(
                f'❌ "{order.mark}": состав (цвета жил) не определён — пропускаем заказ.'
            )
            continue
        if not order.cross_section:
            errors.append(
                f'❌ "{order.mark}": сечение жил не определено — пропускаем заказ.'
            )
            continue
        if not order.wire_type:
            warnings.append(
                f'⚠ "{order.mark}": индекс ТПЖ не указан — используем wire_key="{order.cross_section}".'
            )
        if not order.insulation_material:
            warnings.append(
                f'⚠ "{order.mark}": материал изоляции не указан.'
            )

        # ── Шаг 1: журнал ────────────────────────────────────────────
        journal = _fill_journal(order, params)

        journal_sum = round(sum(journal), 3)
        total = round(order.total_length, 3)
        if abs(journal_sum - total) > 0.5:
            warnings.append(
                f'⚠ "{order.mark}": сумма журнала {journal_sum} м '
                f'≠ длина заказа {total} м.'
            )

        # ── Шаг 2: склад кабеля ──────────────────────────────────────
        stock_uses, remaining_segs = _allocate_cable_stock(
            order, journal, cable_stock)
        all_stock_uses.extend(u for u in stock_uses if u is not None)

        all_segs_for_order:    List[float] = []
        all_sources_for_order: List[str] = []
        for seg, su in zip(journal, stock_uses):
            all_segs_for_order.append(seg)
            all_sources_for_order.append(
                'склад' if su is not None else '__production__')

        # ── Шаг 3: партии скрутки ────────────────────────────────────
        batches, batch_errors = _pack_batches(
            order, remaining_segs, core_drum_caps, params)
        errors.extend(batch_errors)
        all_batches.extend(batches)

        # Обновляем метки источников для производственных отрезков
        for i, src in enumerate(all_sources_for_order):
            if src == '__production__':
                seg = all_segs_for_order[i]
                matched_batch = None
                for batch in batches:
                    if seg in batch.segments:
                        matched_batch = batch.id
                        batch.segments.remove(seg)
                        batch.segments.insert(0, seg)
                        break
                all_sources_for_order[i] = matched_batch or '???'

        # ── Шаг 4: жилы для каждой партии (с обрезкой сегментов) ───────
        # _allocate_batch_all_colors обрабатывает все цвета вместе,
        # чтобы правильно вычислять фактическую длину через все позиции.
        for batch in batches:
            actual_segs, ins_runs, ins_uses, errs, warns = \
                _allocate_batch_all_colors(
                    batch, order.colors, insulated_cores,
                    raw_wires, core_drum_caps, params,
                )
            errors.extend(errs)
            warnings.extend(warns)
            all_ins_runs.extend(ins_runs)
            all_ins_uses.extend(ins_uses)
            batch.segments = actual_segs   # обновляем фактические сегменты

        # ── Шаг 5: выходные барабаны ─────────────────────────────────
        # Перестраиваем final_segs из фактических сегментов партий
        # (batch.segments уже обрезаны на шаге 4)
        final_segs:    List[float] = []
        final_sources: List[str] = []

        # Сначала — отрезки из склада готового кабеля
        for seg, su in zip(journal, stock_uses):
            if su is not None:
                final_segs.append(seg)
                final_sources.append('склад')

        # Затем — производственные отрезки из батчей
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
