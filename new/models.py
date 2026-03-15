"""
Модели данных для Length Helper (v5).

Изменения относительно root/models.py:
  - CableOrder: добавлено поле flexible (из claude_insulation_v4.ipynb)
  - ProcessParams: добавлены min_segment и max_splits (из claude_insulation_v4.ipynb)
  - ProcessParams: добавлен waste_weight (коэффициент оптимизации отходов)
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional
import math


# ═══════════════════════════════════════════════════════════════════════
# ВХОДНЫЕ ДАННЫЕ (склад + заказы + параметры)
# ═══════════════════════════════════════════════════════════════════════

@dataclass
class RawWire:
    """ТПЖ — неизолированная жила на барабане. Источник для изолирования."""
    id: str            # 'Барабан А1'
    name: str          # 'ТПЖ 2,5 ок'
    cross_section: str # '2,5'  — только числовое сечение
    wire_type: str     # 'ок', 'мк', 'мкг' — класс гибкости
    length: float      # полная длина, м
    used: float = 0.0  # использовано, м

    @property
    def available(self) -> float:
        return round(self.length - self.used, 6)

    @property
    def wire_key(self) -> str:
        """Составной ключ: сечение + тип жилы, например '2,5ок'."""
        return f'{self.cross_section}{self.wire_type}'


@dataclass
class InsulatedCore:
    """Готовая изолированная жила на складе."""
    id: str                    # 'Бухта С1'
    name: str                  # 'Синяя 2,5 ок LS'
    color: str                 # 'Синяя'
    cross_section: str         # '2,5'
    wire_type: str             # 'ок', 'мк', 'мкг'
    insulation_material: str   # 'LS', 'HF', …
    fire_resistant: str        # 'FR' или ''
    length: float
    used: float = 0.0

    @property
    def available(self) -> float:
        return round(self.length - self.used, 6)

    @property
    def wire_key(self) -> str:
        return f'{self.cross_section}{self.wire_type}'


@dataclass
class CableStock:
    """Готовый кабель на складе (отрезок под конкретную марку)."""
    id: str            # 'Остаток-1'
    cable_mark: str    # 'ВВГ 3х25мк(N,PE)'
    length: float
    used: float = 0.0

    @property
    def available(self) -> float:
        return round(self.length - self.used, 6)


@dataclass
class CableOrder:
    """Заказ на производство кабеля."""
    mark: str                      # 'ВВГ 4х25мк(N)'
    total_length: float            # общая длина, м
    journal: List[float]           # отрезки кабельного журнала (пусто = нет журнала)
    colors: List[str]              # ['Натуральная', 'Синяя', 'Черная', 'Коричневая']
    cross_section: str             # '25'  — только числовое сечение
    wire_type: str = ''            # 'ок', 'мк', 'мкг'
    fire_resistant: str = ''       # 'FR' или ''
    insulation_material: str = ''  # 'LS', 'HF', …
    flexible: bool = False
    # flexible=True → алгоритм может увеличить длину жилы для поглощения остатков барабана.
    # «Сдача максимальной длиной» вместо «сдача одной длиной».

    @property
    def has_journal(self) -> bool:
        return len(self.journal) > 0

    @property
    def wire_key(self) -> str:
        """Составной ключ: сечение + тип, например '25мк'."""
        return f'{self.cross_section}{self.wire_type}'

    @property
    def n_colors(self) -> int:
        return len(self.colors)


@dataclass
class DrumType:
    """Тип барабана с его ёмкостью."""
    name: str          # 'Б-630', '№12'
    capacity: float    # максимальная длина, м


@dataclass
class CoreDrumCapacity:
    """Ёмкость барабанов/катушек для изолированных жил, по типу жилы."""
    wire_key: str              # '2,5ок', '25мк' — сечение + тип (совмещённый ключ)
    drum_types: List[DrumType] # отсортированы по ёмкости возрастанием

    @property
    def max_capacity(self) -> float:
        return max((d.capacity for d in self.drum_types), default=0.0)

    def smallest_fitting(self, length: float) -> Optional[DrumType]:
        """Наименьший барабан, куда вмещается length метров."""
        for d in sorted(self.drum_types, key=lambda x: x.capacity):
            if d.capacity >= length:
                return d
        return None


@dataclass
class CableDrumCapacity:
    """Ёмкость выходных барабанов для готового кабеля, по марке."""
    cable_mark: str
    drum_types: List[DrumType]

    @property
    def max_capacity(self) -> float:
        return max((d.capacity for d in self.drum_types), default=0.0)

    def smallest_fitting(self, length: float) -> Optional[DrumType]:
        for d in sorted(self.drum_types, key=lambda x: x.capacity):
            if d.capacity >= length:
                return d
        return None


@dataclass
class ProcessParams:
    """Параметры производственного процесса."""
    max_insulation_run: float = 4500.0      # макс. прогон изолирования, м
    max_twisting_run: float = 2100.0        # макс. партия скрутки, м
    min_construction_length: float = 450.0  # мин. строительная длина, м
    allow_splicing: bool = False
    allow_multi_segment_drum: bool = True   # разрешить несколько отрезков на одном приёмном барабане
    strategy: str = 'Экономия'

    # ── Параметры изолирования (ТЗ §10) ─────────────────────────────────
    insulation_startup_loss_m: float = 5.0
    # П1: потери на заправку изолировочной линии (м на каждый прогон).
    # Расход ТПЖ = длина прогона + startup_loss. Плановая длина П/Ф не меняется.

    length_tolerance_m: float = 3.0
    # П4: запас на обрезку торцов (м на каждый сегмент партии скрутки).

    keep_journal_order: bool = False
    # П5: True → при упаковке партий НЕ сортировать, брать в порядке журнала.

    waste_warning_threshold_m: float = 50.0
    # П6: остаток катушки < порога → предупреждение «вероятно в отход».

    # ── Параметры CP-SAT оптимизатора (из claude_insulation_v4) ─────────
    min_segment: float = 330.0
    # Физический минимум одного сегмента жилы на приёмной катушке (м).
    # Меньше значение → больше гибкости, но физически нельзя мотать слишком мало.

    max_splits: int = 1
    # Максимум барабанов ТПЖ на один прогон изолирования одного цвета.
    # 1 = без стыков (по умолчанию; стык жилы = нарушение ТУ).
    # >1 = допустить стыки (только если явно разрешено).

    waste_weight: int = 1000
    # Коэффициент приоритета минимизации отходов ТПЖ в целевой функции CP-SAT.
    # 0 = не минимизировать отходы (только смены); 2000 = отходы критичны.

    cpsat_time_limit: float = 30.0
    # Лимит времени решателя CP-SAT, секунд.


# ═══════════════════════════════════════════════════════════════════════
# РЕЗУЛЬТАТЫ ПЛАНИРОВАНИЯ
# ═══════════════════════════════════════════════════════════════════════

@dataclass
class TwistingBatch:
    """Партия скрутки — один прогон скрутки, несколько отрезков журнала."""
    id: str
    cable_mark: str
    segments: List[float]           # отрезки журнала в этой партии, м
    wire_key: str                   # '2,5ок', '25мк' — тип жилы для поиска жил
    colors: List[str]               # цвета жил (из заказа)
    insulation_material: str = ''   # 'LS', 'HF', …
    fire_resistant: str = ''        # 'FR' или ''
    group_label: str = ''           # метка группы ('5ж', '4ж-А', …) — из Трека B

    @property
    def total_length(self) -> float:
        return sum(self.segments)


@dataclass
class InsulationRun:
    """Прогон изолирования — один цвет, один барабан ТПЖ → один моток жилы."""
    id: str
    color: str
    wire_key: str                   # '2,5ок', '25мк'
    source_id: str                  # id RawWire
    source_name: str                # имя ТПЖ-барабана для инструкции
    length: float                   # длина прогона, м (что получим с линии)
    drum_type: str                  # тип барабана/катушки (Б-630 и т.п.)
    for_batch_id: str               # к какой партии скрутки относится
    insulation_material: str = ''   # 'LS', 'HF', …
    fire_resistant: str = ''        # 'FR' или ''
    raw_wire_consumed: float = 0.0
    # П1+П4: сколько реально снимается с барабана ТПЖ = length + startup_loss + tolerance.
    # Если 0 — считать равным length (обратная совместимость).


@dataclass
class InsulatedCoreUse:
    """Использование готовой изолированной жилы со склада."""
    id: str
    color: str
    wire_key: str                   # '2,5ок'
    source_id: str                  # id InsulatedCore
    source_name: str
    length: float                   # сколько взяли, м
    remainder: float                # остаток на складе после использования, м
    for_batch_id: str
    covered_segments: List[int] = field(default_factory=list)
    # индексы отрезков batch.segments, покрытых этой катушкой
    spool_index: int = 1
    # порядковый номер катушки данного цвета в рамках партии (1, 2, …)


@dataclass
class CableStockUse:
    """Использование готового кабеля со склада под отрезок журнала."""
    id: str
    cable_mark: str
    source_id: str          # id CableStock
    segment_length: float   # длина отрезка журнала, м
    remainder: float        # остаток на складе, м


@dataclass
class DrumAssignment:
    """Назначение выходного барабана под группу отрезков журнала."""
    id: str
    cable_mark: str
    drum_type: str          # '№12'
    drum_capacity: float    # ёмкость барабана, м
    segments: List[float]   # отрезки кабельного журнала на этом барабане
    source: str             # 'партия' или 'склад'
    batch_id: str = ''      # id TwistingBatch (если из партии)

    @property
    def total_length(self) -> float:
        return sum(self.segments)


@dataclass
class PlanResult:
    """Итоговый результат планирования."""
    orders: List[CableOrder]

    batches: List[TwistingBatch]
    insulation_runs: List[InsulationRun]
    insulated_core_uses: List[InsulatedCoreUse]
    cable_stock_uses: List[CableStockUse]
    drum_assignments: List[DrumAssignment]

    # Остатки после планирования
    remaining_raw_wires: List[RawWire]
    remaining_insulated: List[InsulatedCore]
    remaining_cable_stock: List[CableStock]

    errors: List[str]
    warnings: List[str]
