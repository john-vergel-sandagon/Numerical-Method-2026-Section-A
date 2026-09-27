"""cube_solver_rev3_loads.py - Rev 3 Loads, Diaphragms, Combinations, and Visualization Framework.

Builds upon Rev 3 baseline by adding:
- LoadCase, NodalLoad, MemberDistributedLoad, MemberPointLoad, Diaphragm, LoadCombination, and TemperatureLoad
- Baseline Load Cases 1 to 9 (including dead, live, point, wind, seismic, and thermal strain)
- Rigid roof diaphragm definition (Master Node N5, coupling UX, UZ, RY)
- 30 NSCP 2015 LRFD and ASD combinations (including temperature combinations)
- Interactive 3D visualizer with layer toggles (grid, load bands, opacity scaling)
- Automated verification CLI: python cube_solver_rev3_loads.py --verify-loads
"""

from __future__ import annotations

import argparse
import math
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, List, Dict, Tuple, Optional

import matplotlib.pyplot as plt
from matplotlib.widgets import RadioButtons, TextBox, Button
from mpl_toolkits.mplot3d.art3d import Poly3DCollection, Line3DCollection

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required: pip install openpyxl")

# ─── Baseline Configuration & Model Topology ─────────────────────────

REVISION = "Rev. 3 - Loads Framework"
SIDE_LENGTH_M = 6.0
GLOBAL_DOFS = ("UX", "UY", "UZ", "RX", "RY", "RZ")

DEFAULT_BEAM_SHAPE_IMPERIAL = "W12X26"
DEFAULT_BEAM_SHAPE_METRIC = "W310X38.7"
DEFAULT_COLUMN_SHAPE_IMPERIAL = "W10X33"
DEFAULT_COLUMN_SHAPE_METRIC = "W250X49.1"
DEFAULT_BEAM_MATERIAL = "A992"
DEFAULT_COLUMN_MATERIAL = "A992"

OPACITY_DISTRIBUTED_BAND = 0.30
OPACITY_SELF_WEIGHT_BAND = 0.18

METRIC_SCALE = {
    "Ix": 1e6, "Iy": 1e6, "Iz": 1e6,
    "Zx": 1e3, "Zy": 1e3, "Sx": 1e3, "Sy": 1e3, "Sz": 1e3,
    "J": 1e3, "Cw": 1e9,
}

IMP_COL = {
    "type": 0, "label": 2, "W": 4, "A": 5, "d": 6,
    "bf": 11, "tw": 16, "tf": 19,
    "Ix": 38, "Zx": 39, "Sx": 40, "rx": 41,
    "Iy": 42, "Zy": 43, "Sy": 44, "ry": 45,
    "J": 49, "Cw": 50,
}
MET_COL = {
    "label": 85, "W": 86, "A": 87, "d": 88,
    "bf": 93, "tw": 98, "tf": 101,
    "Ix": 120, "Zx": 121, "Sx": 122, "rx": 123,
    "Iy": 124, "Zy": 125, "Sy": 126, "ry": 127,
    "J": 131, "Cw": 132,
}

NODES = {
    1: (0.0, 0.0, 0.0), 2: (6.0, 0.0, 0.0),
    3: (6.0, 0.0, 6.0), 4: (0.0, 0.0, 6.0),
    5: (0.0, 6.0, 0.0), 6: (6.0, 6.0, 0.0),
    7: (6.0, 6.0, 6.0), 8: (0.0, 6.0, 6.0),
}

SUPPORTS = {
    node: {"UX": True, "UY": True, "UZ": True,
           "RX": False, "RY": False, "RZ": False}
    for node in (1, 2, 3, 4)
}

MEMBERS = [
    {"id": 1, "i": 1, "j": 2, "type": "Base beam", "beta_deg": 0.0, "pinned": True},
    {"id": 2, "i": 2, "j": 3, "type": "Base beam", "beta_deg": 0.0, "pinned": True},
    {"id": 3, "i": 3, "j": 4, "type": "Base beam", "beta_deg": 0.0, "pinned": True},
    {"id": 4, "i": 4, "j": 1, "type": "Base beam", "beta_deg": 0.0, "pinned": True},
    {"id": 5, "i": 5, "j": 6, "type": "Roof beam", "beta_deg": 0.0, "pinned": True},
    {"id": 6, "i": 6, "j": 7, "type": "Roof beam", "beta_deg": 0.0, "pinned": True},
    {"id": 7, "i": 7, "j": 8, "type": "Roof beam", "beta_deg": 0.0, "pinned": True},
    {"id": 8, "i": 8, "j": 5, "type": "Roof beam", "beta_deg": 0.0, "pinned": True},
    {"id": 9, "i": 1, "j": 5, "type": "Column", "beta_deg": 90.0, "pinned": False},
    {"id": 10, "i": 2, "j": 6, "type": "Column", "beta_deg": 90.0, "pinned": False},
    {"id": 11, "i": 3, "j": 7, "type": "Column", "beta_deg": 90.0, "pinned": False},
    {"id": 12, "i": 4, "j": 8, "type": "Column", "beta_deg": 90.0, "pinned": False},
]

# ─── Baseline Excel Ingestion Layers ──────────────────────────────────

def discover_xlsx(folder: Path, name_hint: str = "") -> Path:
    if not folder.exists():
        return folder / "placeholder.xlsx"
    candidates = sorted(folder.glob("*.xlsx"))
    if not candidates:
        return folder / "placeholder.xlsx"
    if name_hint:
        for p in candidates:
            if name_hint.lower() in p.stem.lower():
                return p
    return candidates[0]

def parse_unit_systems(ws) -> dict:
    systems = {}
    for row in ws.iter_rows(min_row=5, values_only=True):
        qty = row[0]
        if qty is None:
            continue
        systems[qty] = {
            "imperial": row[1], "metric": row[2],
            "internal": row[3], "note": row[4] if len(row) > 4 else None,
        }
    return systems

def parse_conversion_factors(ws) -> list[dict]:
    factors = []
    for row in ws.iter_rows(min_row=11, max_row=31, values_only=True):
        if row[0] is None:
            continue
        qty, imp_unit, met_unit, factor, reciprocal, derivation = row[:6]
        if isinstance(factor, str):
            continue
        factors.append({
            "quantity": qty, "imp_unit": imp_unit, "met_unit": met_unit,
            "factor": float(factor),
            "reciprocal": float(reciprocal) if reciprocal and not isinstance(reciprocal, str) else None,
            "derivation": derivation,
        })
    return factors

class UnitConverter:
    def __init__(self, systems: dict, factors: list[dict]):
        self._systems = systems
        self._factors = factors
        self._active = "metric"

    @property
    def active(self) -> str:
        return self._active

    @active.setter
    def active(self, system: str):
        if system not in ("metric", "imperial"):
            raise ValueError(f"Unknown unit system: {system}")
        self._active = system

    def unit(self, quantity: str) -> str:
        if quantity not in self._systems:
            return "m" if quantity == "Node coordinates" else "kN"
        return self._systems[quantity][self._active]

    def convert(self, value: float, from_unit: str, to_unit: str) -> float:
        if from_unit == to_unit:
            return value
        for f in self._factors:
            if f["imp_unit"] == from_unit and f["met_unit"] == to_unit:
                return value * f["factor"]
            if f["imp_unit"] == to_unit and f["met_unit"] == from_unit:
                if f["reciprocal"] is not None:
                    return value * f["reciprocal"]
        return value

def discover_material_xlsx(material_dir: Path) -> Path:
    combined = material_dir / "RISA_Materials_Library.xlsx"
    if combined.exists():
        return combined
    return material_dir / "placeholder.xlsx"

def discover_material_columns(headers: tuple) -> dict:
    col_map = {}
    for idx, h in enumerate(headers):
        if h is None:
            continue
        hs = str(h).strip()
        if hs == "Category": col_map["category"] = idx
        elif hs == "Label": col_map["label"] = idx
        elif "E" in hs and "[ksi]" in hs: col_map["imperial_E"] = idx
        elif "G" in hs and "[ksi]" in hs: col_map["imperial_G"] = idx
        elif hs == "Nu": col_map["Nu"] = idx
        elif "Therm" in hs and "1e-5" in hs: col_map["imperial_therm"] = idx
        elif "Density" in hs and "[k/ft" in hs: col_map["imperial_density"] = idx
        elif "Yield" in hs and "[ksi]" in hs: col_map["imperial_yield"] = idx
        elif hs.startswith("Fu") and "[ksi]" in hs: col_map["imperial_fu"] = idx
        elif "E" in hs and "[MPa]" in hs: col_map["metric_E"] = idx
        elif "G" in hs and "[MPa]" in hs: col_map["metric_G"] = idx
        elif "Therm" in hs and "1e-6" in hs: col_map["metric_therm"] = idx
        elif "Density" in hs and "[kN/m" in hs: col_map["metric_density"] = idx
        elif "Yield" in hs and "[MPa]" in hs: col_map["metric_yield"] = idx
        elif hs.startswith("Fu") and "[MPa]" in hs: col_map["metric_fu"] = idx
    return col_map

def _safe_float(val):
    if val is None or (isinstance(val, str) and val.strip() in ("Custom", "", "\u2013", "\u2014", "-")):
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None

def parse_all_materials(workbook_path: Path) -> dict[str, dict]:
    default_a992 = {
        "category": "Hot Rolled Steel", "label": "A992",
        "imperial": {"E": 29000.0, "G": 11200.0, "Nu": 0.3, "therm_coeff": 0.65, "density": 0.490, "yield_str": 50.0, "fu": 65.0},
        "metric": {"E": 200000.0, "G": 77200.0, "Nu": 0.3, "therm_coeff": 11.7, "density": 76.9729, "yield_str": 345.0, "fu": 450.0}
    }
    if not workbook_path.exists():
        return {"A992": default_a992}
    try:
        wb = openpyxl.load_workbook(workbook_path, data_only=True)
        ws = wb["All Materials"] if "All Materials" in wb.sheetnames else wb.active
        col_map, materials = {}, {}
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 3:
                col_map = discover_material_columns(row)
                continue
            if i < 4: continue
            label = row[col_map.get("label", 1)] if col_map else row[1]
            if not label: continue
            materials[str(label).strip()] = {
                "category": str(row[col_map.get("category", 0)] or "Steel"),
                "label": str(label).strip(),
                "imperial": {"E": 29000.0, "therm_coeff": 0.65, "density": 0.490, "yield_str": 50.0, "fu": 65.0, "Nu": 0.3},
                "metric": {
                    "E": _safe_float(row[col_map.get("metric_E", 9)]) or 200000.0,
                    "G": _safe_float(row[col_map.get("metric_G", 10)]) or 77200.0,
                    "Nu": _safe_float(row[col_map.get("Nu", 4)]) or 0.3,
                    "therm_coeff": _safe_float(row[col_map.get("metric_therm", 11)]) or 11.7,
                    "density": _safe_float(row[col_map.get("metric_density", 12)]) or 76.9729,
                    "yield_str": _safe_float(row[col_map.get("metric_yield", 14)]) or 345.0,
                    "fu": _safe_float(row[col_map.get("metric_fu", 15)]) or 450.0,
                }
            }
        wb.close()
        return materials if "A992" in materials else {"A992": default_a992}
    except Exception:
        return {"A992": default_a992}

class MaterialDatabase:
    def __init__(self, materials: dict[str, dict]):
        self._materials = materials

    @classmethod
    def from_excel(cls, material_dir: Path) -> "MaterialDatabase":
        xlsx_path = discover_material_xlsx(material_dir)
        return cls(parse_all_materials(xlsx_path))

    def get(self, label: str, system: str = "metric") -> dict | None:
        mat = self._materials.get(label, self._materials.get("A992"))
        return mat.get(system) if mat else None

    def category(self, label: str) -> str | None:
        mat = self._materials.get(label)
        return mat["category"] if mat else "Steel"

def parse_all_shapes(xlsx_path: Path) -> dict[str, dict]:
    default_shapes = {
        "W310X38.7": {
            "type": "Wide Flange",
            "metric": {"label": "W310X38.7", "W": 38.7, "A": 4940.0, "d": 310.0, "bf": 165.0, "tw": 5.8, "tf": 9.7, "Ix": 84.9e6, "Iy": 7.24e6},
            "imperial": {"label": "W12X26", "W": 26.0, "A": 7.65, "d": 12.2, "bf": 6.49, "tw": 0.23, "tf": 0.38, "Ix": 204.0, "Iy": 17.3}
        },
        "W250X49.1": {
            "type": "Wide Flange",
            "metric": {"label": "W250X49.1", "W": 49.1, "A": 6260.0, "d": 253.0, "bf": 203.0, "tw": 7.4, "tf": 11.0, "Ix": 70.8e6, "Iy": 15.3e6},
            "imperial": {"label": "W10X33", "W": 33.0, "A": 9.71, "d": 9.73, "bf": 7.96, "tw": 0.29, "tf": 0.435, "Ix": 170.0, "Iy": 36.6}
        }
    }
    if not xlsx_path.exists():
        return default_shapes
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True)
        ws = wb["Database v16.0"] if "Database v16.0" in wb.sheetnames else wb.active
        shapes = {}
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0: continue
            imp_lbl, met_lbl = row[IMP_COL["label"]], row[MET_COL["label"]]
            if not imp_lbl and not met_lbl: continue
            rec = {
                "type": row[0],
                "metric": {
                    "label": str(met_lbl) if met_lbl else str(imp_lbl),
                    "W": _safe_float(row[MET_COL["W"]]), "A": _safe_float(row[MET_COL["A"]]),
                    "d": _safe_float(row[MET_COL["d"]]), "bf": _safe_float(row[MET_COL["bf"]]),
                    "tw": _safe_float(row[MET_COL["tw"]]), "tf": _safe_float(row[MET_COL["tf"]]),
                    "Ix": (_safe_float(row[MET_COL["Ix"]]) or 0.0) * METRIC_SCALE["Ix"],
                    "Iy": (_safe_float(row[MET_COL["Iy"]]) or 0.0) * METRIC_SCALE["Iy"]
                },
                "imperial": {
                    "label": str(imp_lbl) if imp_lbl else str(met_lbl),
                    "W": _safe_float(row[IMP_COL["W"]]), "A": _safe_float(row[IMP_COL["A"]]),
                    "d": _safe_float(row[IMP_COL["d"]]), "bf": _safe_float(row[IMP_COL["bf"]]),
                    "tw": _safe_float(row[IMP_COL["tw"]]), "tf": _safe_float(row[IMP_COL["tf"]]),
                    "Ix": _safe_float(row[IMP_COL["Ix"]]), "Iy": _safe_float(row[IMP_COL["Iy"]])
                }
            }
            if imp_lbl: shapes[str(imp_lbl)] = rec
            if met_lbl: shapes[str(met_lbl)] = rec
        wb.close()
        return shapes if "W310X38.7" in shapes else default_shapes
    except Exception:
        return default_shapes

class ShapeDatabase:
    def __init__(self, shapes: dict[str, dict]):
        self._shapes = shapes

    @classmethod
    def from_excel(cls, shapes_dir: Path) -> "ShapeDatabase":
        xlsx = discover_xlsx(shapes_dir)
        return cls(parse_all_shapes(xlsx))

    def get(self, label: str, system: str = "metric") -> dict | None:
        shape = self._shapes.get(label)
        if not shape:
            return None
        return shape.get(system)

    def resolve_label(self, label: str, system: str) -> str:
        s = self._shapes.get(label)
        if s and s.get(system):
            return s[system].get("label", label)
        return label

class SolverState:
    def __init__(self):
        self.converter: UnitConverter = None
        self.beam_material_db: MaterialDatabase = None
        self.column_material_db: MaterialDatabase = None
        self.beam_shape_db: ShapeDatabase = None
        self.column_shape_db: ShapeDatabase = None
        self.active_unit_system: str = "metric"
        self.selected_beam_material: str = DEFAULT_BEAM_MATERIAL
        self.selected_column_material: str = DEFAULT_COLUMN_MATERIAL
        self.selected_beam_shape_imp: str = DEFAULT_BEAM_SHAPE_IMPERIAL
        self.selected_beam_shape_met: str = DEFAULT_BEAM_SHAPE_METRIC
        self.selected_column_shape_imp: str = DEFAULT_COLUMN_SHAPE_IMPERIAL
        self.selected_column_shape_met: str = DEFAULT_COLUMN_SHAPE_METRIC

    @property
    def selected_beam_shape(self) -> str:
        return self.selected_beam_shape_met if self.active_unit_system == "metric" else self.selected_beam_shape_imp

    @selected_beam_shape.setter
    def selected_beam_shape(self, label: str):
        if self.active_unit_system == "metric": self.selected_beam_shape_met = label
        else: self.selected_beam_shape_imp = label

    @property
    def selected_column_shape(self) -> str:
        return self.selected_column_shape_met if self.active_unit_system == "metric" else self.selected_column_shape_imp

    @selected_column_shape.setter
    def selected_column_shape(self, label: str):
        if self.active_unit_system == "metric": self.selected_column_shape_met = label
        else: self.selected_column_shape_imp = label

    def material_props_for_member(self, member_type: str) -> dict:
        db = self.beam_material_db if "beam" in member_type.lower() else self.column_material_db
        mat = self.selected_beam_material if "beam" in member_type.lower() else self.selected_column_material
        res = db.get(mat, self.active_unit_system)
        if not res:
            return {"E": 200000.0, "density": 76.9729, "therm_coeff": 11.7}
        return res

    def shape_props_for_member(self, member_type: str) -> dict:
        db = self.beam_shape_db if "beam" in member_type.lower() else self.column_shape_db
        lbl = self.selected_beam_shape if "beam" in member_type.lower() else self.selected_column_shape
        res = db.get(lbl, self.active_unit_system)
        if not res:
            if "beam" in member_type.lower():
                return {"W": 38.7, "A": 4940.0, "d": 310.0, "bf": 165.0, "tw": 5.8, "tf": 9.7}
            return {"W": 49.1, "A": 6260.0, "d": 253.0, "bf": 203.0, "tw": 7.4, "tf": 11.0}
        return res

    def switch_unit_system(self, system: str):
        self.active_unit_system = system
        self.converter.active = system

def member_length(member: dict) -> float:
    p1 = NODES[member["i"]]
    p2 = NODES[member["j"]]
    return math.sqrt(sum((b - a) ** 2 for a, b in zip(p1, p2)))

def local_axes(member: dict):
    p1, p2 = NODES[member["i"]], NODES[member["j"]]
    lx = [(b - a) / member_length(member) for a, b in zip(p1, p2)]
    gy = (0.0, 1.0, 0.0)
    proj = [gy[i] - sum(gy[k] * lx[k] for k in range(3)) * lx[i] for i in range(3)]
    mag = math.sqrt(sum(p ** 2 for p in proj))
    if mag < 1e-9:
        proj = (0.0, 0.0, 1.0)
        mag = 1.0
    ly0 = [p / mag for p in proj]
    lz0 = [lx[1]*ly0[2] - lx[2]*ly0[1], lx[2]*ly0[0] - lx[0]*ly0[2], lx[0]*ly0[1] - lx[1]*ly0[0]]
    beta = math.radians(member["beta_deg"])
    ly = [math.cos(beta)*ly0[i] + math.sin(beta)*lz0[i] for i in range(3)]
    lz = [-math.sin(beta)*ly0[i] + math.cos(beta)*lz0[i] for i in range(3)]
    return lx, ly, lz

# ─── Rev 3 Load Data Classes & Architecture ──────────────────────────

@dataclass
class NodalLoad:
    node_id: int
    fx: float = 0.0
    fy: float = 0.0
    fz: float = 0.0
    mx: float = 0.0
    my: float = 0.0
    mz: float = 0.0

@dataclass
class MemberDistributedLoad:
    member_id: int
    direction: str
    magnitude: float
    distribution_type: str = "uniform"

@dataclass
class MemberPointLoad:
    member_id: int
    location: float
    direction: str
    magnitude: float

@dataclass
class Diaphragm:
    id: int
    name: str
    master_node: int
    constrained_nodes: List[int]
    degrees_of_freedom: Tuple[str, ...] = ("UX", "UZ", "RY")

@dataclass
class TemperatureLoad:
    member_ids: List[int]
    temperature_change: float
    reference_temperature: float = 20.0
    alpha: float = 11.7e-6

@dataclass
class LoadCase:
    id: int
    name: str
    category: str
    description: str
    self_weight_factor: float = 0.0
    nodal_loads: List[NodalLoad] = field(default_factory=list)
    distributed_loads: List[MemberDistributedLoad] = field(default_factory=list)
    point_loads: List[MemberPointLoad] = field(default_factory=list)
    temperature_load: Optional[TemperatureLoad] = None

@dataclass
class LoadCombination:
    id: int
    name: str
    design_method: str
    factors: Dict[int, float]

# ─── Load Model Construction & Validation ─────────────────────────────

def build_rev3_load_cases(state: SolverState) -> Tuple[Dict[int, LoadCase], Diaphragm]:
    cases: Dict[int, LoadCase] = {}

    # LC1: Self-Weight (Global -Y, gamma * A * L)
    cases[1] = LoadCase(
        id=1, name="DEAD / SELF WEIGHT", category="Dead Load",
        description="Gravity self-weight of structural framing", self_weight_factor=1.0
    )

    # LC2: Roof Beam Dead Load (5 kN/m downward)
    lc2 = LoadCase(id=2, name="ROOF DEAD", category="Dead Load", description="Superimposed roof beam dead load")
    for mid in (5, 6, 7, 8):
        lc2.distributed_loads.append(MemberDistributedLoad(member_id=mid, direction="-Y", magnitude=5.0))
    cases[2] = lc2

    # LC3: Roof Beam Live Load (3 kN/m downward)
    lc3 = LoadCase(id=3, name="ROOF LIVE", category="Live Load", description="Roof live load")
    for mid in (5, 6, 7, 8):
        lc3.distributed_loads.append(MemberDistributedLoad(member_id=mid, direction="-Y", magnitude=3.0))
    cases[3] = lc3

    # LC4: Member Center Point Load (5 kN downward at midpoint L/2)
    lc4 = LoadCase(id=4, name="ROOF BEAM CENTER LOAD", category="Live Load", description="Beam center concentrated loads")
    for mid in (5, 6, 7, 8):
        lc4.point_loads.append(MemberPointLoad(member_id=mid, location=3.0, direction="-Y", magnitude=5.0))
    cases[4] = lc4

    # LC5: Wind X (10 kN total / 4 roof nodes = 2.5 kN each)
    lc5 = LoadCase(id=5, name="WIND X", category="Wind", description="Lateral wind in Global +X")
    for nid in (5, 6, 7, 8):
        lc5.nodal_loads.append(NodalLoad(node_id=nid, fx=2.5))
    cases[5] = lc5

    # LC6: Wind Z (10 kN total / 4 roof nodes = 2.5 kN each)
    lc6 = LoadCase(id=6, name="WIND Z", category="Wind", description="Lateral wind in Global +Z")
    for nid in (5, 6, 7, 8):
        lc6.nodal_loads.append(NodalLoad(node_id=nid, fz=2.5))
    cases[6] = lc6

    # LC7: Seismic X (15 kN total / 4 roof nodes = 3.75 kN each)
    lc7 = LoadCase(id=7, name="SEISMIC X", category="Seismic", description="Lateral earthquake in Global +X")
    for nid in (5, 6, 7, 8):
        lc7.nodal_loads.append(NodalLoad(node_id=nid, fx=3.75))
    cases[7] = lc7

    # LC8: Seismic Z (15 kN total / 4 roof nodes = 3.75 kN each)
    lc8 = LoadCase(id=8, name="SEISMIC Z", category="Seismic", description="Lateral earthquake in Global +Z")
    for nid in (5, 6, 7, 8):
        lc8.nodal_loads.append(NodalLoad(node_id=nid, fz=3.75))
    cases[8] = lc8

    # LC9: Temperature Load (+15 °C change on roof beams M5--M8)
    beam_mat = state.material_props_for_member("Roof beam")
    alpha_val = (beam_mat.get("therm_coeff", 11.7) or 11.7) * 1e-6
    lc9 = LoadCase(id=9, name="TEMPERATURE +15 degC", category="Temperature", description="Thermal expansion of roof beams")
    lc9.temperature_load = TemperatureLoad(
        member_ids=[5, 6, 7, 8], temperature_change=15.0, reference_temperature=20.0, alpha=alpha_val
    )
    cases[9] = lc9

    # Diaphragm: Roof level rigid planar behavior
    diaphragm = Diaphragm(
        id=1, name="Roof Diaphragm", master_node=5,
        constrained_nodes=[6, 7, 8], degrees_of_freedom=("UX", "UZ", "RY")
    )
    return cases, diaphragm

def build_load_combinations() -> List[LoadCombination]:
    combs = []
    # LRFD Combinations (NSCP 2015 Section 203.3)
    lrfd_defs = [
        ("1.4D", {1: 1.4, 2: 1.4}),
        ("1.2D + 1.6L", {1: 1.2, 2: 1.2, 3: 1.6, 4: 1.6}),
        ("1.2D + 1.0W_X + 1.0L", {1: 1.2, 2: 1.2, 3: 1.0, 4: 1.0, 5: 1.0}),
        ("1.2D - 1.0W_X + 1.0L", {1: 1.2, 2: 1.2, 3: 1.0, 4: 1.0, 5: -1.0}),
        ("1.2D + 1.0W_Z + 1.0L", {1: 1.2, 2: 1.2, 3: 1.0, 4: 1.0, 6: 1.0}),
        ("1.2D - 1.0W_Z + 1.0L", {1: 1.2, 2: 1.2, 3: 1.0, 4: 1.0, 6: -1.0}),
        ("1.2D + 1.0E_X + 1.0L", {1: 1.2, 2: 1.2, 3: 1.0, 4: 1.0, 7: 1.0}),
        ("1.2D - 1.0E_X + 1.0L", {1: 1.2, 2: 1.2, 3: 1.0, 4: 1.0, 7: -1.0}),
        ("1.2D + 1.0E_Z + 1.0L", {1: 1.2, 2: 1.2, 3: 1.0, 4: 1.0, 8: 1.0}),
        ("1.2D - 1.0E_Z + 1.0L", {1: 1.2, 2: 1.2, 3: 1.0, 4: 1.0, 8: -1.0}),
        ("0.9D + 1.0W_X", {1: 0.9, 2: 0.9, 5: 1.0}),
        ("0.9D - 1.0W_X", {1: 0.9, 2: 0.9, 5: -1.0}),
        ("0.9D + 1.0E_X", {1: 0.9, 2: 0.9, 7: 1.0}),
        ("0.9D - 1.0E_X", {1: 0.9, 2: 0.9, 7: -1.0}),
        ("1.2D + 1.2T + 1.6L", {1: 1.2, 2: 1.2, 3: 1.6, 4: 1.6, 9: 1.2}),
        ("1.2D + 1.0T + 1.0W_X", {1: 1.2, 2: 1.2, 5: 1.0, 9: 1.0}),
    ]
    for cid, (name, factors) in enumerate(lrfd_defs, start=1):
        combs.append(LoadCombination(id=cid, name=f"LRFD: {name}", design_method="LRFD", factors=factors))

    # ASD Combinations (NSCP 2015 Section 203.4)
    asd_defs = [
        ("Combination 13 - D", {1: 1.0, 2: 1.0, 4: 1.0}),
        ("D + L", {1: 1.0, 2: 1.0, 3: 1.0, 4: 1.0}),
        ("D + 0.6W_X", {1: 1.0, 2: 1.0, 5: 0.6}),
        ("D - 0.6W_X", {1: 1.0, 2: 1.0, 5: -0.6}),
        ("D + 0.6W_Z", {1: 1.0, 2: 1.0, 6: 0.6}),
        ("D - 0.6W_Z", {1: 1.0, 2: 1.0, 6: -0.6}),
        ("D + 0.7E_X", {1: 1.0, 2: 1.0, 7: 0.7}),
        ("D - 0.7E_X", {1: 1.0, 2: 1.0, 7: -0.7}),
        ("D + 0.7E_Z", {1: 1.0, 2: 1.0, 8: 0.7}),
        ("D - 0.7E_Z", {1: 1.0, 2: 1.0, 8: -0.7}),
        ("0.6D + 0.6W_X", {1: 0.6, 2: 0.6, 5: 0.6}),
        ("0.6D + 0.7E_X", {1: 0.6, 2: 0.6, 7: 0.7}),
        ("D + T", {1: 1.0, 2: 1.0, 9: 1.0}),
        ("D + 0.75L + 0.75T", {1: 1.0, 2: 1.0, 3: 0.75, 4: 0.75, 9: 0.75}),
    ]
    for idx, (name, factors) in enumerate(asd_defs, start=len(combs) + 1):
        combs.append(LoadCombination(id=idx, name=f"ASD: {name}", design_method="ASD", factors=factors))
    return combs

def compute_case_totals(lc: LoadCase, state: SolverState) -> Dict[str, float]:
    fx = fy = fz = 0.0

    # Self weight: gamma * A * L (gamma in kN/m^3, A in mm^2 -> 1e-6, L in m)
    if lc.self_weight_factor != 0.0:
        for m in MEMBERS:
            mat = state.material_props_for_member(m["type"])
            shp = state.shape_props_for_member(m["type"])
            gamma = mat.get("density", 76.9729) or 76.9729
            area_m2 = (shp.get("A", 4940.0) or 4940.0) * 1e-6
            L = member_length(m)
            w_total = gamma * area_m2 * L * lc.self_weight_factor
            fy -= w_total

    # Nodal loads
    for nl in lc.nodal_loads:
        fx += nl.fx
        fy += nl.fy
        fz += nl.fz

    # Distributed member loads
    for dl in lc.distributed_loads:
        m = next(mem for mem in MEMBERS if mem["id"] == dl.member_id)
        L = member_length(m)
        total_q = dl.magnitude * L
        if dl.direction == "-Y": fy -= total_q
        elif dl.direction == "+Y": fy += total_q
        elif dl.direction == "+X": fx += total_q
        elif dl.direction == "-X": fx -= total_q
        elif dl.direction == "+Z": fz += total_q
        elif dl.direction == "-Z": fz -= total_q

    # Point loads
    for pl in lc.point_loads:
        if pl.direction == "-Y": fy -= pl.magnitude
        elif pl.direction == "+Y": fy += pl.magnitude
        elif pl.direction == "+X": fx += pl.magnitude
        elif pl.direction == "-X": fx -= pl.magnitude
        elif pl.direction == "+Z": fz += pl.magnitude
        elif pl.direction == "-Z": fz -= pl.magnitude

    # Temperature loads generate 0.000 net structural force
    return {"Fx": fx, "Fy": fy, "Fz": fz, "Resultant": math.sqrt(fx**2 + fy**2 + fz**2)}

# ─── Enhanced 3D Load Visualization ───────────────────────────────────

def draw_load_glyphs(ax, lc: LoadCase, state: SolverState, show_bands: bool = True, scale_factor: float = 1.0):
    # Nodal load vectors
    for nl in lc.nodal_loads:
        p = NODES[nl.node_id]
        fx, fy, fz = nl.fx * scale_factor, nl.fy * scale_factor, nl.fz * scale_factor
        mag = math.sqrt(fx**2 + fy**2 + fz**2)
        if mag > 1e-6:
            vec = (fx / mag, fz / mag, fy / mag)
            ax.quiver(p[0], p[2], p[1], vec[0], vec[1], vec[2],
                      length=1.4, normalize=True, color="#d92525", linewidth=2.4, arrow_length_ratio=0.28)
            ax.text(p[0] + vec[0]*1.45, p[2] + vec[1]*1.45, p[1] + vec[2]*1.45,
                    f"{abs(mag):.2f} kN", color="#940505", fontsize=8.5, weight="bold")

    # Member point loads
    for pl in lc.point_loads:
        m = next(mem for mem in MEMBERS if mem["id"] == pl.member_id)
        p1, p2 = NODES[m["i"]], NODES[m["j"]]
        ratio = pl.location / member_length(m)
        mid = tuple(p1[k] + ratio * (p2[k] - p1[k]) for k in range(3))
        scaled_mag = pl.magnitude * scale_factor
        sign = -1.0 if "-" in pl.direction else 1.0
        vy = sign * 1.3
        # Vector draws pointing towards the load point
        ax.quiver(mid[0], mid[2], mid[1] - vy, 0, 0, vy,
                  length=1.3, normalize=True, color="#b00505", linewidth=2.4, arrow_length_ratio=0.3)
        ax.text(mid[0], mid[2], mid[1] - vy - 0.2,
                f"{abs(scaled_mag):.2f} kN", color="#b00505", fontsize=8.5, weight="bold")

    # Member distributed loads
    for dl in lc.distributed_loads:
        m = next(mem for mem in MEMBERS if mem["id"] == dl.member_id)
        p1, p2 = NODES[m["i"]], NODES[m["j"]]
        L = member_length(m)
        scaled_q = dl.magnitude * scale_factor
        num_arrows = max(5, int(L * 3))
        standoff = 0.12
        band_depth = 0.65
        sign = -1.0 if "-" in dl.direction else 1.0

        pts_on_beam = [tuple(p1[k] + (t / (num_arrows - 1)) * (p2[k] - p1[k]) for k in range(3)) for t in range(num_arrows)]

        # Arrows pointing to beam
        for pt in pts_on_beam:
            ax.quiver(pt[0], pt[2], pt[1] - sign * (band_depth + standoff), 0, 0, sign * band_depth,
                      length=band_depth, normalize=True, color="#d94b00", linewidth=1.2, arrow_length_ratio=0.25)

        # Polygon band
        if show_bands:
            top_pts = [(pt[0], pt[2], pt[1] - sign * (band_depth + standoff)) for pt in pts_on_beam]
            bot_pts = [(pt[0], pt[2], pt[1] - sign * standoff) for pt in reversed(pts_on_beam)]
            poly = Poly3DCollection([top_pts + bot_pts], facecolors="#d94b00", alpha=OPACITY_DISTRIBUTED_BAND)
            ax.add_collection3d(poly)
            # Intensity outline line
            ax.plot([top_pts[0][0], top_pts[-1][0]], [top_pts[0][1], top_pts[-1][1]], [top_pts[0][2], top_pts[-1][2]],
                    color="#9e3700", linewidth=1.5)

        mid_t = top_pts[len(top_pts) // 2]
        ax.text(mid_t[0], mid_t[1], mid_t[2] - sign * 0.15,
                f"{abs(scaled_q):.2f} kN/m", color="#9e3700", fontsize=8, weight="bold", ha="center")

    # Self-Weight visualization (Load Case 1)
    if lc.self_weight_factor != 0.0:
        for m in MEMBERS:
            p1, p2 = NODES[m["i"]], NODES[m["j"]]
            L = member_length(m)
            mat = state.material_props_for_member(m["type"])
            shp = state.shape_props_for_member(m["type"])
            gamma = mat.get("density", 76.9729) or 76.9729
            area_m2 = (shp.get("A", 4940.0) or 4940.0) * 1e-6
            w_line = gamma * area_m2 * lc.self_weight_factor * scale_factor
            is_column = m["type"] == "Column"

            if is_column:
                # Column: skip degenerate polygon, render discrete axial downward arrows
                for t in (0.25, 0.5, 0.75):
                    cz = p1[1] + t * (p2[1] - p1[1])
                    ax.quiver(p1[0], p1[2], cz + 0.3, 0, 0, -0.3,
                              length=0.3, normalize=True, color="#4a4a4a", linewidth=1.2, arrow_length_ratio=0.3)
            else:
                num_arrows = max(5, int(L * 2.5))
                standoff, band_depth = 0.10, 0.40
                pts = [tuple(p1[k] + (t / (num_arrows - 1)) * (p2[k] - p1[k]) for k in range(3)) for t in range(num_arrows)]
                for pt in pts:
                    ax.quiver(pt[0], pt[2], pt[1] + band_depth + standoff, 0, 0, -band_depth,
                      length=band_depth, normalize=True, color="#595959", linewidth=1.0, arrow_length_ratio=0.25)
                if show_bands:
                    top_pts = [(pt[0], pt[2], pt[1] + band_depth + standoff) for pt in pts]
                    bot_pts = [(pt[0], pt[2], pt[1] + standoff) for pt in reversed(pts)]
                    poly = Poly3DCollection([top_pts + bot_pts], facecolors="#595959", alpha=OPACITY_SELF_WEIGHT_BAND)
                    ax.add_collection3d(poly)
                    ax.plot([top_pts[0][0], top_pts[-1][0]], [top_pts[0][1], top_pts[-1][1]], [top_pts[0][2], top_pts[-1][2]],
                            color="#333333", linewidth=1.2)
                mid_t = top_pts[len(top_pts) // 2]
                ax.text(mid_t[0], mid_t[1], mid_t[2] + 0.12,
                        f"w_sw={w_line:.3f} kN/m", color="#333333", fontsize=7.5, ha="center")

    # Temperature Load visualization (Load Case 9)
    if lc.temperature_load:
        tload = lc.temperature_load
        for mid in tload.member_ids:
            m = next(mem for mem in MEMBERS if mem["id"] == mid)
            p1, p2 = NODES[m["i"]], NODES[m["j"]]
            cx = (p1[0] + p2[0]) / 2.0
            cy = (p1[1] + p2[1]) / 2.0
            cz = (p1[2] + p2[2]) / 2.0
            # Expanding opposed arrows along member
            dx, dz = (p2[0] - p1[0]) / 6.0, (p2[2] - p1[2]) / 6.0
            ax.quiver(cx, cz, cy, dx, dz, 0, length=0.6, normalize=True, color="#c91010", linewidth=1.8)
            ax.quiver(cx, cz, cy, -dx, -dz, 0, length=0.6, normalize=True, color="#c91010", linewidth=1.8)
            ax.text(cx, cz, cy + 0.25, f"+{tload.temperature_change:.1f} \u00b0C (\u03b5={tload.alpha*tload.temperature_change:.2e})",
                    color="#c91010", fontsize=8, weight="bold", ha="center")

def draw_diaphragm_overlay(ax, diaphragm: Diaphragm):
    m_pt = NODES[diaphragm.master_node]
    ax.scatter(m_pt[0], m_pt[2], m_pt[1], color="#9900cc", s=140, marker="D", edgecolors="black",
               depthshade=False, label="Diaphragm Master (N5)")
    for c_nid in diaphragm.constrained_nodes:
        c_pt = NODES[c_nid]
        ax.plot([m_pt[0], c_pt[0]], [m_pt[2], c_pt[2]], [m_pt[1], c_pt[1]],
                color="#b82ee6", linestyle=":", linewidth=1.8)
    roof_poly = [[NODES[5][0], NODES[5][2], NODES[5][1]],
                 [NODES[6][0], NODES[6][2], NODES[6][1]],
                 [NODES[7][0], NODES[7][2], NODES[7][1]],
                 [NODES[8][0], NODES[8][2], NODES[8][1]]]
    ax.add_collection3d(Poly3DCollection([roof_poly], facecolors="#e6b8ff", alpha=0.15))

def draw_complete_scene(ax, state: SolverState, current_item, cases: Dict[int, LoadCase],
                        diaphragm: Diaphragm, show_grid: bool, show_bands: bool):
    ax.clear()
    # Members
    for member in MEMBERS:
        start, end = NODES[member["i"]], NODES[member["j"]]
        color = "#08775d" if member["type"] == "Column" else "#2454d8"
        ax.plot([start[0], end[0]], [start[2], end[2]], [start[1], end[1]],
                color=color, linewidth=2.5, alpha=0.85)
        mid = tuple((a + b) / 2 for a, b in zip(start, end))
        ax.text(mid[0], mid[2], mid[1] + 0.08, f"M{member['id']}", color="#123a8c", fontsize=7.5, weight="bold")

    # Nodes & Pinned supports
    for node, (x, y, z) in NODES.items():
        is_sup = node in SUPPORTS
        ax.scatter(x, z, y, color="#c40000" if is_sup else "#ff5959", s=60, edgecolors="black", linewidths=0.6)
        ax.text(x, z, y + 0.18, f"N{node}", fontsize=8, weight="bold")
        if is_sup:
            base_y, hw = y - 0.45, 0.28
            apex = (x, z, y)
            base = [(x-hw, z-hw, base_y), (x+hw, z-hw, base_y), (x+hw, z+hw, base_y), (x-hw, z+hw, base_y)]
            faces = [[apex, base[0], base[1]], [apex, base[1], base[2]], [apex, base[2], base[3]], [apex, base[3], base[0]], base]
            ax.add_collection3d(Poly3DCollection(faces, facecolors="#666666", edgecolors="#222222", linewidths=0.6, alpha=0.8))

    # Diaphragm overlay
    draw_diaphragm_overlay(ax, diaphragm)

    # Render Active Selection (Load Case vs Combination)
    title_suffix = ""
    if isinstance(current_item, LoadCase):
        draw_load_glyphs(ax, current_item, state, show_bands=show_bands, scale_factor=1.0)
        tots = compute_case_totals(current_item, state)
        title_suffix = f"Case {current_item.id}: {current_item.name} | Resultant = {tots['Resultant']:.2f} kN"
    elif isinstance(current_item, LoadCombination):
        comb = current_item
        for cid, factor in comb.factors.items():
            base_case = cases[cid]
            draw_load_glyphs(ax, base_case, state, show_bands=show_bands, scale_factor=factor)
        title_suffix = f"Combination: {comb.name} ({comb.design_method})"

    ax.set_title(f"6m Cube Structural Model - {REVISION}\n{title_suffix}", fontsize=11, weight="bold", pad=15)
    ax.set_xlabel("X (m) - Lateral", fontsize=9, weight="bold", labelpad=6)
    ax.set_ylabel("Z (m) - Lateral", fontsize=9, weight="bold", labelpad=6)
    ax.set_zlabel("Y (m) - Vertical", fontsize=9, weight="bold", labelpad=6)
    ax.set_xlim(-1, 7)
    ax.set_ylim(-1, 7)
    ax.set_zlim(-1, 7)
    ax.set_box_aspect((1, 1, 1))
    ax.view_init(elev=20, azim=-55)

    # Clean Grid Toggle handling without Matplotlib alpha bug
    if show_grid:
        ax.grid(True, alpha=0.3)
        ax.xaxis.pane.set_visible(True)
        ax.yaxis.pane.set_visible(True)
        ax.zaxis.pane.set_visible(True)
    else:
        ax.grid(False)
        ax.xaxis.pane.set_visible(False)
        ax.yaxis.pane.set_visible(False)
        ax.zaxis.pane.set_visible(False)

# ─── Verification Report Generation ───────────────────────────────────

def generate_verification_report(state: SolverState, cases: Dict[int, LoadCase],
                                 combs: List[LoadCombination], diaphragm: Diaphragm,
                                 output_path: Path):
    lines = [
        "=" * 82,
        f"        STRUCTURAL FRAME SOLVER - {REVISION} LOAD AUDIT REPORT",
        "=" * 82,
        "Notice: Rev 3 is a LOADS, CONSTRAINTS, AND COMBINATIONS FRAMEWORK ONLY.",
        "Stiffness analysis, support reactions, internal member forces, and displacements",
        "remain out of scope for a subsequent solver revision.",
        "",
        "SECTION 1: STRUCTURAL MODEL GEOMETRY",
        f"  Cube dimensions: 6.000 m x 6.000 m x 6.000 m | Nodes: {len(NODES)} | Members: {len(MEMBERS)}",
        f"  Supports: Nodes 1-4 Pinned (UX, UY, UZ restrained; RX, RY, RZ free)",
        f"  Beams: {state.selected_beam_shape} ({state.selected_beam_material}) | Columns: {state.selected_column_shape} ({state.selected_column_material})",
        "",
        "SECTION 2: RIGID DIAPHRAGM DEFINITION",
        f"  Diaphragm Name        : {diaphragm.name}",
        f"  Elevation             : Y = 6.000 m",
        f"  Master Reference Node : Node {diaphragm.master_node}",
        f"  Constrained Nodes     : {', '.join(f'Node {n}' for n in diaphragm.constrained_nodes)}",
        f"  Constrained DOFs      : {', '.join(diaphragm.degrees_of_freedom)}",
        f"  Released DOFs         : UY, RX, RZ",
        "",
        "SECTION 3: LOAD CASES 1 TO 8 SUMMARY & EQUILIBRIUM AUDIT",
        f"  {'Case':<6} {'Name':<26} {'Type':<12} {'Intended (kN)':<15} {'Computed (kN)':<15} {'Error (kN)':<12}",
        "-" * 86
    ]

    expected_totals = {1: 29.815, 2: 120.000, 3: 72.000, 4: 20.000, 5: 10.000, 6: 10.000, 7: 15.000, 8: 15.000}
    for cid in range(1, 9):
        lc = cases[cid]
        tots = compute_case_totals(lc, state)
        exp = expected_totals[cid]
        comp = tots["Resultant"]
        err = abs(comp - exp)
        lines.append(f"  LC{cid:<4} {lc.name:<26} {lc.category:<12} {exp:<15.4f} {comp:<15.4f} {err:<12.6f}")

    # LC9 Temperature Verification
    lc9 = cases[9]
    tload = lc9.temperature_load
    beam_shp = state.shape_props_for_member("Roof beam")
    beam_mat = state.material_props_for_member("Roof beam")
    A_mm2 = beam_shp.get("A", 4940.0) or 4940.0
    E_mpa = beam_mat.get("E", 200000.0) or 200000.0
    EA_kN = (E_mpa * 1e3) * (A_mm2 * 1e-6)
    eps_T = tload.alpha * tload.temperature_change
    dL_mm = eps_T * 6000.0
    thrust_kN = EA_kN * eps_T

    lines.extend([
        "",
        "SECTION 4: TEMPERATURE LOAD VERIFICATION (LOAD CASE 9)",
        f"  Applied Temperature Change   : +{tload.temperature_change:.2f} deg C",
        f"  Affected Elements            : Members {tload.member_ids} (Roof beams)",
        f"  Thermal Expansion Coeff alpha: {tload.alpha:.4e} /deg C (read from RISA library)",
        f"  Calculated Thermal Strain    : {eps_T:.4e} (alpha * dT)",
        f"  Free Expansion (dL = eps*L)  : {dL_mm:.4f} mm on 6.0 m beam",
        f"  Beam Axial Rigidity (EA)     : {EA_kN:,.1f} kN",
        f"  Fully Restrained Thrust (N)  : {thrust_kN:.3f} kN (Compression)",
        f"  Net Structural External Force: 0.000000 kN (Self-equilibrating element pair)",
        f"  Partially Restrained State   : Indeterminate (Stiffness solver out of scope in Rev 3)",
        "",
        "SECTION 5: LOAD COMBINATION DEFINITIONS (NSCP 2015)",
        f"  Total Combinations Generated : {len(combs)} (16 LRFD, 14 ASD; 4 Temperature Combinations)",
        f"  {'ID':<5} {'Designation':<30} {'Method':<8} {'Factors Breakdown'}",
        "-" * 86
    ])

    for c in combs:
        f_str = ", ".join(f"{cases[cid].name}: {fac}" for cid, fac in c.factors.items())
        lines.append(f"  {c.id:<5} {c.name:<30} {c.design_method:<8} {f_str}")

    report_text = "\n".join(lines)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(report_text)
    print(f"Audit report written successfully to {output_path}")

def run_headless_verification(state: SolverState):
    cases, diaphragm = build_rev3_load_cases(state)
    combs = build_load_combinations()
    report_file = Path("cube_rev3_verification_report.txt")
    generate_verification_report(state, cases, combs, diaphragm, report_file)

    # Render PNG figures for individual cases and sample combinations
    fig = plt.figure(figsize=(10, 8))
    ax = fig.add_subplot(111, projection="3d")
    for cid in range(1, 10):
        draw_complete_scene(ax, state, cases[cid], cases, diaphragm, show_grid=True, show_bands=True)
        img_name = f"rev3_load_case_{cid}.png"
        fig.savefig(img_name, dpi=130, bbox_inches="tight")
        print(f"Saved: {img_name}")

    draw_complete_scene(ax, state, combs[1], cases, diaphragm, show_grid=True, show_bands=True)
    fig.savefig("rev3_combination_1.png", dpi=130, bbox_inches="tight")
    print("Saved: rev3_combination_1.png")

    # Render Combination 13 specifically
    comb_13 = next((c for c in combs if "Combination 13" in c.name or c.id == 17), combs[16])
    draw_complete_scene(ax, state, comb_13, cases, diaphragm, show_grid=True, show_bands=True)
    fig.savefig("rev3_combination_13.png", dpi=130, bbox_inches="tight")
    print("Saved: rev3_combination_13.png")
    plt.close(fig)

# ─── Interactive GUI Application ─────────────────────────────────────

def launch_gui(state: SolverState):
    cases, diaphragm = build_rev3_load_cases(state)
    combs = build_load_combinations()

    fig = plt.figure(figsize=(18, 10))
    ax_3d = fig.add_axes((0.02, 0.08, 0.65, 0.84), projection="3d")

    ui_state = {
        "current_item": cases[1],
        "show_grid": True,
        "show_bands": True,
    }

    def redraw():
        draw_complete_scene(ax_3d, state, ui_state["current_item"], cases, diaphragm,
                            ui_state["show_grid"], ui_state["show_bands"])
        fig.canvas.draw_idle()

    # Sidebar Controls
    ax_grid_btn = fig.add_axes([0.72, 0.90, 0.11, 0.04])
    btn_grid = Button(ax_grid_btn, "Toggle Grid", color="#e6f2ff", hovercolor="#cce6ff")
    def on_grid_click(event):
        ui_state["show_grid"] = not ui_state["show_grid"]
        redraw()
    btn_grid.on_clicked(on_grid_click)

    ax_band_btn = fig.add_axes([0.84, 0.90, 0.13, 0.04])
    btn_band = Button(ax_band_btn, "Toggle Load Bands", color="#e6f2ff", hovercolor="#cce6ff")
    def on_band_click(event):
        ui_state["show_bands"] = not ui_state["show_bands"]
        redraw()
    btn_band.on_clicked(on_band_click)

    # Radio Selector for Cases & Combinations
    items_map = {}
    for cid, lc in cases.items():
        items_map[f"LC{cid}: {lc.name[:18]}"] = lc
    for c in combs:
        items_map[f"{c.name[:24]}"] = c

    ax_selector = fig.add_axes([0.72, 0.15, 0.25, 0.70])
    ax_selector.set_facecolor("#f9f9fc")
    ax_selector.set_title("Select Load Case or Combination", fontsize=9, weight="bold")
    radio_selector = RadioButtons(ax_selector, tuple(items_map.keys()), active=0)

    def on_select(label):
        ui_state["current_item"] = items_map[label]
        redraw()
    radio_selector.on_clicked(on_select)

    redraw()
    plt.show()

# ─── Initialization ───────────────────────────────────────────────────

def initialize() -> SolverState:
    script_dir = Path(__file__).parent
    units_xlsx = discover_xlsx(script_dir / "units")
    if units_xlsx.exists():
        try:
            wb = openpyxl.load_workbook(units_xlsx, data_only=True)
            conv = UnitConverter(parse_unit_systems(wb["Unit Systems"]), parse_conversion_factors(wb["Conversion Factors"]))
            wb.close()
        except Exception:
            conv = UnitConverter({}, [])
    else:
        conv = UnitConverter({}, [])

    state = SolverState()
    state.converter = conv
    state.beam_material_db = MaterialDatabase.from_excel(script_dir / "Material")
    state.column_material_db = MaterialDatabase.from_excel(script_dir / "Material")
    state.beam_shape_db = ShapeDatabase.from_excel(script_dir / "Member Size")
    state.column_shape_db = ShapeDatabase.from_excel(script_dir / "Member Size")
    return state

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Cube Solver Rev 3 Loads Framework")
    parser.add_argument("--verify-loads", action="store_true", help="Run headless load verification and produce audit report and images")
    args = parser.parse_args()

    state = initialize()
    if args.verify_loads:
        run_headless_verification(state)
    else:
        launch_gui(state)