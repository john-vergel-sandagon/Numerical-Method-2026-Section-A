"""Rev. 3 - Navigable 3-D structural model with Excel-driven units, materials, and member sizes.

Global coordinate system: X and Z are lateral; Y is vertical.  Each node has
six global degrees of freedom: UX, UY, UZ, RX, RY, and RZ.  Nodes 1--4 have
pinned supports (translations restrained, rotations free).

Unit system, material database, and section properties are read from Excel
files at startup.  The Excel files are the source of truth.

Run: python Rev.3_sandagon.py
This opens the navigable model and writes structural_model_Rev_3.xlsx.
"""

from __future__ import annotations

import math
import sys
from pathlib import Path
from typing import Iterable

import matplotlib.pyplot as plt
from matplotlib.widgets import RadioButtons, TextBox
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required: pip install openpyxl")


REVISION = "Rev. 3"
SIDE_LENGTH_M = 6.0
GLOBAL_DOFS = ("UX", "UY", "UZ", "RX", "RY", "RZ")

DEFAULT_BEAM_SHAPE_IMPERIAL = "W12X26"
DEFAULT_BEAM_SHAPE_METRIC = "W310X38.7"
DEFAULT_COLUMN_SHAPE_IMPERIAL = "W12X26"
DEFAULT_COLUMN_SHAPE_METRIC = "W310X38.7"
DEFAULT_BEAM_MATERIAL = "A992"
DEFAULT_COLUMN_MATERIAL = "A992"

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


# ─── Unit System Layer ────────────────────────────────────────────────

def discover_xlsx(folder: Path, name_hint: str = "") -> Path:
    candidates = sorted(folder.glob("*.xlsx"))
    if not candidates:
        raise FileNotFoundError(f"No .xlsx file found in {folder}")
    if name_hint:
        for p in candidates:
            if name_hint.lower() in p.stem.lower():
                return p
        raise FileNotFoundError(
            f"No .xlsx file matching '{name_hint}' in {folder}. "
            f"Found: {[p.name for p in candidates]}"
        )
    return candidates[0]


def parse_unit_systems(ws) -> dict:
    systems = {}
    for row in ws.iter_rows(min_row=5, values_only=True):
        quantity = row[0]
        if quantity is None:
            continue
        systems[quantity] = {
            "imperial": row[1],
            "metric": row[2],
            "internal": row[3],
            "note": row[4] if len(row) > 4 else None,
        }
    return systems


def parse_conversion_factors(ws) -> list[dict]:
    factors = []
    for row in ws.iter_rows(min_row=11, max_row=31, values_only=True):
        if row[0] is None:
            continue
        quantity, imp_unit, met_unit, factor, reciprocal, derivation = row[:6]
        if isinstance(factor, str):
            continue
        factors.append({
            "quantity": quantity,
            "imp_unit": imp_unit,
            "met_unit": met_unit,
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
        raise ValueError(f"No conversion path: {from_unit} -> {to_unit}")


# ─── Material Layer ───────────────────────────────────────────────────

def discover_material_xlsx(material_dir: Path) -> Path:
    combined = material_dir / "RISA_Materials_Library.xlsx"
    if combined.exists():
        return combined
    imp = material_dir / "RISA_Materials_Library_Imperial.xlsx"
    met = material_dir / "RISA_Materials_Library_Metric.xlsx"
    if imp.exists() or met.exists():
        return material_dir
    raise FileNotFoundError(
        f"No material database found in {material_dir}. "
        f"Expected 'RISA_Materials_Library.xlsx' or separate Imperial/Metric files."
    )


def discover_material_columns(headers: tuple) -> dict:
    col_map = {}
    for idx, h in enumerate(headers):
        if h is None:
            continue
        h = str(h).strip()
        if h == "Category":
            col_map["category"] = idx
        elif h == "Label":
            col_map["label"] = idx
        elif "E" in h and "[ksi]" in h and "imperial_E" not in col_map:
            col_map["imperial_E"] = idx
        elif "G" in h and "[ksi]" in h and "imperial_G" not in col_map:
            col_map["imperial_G"] = idx
        elif h == "Nu" and "Nu" not in col_map:
            col_map["Nu"] = idx
        elif "Therm" in h and "1e-5" in h:
            col_map["imperial_therm"] = idx
        elif "Density" in h and "[k/ft" in h:
            col_map["imperial_density"] = idx
        elif "Yield" in h and "[ksi]" in h:
            col_map["imperial_yield"] = idx
        elif h.startswith("Fu") and "[ksi]" in h:
            col_map["imperial_fu"] = idx
        elif "E" in h and "[MPa]" in h and "metric_E" not in col_map:
            col_map["metric_E"] = idx
        elif "G" in h and "[MPa]" in h and "metric_G" not in col_map:
            col_map["metric_G"] = idx
        elif "Therm" in h and "1e-6" in h:
            col_map["metric_therm"] = idx
        elif "Density" in h and "[kN/m" in h:
            col_map["metric_density"] = idx
        elif "Mass Density" in h:
            col_map["metric_mass_density"] = idx
        elif "Yield" in h and "[MPa]" in h:
            col_map["metric_yield"] = idx
        elif h.startswith("Fu") and "[MPa]" in h:
            col_map["metric_fu"] = idx
    return col_map


def _safe_float(val):
    if val is None:
        return None
    if isinstance(val, str) and val.strip() in ("Custom", "", "\u2013", "\u2014"):
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def parse_material_row(row: tuple, col_map: dict) -> dict | None:
    label = row[col_map.get("label", 1)]
    category = row[col_map.get("category", 0)]
    if label is None or category is None:
        return None
    cat_str = str(category).strip()
    if "Count" in cat_str or "count" in cat_str:
        return None

    def g(key):
        idx = col_map.get(key)
        return _safe_float(row[idx]) if idx is not None and idx < len(row) else None

    return {
        "category": cat_str,
        "label": str(label).strip(),
        "imperial": {
            "E": g("imperial_E"), "G": g("imperial_G"), "Nu": g("Nu"),
            "therm_coeff": g("imperial_therm"), "density": g("imperial_density"),
            "yield_str": g("imperial_yield"), "fu": g("imperial_fu"),
        },
        "metric": {
            "E": g("metric_E"), "G": g("metric_G"), "Nu": g("Nu"),
            "therm_coeff": g("metric_therm"), "density": g("metric_density"),
            "mass_density": g("metric_mass_density"),
            "yield_str": g("metric_yield"), "fu": g("metric_fu"),
        },
    }


def parse_all_materials(workbook_path: Path) -> dict[str, dict]:
    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    if "All Materials" in wb.sheetnames:
        ws = wb["All Materials"]
        headers = None
        materials = {}
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 3:
                headers = row
                col_map = discover_material_columns(headers)
                continue
            if i < 4 or i > 44:
                continue
            rec = parse_material_row(row, col_map)
            if rec is not None:
                materials[rec["label"]] = rec
        wb.close()
        return materials
    wb.close()
    return _parse_two_file_fallback(workbook_path.parent)


def _parse_two_file_fallback(base_dir: Path) -> dict[str, dict]:
    imp_path = base_dir / "RISA_Materials_Library_Imperial.xlsx"
    met_path = base_dir / "RISA_Materials_Library_Metric.xlsx"

    def _parse_file(path):
        if not path.exists():
            return {}
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["All Materials"]
        headers = None
        result = {}
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 3:
                headers = row
                col_map = discover_material_columns(headers)
                continue
            if i < 4 or i > 44:
                continue
            rec = parse_material_row(row, col_map)
            if rec is not None:
                result[rec["label"]] = rec
        wb.close()
        return result

    imp_mats = _parse_file(imp_path)
    met_mats = _parse_file(met_path)
    merged = dict(imp_mats)
    for label, met_rec in met_mats.items():
        if label in merged:
            merged[label]["metric"] = met_rec["metric"]
        else:
            merged[label] = met_rec
    return merged


class MaterialDatabase:
    def __init__(self, materials: dict[str, dict]):
        self._materials = materials

    @classmethod
    def from_excel(cls, material_dir: Path) -> "MaterialDatabase":
        xlsx_path = discover_material_xlsx(material_dir)
        materials = parse_all_materials(xlsx_path)
        return cls(materials)

    def get(self, label: str, system: str = "metric") -> dict | None:
        mat = self._materials.get(label)
        if mat is None:
            return None
        return mat.get(system)

    def category(self, label: str) -> str | None:
        m = self._materials.get(label)
        return m["category"] if m else None

    def labels(self) -> list[str]:
        return sorted(self._materials.keys())

    def categories(self) -> list[str]:
        seen, result = set(), []
        for m in self._materials.values():
            c = m["category"]
            if c not in seen:
                seen.add(c)
                result.append(c)
        return result

    def labels_for_category(self, category: str) -> list[str]:
        return sorted(l for l, m in self._materials.items() if m["category"] == category)


# ─── Shape Layer ──────────────────────────────────────────────────────

def parse_all_shapes(xlsx_path: Path) -> dict[str, dict]:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    ws = wb["Database v16.0"]
    shapes = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        shape_type = row[0]
        imp_label = row[IMP_COL["label"]]
        met_label = row[MET_COL["label"]]

        def g(col_map, prop):
            idx = col_map[prop]
            val = row[idx]
            if val is None or (isinstance(val, str) and val in ("\u2013", "\u2014", "-")):
                return None
            try:
                v = float(val)
                if idx >= 84:
                    v *= METRIC_SCALE.get(prop, 1)
                return v
            except (ValueError, TypeError):
                return None

        record = {
            "type": shape_type,
            "imperial": {
                "label": str(imp_label) if imp_label else None,
                "W": g(IMP_COL, "W"), "A": g(IMP_COL, "A"), "d": g(IMP_COL, "d"),
                "bf": g(IMP_COL, "bf"), "tw": g(IMP_COL, "tw"), "tf": g(IMP_COL, "tf"),
                "Ix": g(IMP_COL, "Ix"), "Zx": g(IMP_COL, "Zx"), "Sx": g(IMP_COL, "Sx"),
                "rx": g(IMP_COL, "rx"), "Iy": g(IMP_COL, "Iy"),
                "J": g(IMP_COL, "J"), "Cw": g(IMP_COL, "Cw"),
            },
            "metric": {
                "label": str(met_label) if met_label else None,
                "W": g(MET_COL, "W"), "A": g(MET_COL, "A"), "d": g(MET_COL, "d"),
                "bf": g(MET_COL, "bf"), "tw": g(MET_COL, "tw"), "tf": g(MET_COL, "tf"),
                "Ix": g(MET_COL, "Ix"), "Zx": g(MET_COL, "Zx"), "Sx": g(MET_COL, "Sx"),
                "rx": g(MET_COL, "rx"), "Iy": g(MET_COL, "Iy"),
                "J": g(MET_COL, "J"), "Cw": g(MET_COL, "Cw"),
            },
        }
        if imp_label:
            shapes[str(imp_label)] = record
        if met_label and met_label != imp_label:
            shapes[str(met_label)] = record
    wb.close()
    return shapes


class ShapeDatabase:
    def __init__(self, shapes: dict[str, dict]):
        self._shapes = shapes

    @classmethod
    def from_excel(cls, shapes_dir: Path) -> "ShapeDatabase":
        xlsx = discover_xlsx(shapes_dir)
        shapes = parse_all_shapes(xlsx)
        return cls(shapes)

    def get(self, label: str, system: str = "metric") -> dict | None:
        shape = self._shapes.get(label)
        if shape is None:
            return None
        return shape.get(system)

    def resolve_label(self, label: str, system: str) -> str | None:
        shape = self._shapes.get(label)
        if shape is None:
            return None
        rec = shape.get(system)
        if rec:
            return rec.get("label")
        return label

    def labels(self) -> list[str]:
        return sorted(self._shapes.keys())


# ─── Solver State ─────────────────────────────────────────────────────

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

    def _label_for(self, imp: str, met: str) -> str:
        return met if self.active_unit_system == "metric" else imp

    @property
    def selected_beam_shape(self) -> str:
        return self._label_for(self.selected_beam_shape_imp, self.selected_beam_shape_met)

    @selected_beam_shape.setter
    def selected_beam_shape(self, label: str):
        if self.active_unit_system == "metric":
            self.selected_beam_shape_met = label
        else:
            self.selected_beam_shape_imp = label

    @property
    def selected_column_shape(self) -> str:
        return self._label_for(self.selected_column_shape_imp, self.selected_column_shape_met)

    @selected_column_shape.setter
    def selected_column_shape(self, label: str):
        if self.active_unit_system == "metric":
            self.selected_column_shape_met = label
        else:
            self.selected_column_shape_imp = label

    @property
    def active_beam_material_props(self) -> dict | None:
        return self.beam_material_db.get(self.selected_beam_material, self.active_unit_system)

    @property
    def active_column_material_props(self) -> dict | None:
        return self.column_material_db.get(self.selected_column_material, self.active_unit_system)

    def material_for_member(self, member_type: str) -> str:
        if member_type in ("Base beam", "Roof beam"):
            return self.selected_beam_material
        return self.selected_column_material

    def material_props_for_member(self, member_type: str) -> dict | None:
        db = self.beam_material_db if member_type in ("Base beam", "Roof beam") else self.column_material_db
        return db.get(self.material_for_member(member_type), self.active_unit_system)

    @property
    def active_beam_shape_props(self) -> dict | None:
        return self.beam_shape_db.get(self.selected_beam_shape, self.active_unit_system)

    @property
    def active_column_shape_props(self) -> dict | None:
        return self.column_shape_db.get(self.selected_column_shape, self.active_unit_system)

    def beam_shape_label_for(self, system: str) -> str:
        return self.selected_beam_shape_met if system == "metric" else self.selected_beam_shape_imp

    def column_shape_label_for(self, system: str) -> str:
        return self.selected_column_shape_met if system == "metric" else self.selected_column_shape_imp

    def shape_for_member(self, member_type: str) -> str:
        if member_type in ("Base beam", "Roof beam"):
            return self.selected_beam_shape
        return self.selected_column_shape

    def shape_props_for_member(self, member_type: str) -> dict | None:
        db = self.beam_shape_db if member_type in ("Base beam", "Roof beam") else self.column_shape_db
        return db.get(self.shape_for_member(member_type), self.active_unit_system)

    def switch_unit_system(self, system: str):
        self.active_unit_system = system
        self.converter.active = system


# ─── Geometry Functions ───────────────────────────────────────────────

def vector_subtract(a: tuple[float, float, float], b: tuple[float, float, float]):
    return tuple(a_i - b_i for a_i, b_i in zip(a, b))


def dot(a: Iterable[float], b: Iterable[float]) -> float:
    return sum(a_i * b_i for a_i, b_i in zip(a, b))


def cross(a: tuple[float, float, float], b: tuple[float, float, float]):
    return (a[1] * b[2] - a[2] * b[1],
            a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0])


def normalize(v: tuple[float, float, float]):
    magnitude = math.sqrt(dot(v, v))
    if magnitude == 0:
        raise ValueError("A zero-length member does not have local axes.")
    return tuple(value / magnitude for value in v)


def local_axes(member: dict):
    local_x = normalize(vector_subtract(NODES[member["j"]], NODES[member["i"]]))
    global_y = (0.0, 1.0, 0.0)
    projection = tuple(global_y[index] - dot(global_y, local_x) * local_x[index]
                       for index in range(3))
    if math.sqrt(dot(projection, projection)) < 1e-9:
        projection = (0.0, 0.0, 1.0)
    local_y_0 = normalize(projection)
    local_z_0 = normalize(cross(local_x, local_y_0))
    beta = math.radians(member["beta_deg"])
    local_y = tuple(math.cos(beta) * local_y_0[index] + math.sin(beta) * local_z_0[index]
                    for index in range(3))
    local_z = tuple(-math.sin(beta) * local_y_0[index] + math.cos(beta) * local_z_0[index]
                    for index in range(3))
    return local_x, local_y, local_z


def node_dof_state(node: int):
    restraints = SUPPORTS.get(node, {})
    return {dof: "Restrained" if restraints.get(dof, False) else "Free" for dof in GLOBAL_DOFS}


def member_length(member: dict) -> float:
    delta = vector_subtract(NODES[member["j"]], NODES[member["i"]])
    return math.sqrt(dot(delta, delta))


# ─── Self-Weight ──────────────────────────────────────────────────────

def self_weight_from_shape(state: SolverState, shape) -> float | None:
    if shape is None or shape.get("W") is None:
        return None
    w = shape["W"]
    if state.active_unit_system == "metric":
        return w * 9.80665 / 1000.0
    else:
        return w / 1000.0


def compute_self_weight(state: SolverState, member_type: str) -> float | None:
    return self_weight_from_shape(state, state.shape_props_for_member(member_type))


# ─── Display Formatting ───────────────────────────────────────────────

def fmt_val(value, unit: str) -> str:
    if value is None:
        return "N/A"
    if "mm" in unit and ("4" in unit or "\u2074" in unit) and "3" not in unit:
        if abs(value) >= 1e6:
            return f"{value / 1e6:.2f} \u00d7 10\u2076 mm\u2074"
    elif "mm" in unit and ("3" in unit or "\u00b3" in unit):
        if abs(value) >= 1e3:
            return f"{value / 1e3:.1f} \u00d7 10\u00b3 mm\u00b3"
    elif "mm" in unit and ("6" in unit or "\u2076" in unit):
        if abs(value) >= 1e9:
            return f"{value / 1e9:.2f} \u00d7 10\u2079 mm\u2076"
    return f"{value:.4g} {unit}"


def build_info_text(state: SolverState) -> str:
    u = state.active_unit_system
    conv = state.converter
    beam_mat = state.active_beam_material_props
    column_mat = state.active_column_material_props
    beam_shape = state.active_beam_shape_props
    column_shape = state.active_column_shape_props

    length_unit = conv.unit("Node coordinates")
    stress_unit = conv.unit("Stress, modulus").split("(")[0].strip() if "(" in conv.unit("Stress, modulus") else conv.unit("Stress, modulus")
    section_unit = conv.unit("Section dimensions")
    density_unit = conv.unit("Unit weight (density)")
    dist_unit = conv.unit("Distributed load")

    side = SIDE_LENGTH_M if u == "metric" else conv.convert(SIDE_LENGTH_M, "m", "ft")

    def p(label, val, default="N/A"):
        return f"    {label:<28}{val if val is not None else default}"

    def pf(label, val, unit_str):
        return f"    {label:<28}{fmt_val(val, unit_str)}"

    def material_block(title, label, mat, material_db):
        block = [p("Selected", label)]
        block.append(p("Source", "RISA_Materials_Library.xlsx"))
        if mat:
            mat_cat = material_db.category(label)
            block.append(p("Category", mat_cat if mat_cat else "N/A"))
            block.append(pf("E", mat.get("E"), stress_unit))
            block.append(pf("G", mat.get("G"), stress_unit))
            block.append(p("Nu", f"{mat['Nu']:.3f}" if mat.get("Nu") is not None else "N/A"))
            block.append(pf("Density", mat.get("density"), density_unit))
            block.append(pf("Yield", mat.get("yield_str"), stress_unit))
            block.append(pf("Fu", mat.get("fu"), stress_unit))
        lines.append(f"    {title}")
        lines.extend(block)

    def size_block(title, label, shape):
        block = [p("Selected", label)]
        block.append(p("Source", "aisc-shapes-database-v160-2.xlsx"))
        if shape:
            block.append(pf("A", shape.get("A"), conv.unit("Area")))
            block.append(pf("d", shape.get("d"), section_unit))
            block.append(pf("bf", shape.get("bf"), section_unit))
            block.append(pf("tw", shape.get("tw"), section_unit))
            block.append(pf("tf", shape.get("tf"), section_unit))
            block.append(pf("Ix", shape.get("Ix"), conv.unit("Moment of inertia")))
            block.append(pf("Iy", shape.get("Iy"), conv.unit("Moment of inertia")))
            sw = self_weight_from_shape(state, shape)
            if sw is not None:
                block.append(p("Self-weight", f"{sw:.4f} {dist_unit}"))
        lines.append(f"    {title}")
        lines.extend(block)

    lines = [f"MODEL DATA - {REVISION}", ""]
    lines.append("Geometry")
    lines.append(p("Cube edge", f"{side:.3f} {length_unit}"))
    lines.append(p("Nodes", "8"))
    lines.append(p("Members", "12"))
    lines.append(p("Vertical axis", "global Y"))
    lines.append("")

    lines.append("Unit System")
    lines.append(p("Active", "Standard Metric" if u == "metric" else "Imperial"))
    lines.append(p("Source", "Units_Imperial_Metric.xlsx"))
    lines.append("")

    lines.append("Material")
    material_block("Beams (base + roof)", state.selected_beam_material, beam_mat, state.beam_material_db)
    lines.append("")
    material_block("Columns", state.selected_column_material, column_mat, state.column_material_db)
    lines.append("")

    lines.append("Member Size")
    size_block("Beams (base + roof)", state.selected_beam_shape, beam_shape)
    lines.append("")
    size_block("Columns", state.selected_column_shape, column_shape)
    lines.append("")

    lines.append("Supports")
    lines.append(p("Type", "pinned"))
    lines.append(p("Nodes", "1, 2, 3, 4"))
    lines.append(p("Restrained", "UX, UY, UZ"))
    lines.append(p("Released", "RX, RY, RZ"))
    lines.append("")

    lines.append("Degrees of freedom")
    lines.append(p("DOF per node", "6"))
    lines.append(p("Total DOF", "48"))
    lines.append(p("Restrained DOF", "12"))
    lines.append(p("Active DOF", "36"))
    lines.append(p("Numbering", "(node - 1) x 6 + 1...6"))
    lines.append("")

    lines.append("Beta angles")
    lines.append(p("Base Beam", "0 deg"))
    lines.append(p("Roof Beam", "0 deg"))
    lines.append(p("Column", "90 deg"))
    lines.append("")

    lines.append("Local axes")
    lines.append("    local x   start node i -> end node j")
    lines.append("    local y   in the vertical plane, when possible")
    lines.append("    local z   completes the right-handed set")
    lines.append("    Vertical members: local z parallel to global Z")

    return "\n".join(lines)


# ─── Visualization ────────────────────────────────────────────────────

def draw_model(ax_3d, state: SolverState):
    beam_color, column_color = "#2454d8", "#08775d"

    length_unit = state.converter.unit("Node coordinates")
    is_metric = state.active_unit_system == "metric"
    axis_label_x = f"X ({length_unit}) - lateral"
    axis_label_y = f"Y ({length_unit}) - vertical"
    axis_label_z = f"Z ({length_unit}) - lateral"

    for member in MEMBERS:
        start, end = NODES[member["i"]], NODES[member["j"]]
        if not is_metric:
            start = tuple(state.converter.convert(v, "m", "ft") for v in start)
            end = tuple(state.converter.convert(v, "m", "ft") for v in end)

        member_color = column_color if member["type"] == "Column" else beam_color
        ax_3d.plot([start[0], end[0]], [start[2], end[2]], [start[1], end[1]],
                   color=member_color, linewidth=2.7,
                   label="Column" if member["id"] == 9 else "Beam" if member["id"] == 1 else None)

        mid = tuple((a + b) / 2 for a, b in zip(start, end))
        lx, ly, lz = local_axes(member)
        for axis, color, lbl in zip((lx, ly, lz),
                                    ("#dd2929", "#22aa31", "#9361d1"),
                                    ("Local x axis", "Local y axis", "Local z axis")):
            ax_3d.quiver(mid[0], mid[2], mid[1], axis[0], axis[2], axis[1],
                         length=0.65, normalize=True, color=color, arrow_length_ratio=0.24,
                         linewidth=1.5, label=lbl if member["id"] == 1 else None)
        beta_label = f" (\u03b2={member['beta_deg']:.0f}\u00b0)" if member["beta_deg"] else ""
        ax_3d.text(mid[0], mid[2], mid[1] + 0.10,
                   f"M{member['id']}{beta_label}", color="#123a8c", fontsize=8, weight="bold")

    for node, (x, y, z) in NODES.items():
        if not is_metric:
            x, y, z = state.converter.convert(x, "m", "ft"), state.converter.convert(y, "m", "ft"), state.converter.convert(z, "m", "ft")
        supported = node in SUPPORTS
        color = "#c40000" if supported else "#ff5959"
        ax_3d.scatter(x, z, y, color=color, edgecolors="black", linewidths=0.7, s=76,
                      depthshade=False,
                      label="Supported node (pinned)" if node == 1 else "Free node" if node == 5 else None)
        dof_start = (node - 1) * 6 + 1
        ax_3d.text(x, z, y + 0.20, f"N{node}\nDOF {dof_start}-{dof_start + 5}",
                   fontsize=8.5, weight="bold", color="#111111")
        if supported:
            draw_pinned_support(ax_3d, x, y, z)

    ax_3d.scatter(0, 0, 0, marker="*", color="#12952c", edgecolors="black", s=115,
                  depthshade=False, label="Origin (0, 0, 0)")

    unit_label = "Standard Metric" if is_metric else "Imperial"
    ax_3d.set_title(f"6m x 6m x 6m Cube - Structural Model, {REVISION}\n"
                    f"Pinned supports at nodes 1-4 | Unit system: {unit_label}",
                    fontsize=13, weight="bold", pad=25)
    ax_3d.set_xlabel(axis_label_x, fontsize=11, weight="bold", labelpad=10)
    ax_3d.set_ylabel(axis_label_z, fontsize=11, weight="bold", labelpad=10)
    ax_3d.set_zlabel(axis_label_y, fontsize=11, weight="bold", labelpad=10)
    limit = 7 if is_metric else state.converter.convert(7, "m", "ft")
    ax_3d.set_xlim(-1, limit)
    ax_3d.set_ylim(-1, limit)
    ax_3d.set_zlim(-1, limit)
    ax_3d.set_box_aspect((1, 1, 1))
    ax_3d.view_init(elev=17, azim=-58)
    ax_3d.legend(loc="upper left", bbox_to_anchor=(-0.08, 0.98), fontsize=8, framealpha=0.94)
    ax_3d.grid(True)


def plot_structure(state: SolverState):
    fig = plt.figure(figsize=(20, 12))

    ax_3d = fig.add_axes((0.02, 0.08, 0.50, 0.84), projection="3d")
    draw_model(ax_3d, state)

    info_text = build_info_text(state)
    info_text_obj = fig.text(0.54, 0.92, info_text, va="top", ha="left", fontsize=8.5,
                             family="monospace",
                             bbox={"boxstyle": "round,pad=0.6", "facecolor": "#f5f6fc",
                                   "edgecolor": "#153c77", "linewidth": 1.25})

    widget_elements = setup_widgets(fig, ax_3d, state, info_text_obj)
    plt.show()


def draw_pinned_support(ax, x: float, y: float, z: float):
    base_y, half_width = y - 0.55, 0.33
    apex = (x, z, y)
    base = [(x - half_width, z - half_width, base_y), (x + half_width, z - half_width, base_y),
            (x + half_width, z + half_width, base_y), (x - half_width, z + half_width, base_y)]
    faces = [[apex, base[0], base[1]], [apex, base[1], base[2]],
             [apex, base[2], base[3]], [apex, base[3], base[0]], base]
    ax.add_collection3d(Poly3DCollection(faces, facecolors="#555555", edgecolors="#222222",
                                         linewidths=0.8, alpha=0.9))


def setup_widgets(fig, ax_3d, state: SolverState, info_text_obj):
    ax_unit = fig.add_axes([0.78, 0.76, 0.20, 0.10])
    ax_unit.set_facecolor("#f0f0f5")
    ax_unit.set_title("Unit System", fontsize=9, weight="bold", pad=4)
    radio_unit = RadioButtons(ax_unit, ("Standard Metric", "Imperial"),
                              active=0 if state.active_unit_system == "metric" else 1)

    ax_beam_mat = fig.add_axes([0.78, 0.68, 0.20, 0.04])
    ax_beam_mat.set_facecolor("#f0f0f5")
    ax_beam_mat.set_title("Beam Material", fontsize=9, weight="bold", pad=4)
    box_beam_mat = TextBox(ax_beam_mat, "", initial=state.selected_beam_material)

    ax_column_mat = fig.add_axes([0.78, 0.59, 0.20, 0.04])
    ax_column_mat.set_facecolor("#f0f0f5")
    ax_column_mat.set_title("Column Material", fontsize=9, weight="bold", pad=4)
    box_column_mat = TextBox(ax_column_mat, "", initial=state.selected_column_material)

    ax_beam = fig.add_axes([0.78, 0.50, 0.20, 0.04])
    ax_beam.set_facecolor("#f0f0f5")
    ax_beam.set_title("Beam Size", fontsize=9, weight="bold", pad=4)
    box_beam = TextBox(ax_beam, "", initial=state.selected_beam_shape)

    ax_column = fig.add_axes([0.78, 0.41, 0.20, 0.04])
    ax_column.set_facecolor("#f0f0f5")
    ax_column.set_title("Column Size", fontsize=9, weight="bold", pad=4)
    box_column = TextBox(ax_column, "", initial=state.selected_column_shape)

    error_ax = fig.add_axes([0.78, 0.36, 0.20, 0.03])
    error_ax.set_facecolor("#fff0f0")
    error_ax.set_axis_off()
    error_text = error_ax.text(0.0, 0.5, "", fontsize=7.5, color="red",
                               va="center", transform=error_ax.transAxes)

    def update_info():
        info_text_obj.set_text(build_info_text(state))
        fig.canvas.draw_idle()

    def on_unit_change(label):
        sys_name = "metric" if "Metric" in label else "imperial"
        state.switch_unit_system(sys_name)
        box_beam.set_val(state.selected_beam_shape)
        box_column.set_val(state.selected_column_shape)
        error_text.set_text("")
        ax_3d.clear()
        draw_model(ax_3d, state)
        update_info()

    def _resolve_material(text, material_db):
        label = text.strip()
        if material_db.get(label, state.active_unit_system) is not None:
            return label
        return None

    def on_beam_mat_submit(text):
        label = _resolve_material(text, state.beam_material_db)
        if label is None:
            error_text.set_text(f"Material '{text.strip()}' not found")
            fig.canvas.draw_idle()
            return
        error_text.set_text("")
        state.selected_beam_material = label
        update_info()

    def on_column_mat_submit(text):
        label = _resolve_material(text, state.column_material_db)
        if label is None:
            error_text.set_text(f"Material '{text.strip()}' not found")
            fig.canvas.draw_idle()
            return
        error_text.set_text("")
        state.selected_column_material = label
        update_info()

    def _resolve_shape(text, shape_db):
        label = text.strip()
        props = shape_db.get(label, state.active_unit_system)
        if props is None:
            resolved = shape_db.resolve_label(label, state.active_unit_system)
            if resolved:
                props = shape_db.get(resolved, state.active_unit_system)
                if props:
                    label = resolved
        return label if props is not None else None

    def on_beam_submit(text):
        label = _resolve_shape(text, state.beam_shape_db)
        if label is None:
            error_text.set_text(f"Shape '{text.strip()}' not found")
            fig.canvas.draw_idle()
            return
        error_text.set_text("")
        state.selected_beam_shape = label
        update_info()

    def on_column_submit(text):
        label = _resolve_shape(text, state.column_shape_db)
        if label is None:
            error_text.set_text(f"Shape '{text.strip()}' not found")
            fig.canvas.draw_idle()
            return
        error_text.set_text("")
        state.selected_column_shape = label
        update_info()

    radio_unit.on_clicked(on_unit_change)
    box_beam_mat.on_submit(on_beam_mat_submit)
    box_column_mat.on_submit(on_column_mat_submit)
    box_beam.on_submit(on_beam_submit)
    box_column.on_submit(on_column_submit)

    return {"unit": radio_unit, "beam_mat": box_beam_mat, "column_mat": box_column_mat,
            "beam": box_beam, "column": box_column, "error": error_text}


# ─── Excel Export ─────────────────────────────────────────────────────

def column_letter(index: int) -> str:
    result = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def write_excel_output(filename: Path, state: SolverState):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = openpyxl.Workbook()
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="17365D", end_color="17365D", fill_type="solid")
    green_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    orange_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    wrap_align = Alignment(wrap_text=True)

    def style_header(ws, row_num=1):
        for cell in ws[row_num]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

    def auto_width(ws):
        for col_cells in ws.columns:
            max_len = 0
            col_letter = column_letter(col_cells[0].column - 1)
            for cell in col_cells:
                try:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 3, 28)

    u = state.active_unit_system
    conv = state.converter
    beam_mat = state.active_beam_material_props
    column_mat = state.active_column_material_props
    beam_shape = state.active_beam_shape_props
    column_shape = state.active_column_shape_props

    ws_sum = wb.active
    ws_sum.title = "Model Summary"
    summary_data = [
        ["Item", "Value"],
        ["Revision", REVISION],
        ["Unit system", "Standard Metric" if u == "metric" else "Imperial"],
        ["Cube edge length", f"{SIDE_LENGTH_M} m (internal)"],
        ["Number of nodes", 8], ["Number of members", 12],
        ["Supported nodes", 4], ["Support type", "Pinned"],
        ["DOF per node", 6], ["Total DOF", 48],
        ["Restrained DOF", 12], ["Active DOF (equations)", 36],
        ["Global vertical axis", "Y"], ["Global lateral axes", "X and Z"],
        ["Selected beam material", state.selected_beam_material],
        ["Beam material category", state.beam_material_db.category(state.selected_beam_material) or ""],
        ["Selected column material", state.selected_column_material],
        ["Column material category", state.column_material_db.category(state.selected_column_material) or ""],
        ["Selected beam size", f"{state.selected_beam_shape} (base + roof)"],
        ["Selected column size", state.selected_column_shape],
    ]
    for row in summary_data:
        ws_sum.append(row)
    style_header(ws_sum)
    auto_width(ws_sum)

    ws_nodes = wb.create_sheet("Nodes")
    ws_nodes.append(["Node", f"X ({conv.unit('Node coordinates')})",
                     f"Y ({conv.unit('Node coordinates')})",
                     f"Z ({conv.unit('Node coordinates')})", "Support"])
    for node, (x, y, z) in NODES.items():
        xu = conv.convert(x, "m", conv.unit("Node coordinates")) if u == "imperial" else x
        yu = conv.convert(y, "m", conv.unit("Node coordinates")) if u == "imperial" else y
        zu = conv.convert(z, "m", conv.unit("Node coordinates")) if u == "imperial" else z
        ws_nodes.append([node, round(xu, 4), round(yu, 4), round(zu, 4),
                         "Pinned" if node in SUPPORTS else "Free"])
    style_header(ws_nodes)
    auto_width(ws_nodes)

    ws_mem = wb.create_sheet("Member Incidences")
    ws_mem.append(["Member", "Node i", "Node j", "Type",
                   f"Length ({conv.unit('Node coordinates')})", "Beta (deg)",
                   "Size", "Material"])
    for member in MEMBERS:
        length_m = member_length(member)
        if u == "imperial":
            length_u = conv.convert(length_m, "m", conv.unit("Node coordinates"))
        else:
            length_u = length_m
        ws_mem.append([member["id"], member["i"], member["j"], member["type"],
                       round(length_u, 4), member["beta_deg"],
                       state.shape_for_member(member["type"]),
                       state.material_for_member(member["type"])])
    style_header(ws_mem)
    auto_width(ws_mem)

    ws_sup = wb.create_sheet("Supports")
    ws_sup.append(["Node", "Support Type", *GLOBAL_DOFS, "Restraint Code"])
    for node in (1, 2, 3, 4):
        states = node_dof_state(node)
        code = "".join("1" if states[d] == "Restrained" else "0" for d in GLOBAL_DOFS)
        ws_sup.append([node, "Pinned", *(states[d] for d in GLOBAL_DOFS), code])
    style_header(ws_sup)
    auto_width(ws_sup)

    def write_material_sheet(wb, sheet_name, mat):
        ws_mat = wb.create_sheet(sheet_name)
        ws_mat.append(["Property", f"Value ({conv.unit('Stress, modulus')})", "Unit"])
        if mat:
            for key, label in [("E", "Modulus of Elasticity"), ("G", "Shear Modulus"),
                               ("Nu", "Poisson's Ratio"), ("therm_coeff", "Thermal Coeff."),
                               ("density", "Unit Weight"), ("yield_str", "Yield Strength"),
                               ("fu", "Ultimate Strength")]:
                val = mat.get(key)
                if key in ("E", "G", "yield_str", "fu"):
                    unit_str = conv.unit("Stress, modulus")
                elif key == "density":
                    unit_str = conv.unit("Unit weight (density)")
                elif key == "therm_coeff":
                    unit_str = "1e-5/°F" if u == "imperial" else "1e-6/°C"
                else:
                    unit_str = ""
                ws_mat.append([label, val, unit_str])
        style_header(ws_mat)
        auto_width(ws_mat)

    write_material_sheet(wb, "Beam Material Properties", beam_mat)
    write_material_sheet(wb, "Column Material Properties", column_mat)

    def write_section_sheet(wb, sheet_name, shape):
        ws_sec = wb.create_sheet(sheet_name)
        ws_sec.append(["Property", "Value", "Unit"])
        if shape:
            sec_unit = conv.unit("Section dimensions")
            props_list = [
                ("A", "Area", conv.unit("Area")),
                ("d", "Depth", sec_unit),
                ("bf", "Flange Width", sec_unit),
                ("tw", "Web Thickness", sec_unit),
                ("tf", "Flange Thickness", sec_unit),
                ("Ix", "Moment of Inertia X", conv.unit("Moment of inertia")),
                ("Iy", "Moment of Inertia Y", conv.unit("Moment of inertia")),
                ("Zx", "Plastic Modulus X", "mm\u00b3" if u == "metric" else "in\u00b3"),
                ("Sx", "Elastic Modulus X", "mm\u00b3" if u == "metric" else "in\u00b3"),
                ("J", "Torsion Constant", conv.unit("Moment of inertia")),
            ]
            for key, label, unit_str in props_list:
                ws_sec.append([label, shape.get(key), unit_str])
        style_header(ws_sec)
        auto_width(ws_sec)

    write_section_sheet(wb, "Beam Section Properties", beam_shape)
    write_section_sheet(wb, "Column Section Properties", column_shape)

    ws_units = wb.create_sheet("Unit System")
    ws_units.append(["Quantity", "Imperial", "Standard Metric", "Active", "Internal"])
    for qty, info in conv._systems.items():
        ws_units.append([qty, info["imperial"], info["metric"],
                         info[u], info.get("internal", "")])
    style_header(ws_units)
    auto_width(ws_units)

    wb.save(filename)


# ─── Initialization ───────────────────────────────────────────────────

def initialize() -> SolverState:
    script_dir = Path(__file__).parent

    units_xlsx = discover_xlsx(script_dir / "units")
    wb = openpyxl.load_workbook(units_xlsx, data_only=True)
    unit_systems = parse_unit_systems(wb["Unit Systems"])
    conversion_factors = parse_conversion_factors(wb["Conversion Factors"])
    wb.close()
    converter = UnitConverter(unit_systems, conversion_factors)

    state = SolverState()
    state.converter = converter
    state.beam_material_db = MaterialDatabase.from_excel(script_dir / "Material")
    state.column_material_db = MaterialDatabase.from_excel(script_dir / "Material")
    state.beam_shape_db = ShapeDatabase.from_excel(script_dir / "Member Size")
    state.column_shape_db = ShapeDatabase.from_excel(script_dir / "Member Size")
    return state


if __name__ == "__main__":
    state = initialize()
    output_file = Path(__file__).with_name("structural_model_Rev_3_detailed.xlsx")
    write_excel_output(output_file, state)
    print(f"Wrote Excel model data: {output_file}")
    plot_structure(state)
