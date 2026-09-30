"""Export the boundary walls of the stair and smoke filter rooms with their fire related data (read only)"""

__title__ = "Muri REI\nscale"
__author__ = "Luca Rosati"
__guida__ = "Estrae in CSV (Excel) i muri perimetrali dei locali vano scala e filtro fumo (codice VS/FF nel numero o nome del locale): tipo, descrizione del tipo, spessore, stratigrafia con materiali, parametri di resistenza al fuoco e compartimentazione, vincoli di base e di testa, fase, porte ospitate, lunghezza sul confine. Per ogni muro cerca nei modelli collegati il muro strutturale parallelo dietro (addossato o con intercapedine fino a 500 mm): tipo, spessore, materiale, intercapedine, copertura. Segnala tratti senza struttura dietro, confini che non sono muri e locali con nome e codice incoerenti. Non modifica il modello."

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# Extraction only: the REI evaluation is done afterwards on the CSV (decision 29/09/2026).
# Rooms: function code of the number PIANO.FUNZIONE.NN (PGI 3.11.12) VS / FF, or the name starting with
# "Vano scala" / "Filtro fumo"; both are reported because they can disagree (e.g. "Filtro fumo L02A.VS.17").
# Boundary: Room.GetBoundarySegments (finish faces), one row per room and bounding element with the summed
# segment length. Elements of linked models (room bounding links) are read from the link document.
# Structure behind (29/09/2026: linings stand against the 300 mm structural walls of the linked ST model or
# with a 240 mm cavity): for each host wall with a straight axis, walls of every loaded link that are
# parallel, on the side opposite to the room (Room.IsPointInRoom), with a face gap 0..MAX_GAP mm, a plan
# overlap and a vertical overlap; plan test in bmc_rules.wall_behind (self-checked).
# Coverage = plan overlap / wall length.

import codecs

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

import bmc_rules as R
import bmc_utils as U

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

KINDS = {"VS": "vano scala", "FF": "filtro fumo"}
# fire parameters of the PIR (PRE type, VVF instance)
FIRE_PARAMS = [
    (name, level)
    for name, level in R.PRE + R.VVF
    if "fuoco" in name or "Compartimento" in name
]
P_DOOR_FIRE = "PRE_TE_Resistenza al fuoco"
MAX_GAP, MIN_OVERLAP, MIN_Z_OVERLAP = 500, 300, 300  # mm
COVERED = 90  # % of the wall length with structure behind
MAX_WALL = 1000  # mm: thicker IFC shapes are not a single straight wall
BEHIND = [
    "Muro dietro (modello collegato: tipo)",
    "Muro dietro spessore mm",
    "Muro dietro materiale strutturale",
    "Intercapedine mm",
    "Copertura muro dietro %",
]

COLUMNS = (
    [
        "Livello locale",
        "Numero locale",
        "Nome locale",
        "Fase locale",
        "Funzione da codice",
        "Funzione da nome",
        "Modello",
        "Id elemento",
        "Categoria",
        "Tipo",
        "Contrassegno tipo",
        "Descrizione tipo",
        "Tipologia muro",
        "Funzione muro",
        "Spessore mm",
        "Stratigrafia (funzione, materiale, mm)",
    ]
    + [name for name, _l in FIRE_PARAMS]
    + [
        "Fire Rating (Revit)",
        "Vincolo base",
        "Offset base mm",
        "Vincolo superiore",
        "Offset superiore mm",
        "Altezza mm",
        "Testa collegata",
        "Fase creazione",
        "Fase demolizione",
        "Lunghezza sul confine mm",
        "Porte ospitate (tipo, %s)" % P_DOOR_FIRE,
    ]
    + BEHIND
)


def text(element, bip):
    p = element.get_Parameter(bip)
    if p is None:
        return ""
    return p.AsValueString() or p.AsString() or ""


def mm(element, bip):
    p = element.get_Parameter(bip)
    return "%.0f" % U.to_mm(p.AsDouble()) if p is not None and p.HasValue else ""


def value(src_doc, element, name, level):
    """Parameter value for the CSV: the value, or (assente) / (vuoto) / nd; Yes/No as Si/No."""
    _target, state, v = U.param_state_any(src_doc, element, name, level)
    if state == "missing":
        return "(assente)"
    if state == "empty":
        return "(vuoto)"
    if "_SN_" in name and state == "value":
        return "Si" if v == 1 else "No"
    return v


def phase_name(src_doc, phase_id):
    if phase_id is None or phase_id == ElementId.InvalidElementId:
        return ""
    return U.ename(src_doc.GetElement(phase_id))


def room_kinds(room):
    """(function from the number code, function from the name) as VS / FF / ''."""
    parts = (room.Number or "").split(".")
    code = parts[1] if len(parts) == 3 and parts[1] in KINDS else ""
    name = (text(room, BuiltInParameter.ROOM_NAME) or "").strip().lower()
    by_name = [k for k, label in KINDS.items() if name.startswith(label)]
    return code, by_name[0] if by_name else ""


def layers(src_doc, wall_type):
    cs = wall_type.GetCompoundStructure() if isinstance(wall_type, WallType) else None
    if cs is None:
        return ""
    rows = []
    for layer in cs.GetLayers():
        material = src_doc.GetElement(layer.MaterialId)
        rows.append(
            "%s %s %.0f"
            % (
                layer.Function,
                U.ename(material) if material is not None else "(nessun materiale)",
                U.to_mm(layer.Width),
            )
        )
    return " | ".join(rows)


def hosted_doors():
    """Host wall id -> ['door type (fire resistance)'] (host model only)."""
    result = {}
    for door in U.instances(doc, [BuiltInCategory.OST_Doors]):
        host = getattr(door, "Host", None)
        if host is None:
            continue
        door_type = U.type_of(doc, door)
        result.setdefault(U.idv(host.Id), []).append(
            "%s (%s)"
            % (
                U.ename(door_type) if door_type is not None else "?",
                value(doc, door, P_DOOR_FIRE, "T"),
            )
        )
    return result


def resolve(segment):
    """(model label, document, element) of a boundary segment; None when the element is gone."""
    if segment.LinkElementId != ElementId.InvalidElementId:
        link = doc.GetElement(segment.ElementId)
        link_doc = link.GetLinkDocument() if link is not None else None
        if link_doc is None:
            return None
        return U.ename(link), link_doc, link_doc.GetElement(segment.LinkElementId)
    return "host", doc, doc.GetElement(segment.ElementId)


def element_row(model, src_doc, element, doors):
    """Columns from 'Modello' on, without the boundary length."""
    category = element.Category.Name if element.Category is not None else ""
    base = [model, str(U.idv(element.Id)), category]
    if not isinstance(element, Wall):
        # 6 room columns + model, id, category, name
        return base + [U.ename(element)] + [""] * (len(COLUMNS) - 10)
    wall_type = element.WallType
    fire = [value(src_doc, element, name, level) for name, level in FIRE_PARAMS]
    return (
        base
        + [
            U.ename(wall_type),
            U.type_mark(wall_type),
            value(src_doc, element, R.P_DESCRIPTION, "T"),
            str(wall_type.Kind),
            text(wall_type, BuiltInParameter.FUNCTION_PARAM),
            "%.0f" % U.to_mm(element.Width),
            layers(src_doc, wall_type),
        ]
        + fire
        + [
            text(wall_type, BuiltInParameter.FIRE_RATING),
            text(element, BuiltInParameter.WALL_BASE_CONSTRAINT),
            mm(element, BuiltInParameter.WALL_BASE_OFFSET),
            text(element, BuiltInParameter.WALL_HEIGHT_TYPE),
            mm(element, BuiltInParameter.WALL_TOP_OFFSET),
            mm(element, BuiltInParameter.WALL_USER_HEIGHT_PARAM),
            text(element, BuiltInParameter.WALL_TOP_IS_ATTACHED),
            phase_name(src_doc, element.CreatedPhaseId),
            phase_name(src_doc, element.DemolishedPhaseId),
        ]
        + [
            None,  # boundary length, filled by the caller
            (
                "; ".join(doors.get(U.idv(element.Id), []))
                if src_doc is doc
                else "(modello collegato)"
            ),
        ]
        + [None] * len(BEHIND)  # filled by the caller
    )


def xy_mm(p):
    return U.to_mm(p.X), U.to_mm(p.Y)


def structural_materials(src_doc, wall_type):
    """Materials of the Structure layers (all layers when none is marked Structure)."""
    cs = wall_type.GetCompoundStructure() if isinstance(wall_type, WallType) else None
    if cs is None:
        return ""
    all_layers = list(cs.GetLayers())
    structure = [
        l for l in all_layers if l.Function == MaterialFunctionAssignment.Structure
    ] or all_layers
    names = []
    for layer in structure:
        material = src_doc.GetElement(layer.MaterialId)
        name = U.ename(material) if material is not None else "(nessun materiale)"
        if name not in names:
            names.append(name)
    return ", ".join(names)


def link_walls():
    """Straight walls of every loaded link, in host coordinates (mm)."""
    result = []
    for link in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = link.GetLinkDocument()
        if link_doc is None:
            continue
        tf = link.GetTotalTransform()
        counts[U.ename(link)] = 0
        # category, not class: walls of an IFC link are DirectShapes (the ST model is linked as IFC)
        for w in (
            FilteredElementCollector(link_doc)
            .OfCategory(BuiltInCategory.OST_Walls)
            .WhereElementIsNotElementType()
        ):
            if isinstance(w, Wall):
                curve = getattr(w.Location, "Curve", None)
                bb = w.get_BoundingBox(None)
                if not isinstance(curve, Line) or bb is None:
                    continue
                z = sorted([tf.OfPoint(bb.Min).Z, tf.OfPoint(bb.Max).Z])
                item = {
                    "label": "%s: %s" % (U.ename(link), U.ename(w.WallType)),
                    "width": U.to_mm(w.Width),
                    "a": xy_mm(tf.OfPoint(curve.GetEndPoint(0))),
                    "b": xy_mm(tf.OfPoint(curve.GetEndPoint(1))),
                    "z": (U.to_mm(z[0]), U.to_mm(z[1])),
                    "material": structural_materials(link_doc, w.WallType),
                }
            else:
                item = shape_axis(link_doc, w, tf)
                if item is None:
                    continue
                item["label"] = "%s: %s" % (U.ename(link), U.ename(w))
            result.append(item)
            counts[U.ename(link)] += 1
    return result


def solids_of(element):
    out = []

    def collect(geo):
        for g in geo:
            if isinstance(g, Solid) and g.Volume > 0:
                out.append(g)
            elif isinstance(g, GeometryInstance):
                collect(g.GetInstanceGeometry())

    geo = element.get_Geometry(Options())
    if geo is not None:
        collect(geo)
    return out


def shape_axis(link_doc, element, tf):
    """Axis, thickness and height (host mm) of a wall without a location line (IFC DirectShape): the largest
    vertical planar face gives the normal, the solid vertices projected on it give the thickness and on the
    face direction the length. None for shapes that are not a straight wall (no vertical face, > MAX_WALL).
    """
    solids = solids_of(element)
    best = None
    for s in solids:
        for f in s.Faces:
            if isinstance(f, PlanarFace) and abs(f.FaceNormal.Z) < 0.01:
                if best is None or f.Area > best.Area:
                    best = f
    if best is None:
        return None
    n = tf.OfVector(best.FaceNormal).Normalize()
    d = XYZ(-n.Y, n.X, 0)
    pts = []
    for s in solids:
        for e in s.Edges:
            c = e.AsCurve()
            pts += [tf.OfPoint(c.GetEndPoint(0)), tf.OfPoint(c.GetEndPoint(1))]
    ts = [p.X * d.X + p.Y * d.Y for p in pts]
    ss = [p.X * n.X + p.Y * n.Y for p in pts]
    width = max(ss) - min(ss)
    if U.to_mm(width) > MAX_WALL:
        return None
    mid = (max(ss) + min(ss)) / 2.0

    def at(t):
        return xy_mm(XYZ(d.X * t + n.X * mid, d.Y * t + n.Y * mid, 0))

    mp = element.LookupParameter("IfcMaterial")
    material = mp.AsString() if mp is not None and mp.HasValue else ""
    if not material:
        m = link_doc.GetElement(best.MaterialElementId)
        material = U.ename(m) if m is not None else ""
    return {
        "width": U.to_mm(width),
        "a": at(min(ts)),
        "b": at(max(ts)),
        "z": (U.to_mm(min(p.Z for p in pts)), U.to_mm(max(p.Z for p in pts))),
        "material": material,
    }


def search_side(room, wall, curve):
    """-1 / +1 along the normal (-dy, dx) of the wall axis: the side away from the room; 0 when unknown."""
    d = curve.Direction
    n = XYZ(-d.Y, d.X, 0)
    mid = curve.Evaluate(0.5, True)
    bb = room.get_BoundingBox(None)
    z = (bb.Min.Z if bb is not None else mid.Z) + 3.3  # ~1 m above the room base, feet
    off = wall.Width / 2.0 + 0.5  # ~150 mm beyond the face, feet
    if room.IsPointInRoom(XYZ(mid.X + n.X * off, mid.Y + n.Y * off, z)):
        return -1
    if room.IsPointInRoom(XYZ(mid.X - n.X * off, mid.Y - n.Y * off, z)):
        return 1
    return 0


def unique(values):
    out = []
    for v in values:
        if v not in out:
            out.append(v)
    return " | ".join(out)


def behind(room, wall, candidates):
    """Values of the BEHIND columns for a host wall, and the coverage % (None when not checked)."""
    curve = getattr(wall.Location, "Curve", None)
    bb = wall.get_BoundingBox(None)
    if not isinstance(curve, Line) or bb is None:
        return ["(muro curvo: non verificato)", "", "", "", ""], None
    a0, a1 = xy_mm(curve.GetEndPoint(0)), xy_mm(curve.GetEndPoint(1))
    za = (U.to_mm(bb.Min.Z), U.to_mm(bb.Max.Z))
    side = search_side(room, wall, curve)
    matches = []
    for c in candidates:
        if min(za[1], c["z"][1]) - max(za[0], c["z"][0]) < MIN_Z_OVERLAP:
            continue
        m = R.wall_behind(
            a0,
            a1,
            U.to_mm(wall.Width) / 2.0,
            side,
            c["a"],
            c["b"],
            c["width"] / 2.0,
            MAX_GAP,
            MIN_OVERLAP,
        )
        if m:
            matches.append((m, c))
    if not matches:
        return ["(nessuno)", "", "", "", "0"], 0.0
    coverage = min(
        100.0, 100.0 * sum(m[1] for m, _c in matches) / U.to_mm(curve.Length)
    )
    return [
        unique(c["label"] for _m, c in matches),
        unique("%.0f" % c["width"] for _m, c in matches),
        unique(c["material"] for _m, c in matches),
        unique("%.0f" % m[0] for m, _c in matches),
        "%.0f" % coverage,
    ], coverage


def csv_cell(v):
    s = "" if v is None else "%s" % v
    if any(c in s for c in ';"\n\r'):
        s = '"%s"' % s.replace('"', '""')
    return s


# ---------------------------------------------------------------- collect

rooms, skipped = [], []
for room in U.instances(doc, [BuiltInCategory.OST_Rooms]):
    code, by_name = room_kinds(room)
    if not (code or by_name):
        continue
    if room.Area <= 0:  # not placed or not enclosed: no boundary
        skipped.append(room)
    else:
        rooms.append((room, code, by_name))
if not rooms:
    forms.alert(
        "Nessun locale vano scala / filtro fumo posizionato e chiuso trovato.",
        exitscript=True,
    )

doors = hosted_doors()
counts = {}  # link name -> walls read
candidates = link_walls()
uncovered = []  # (room, type, id, boundary length, coverage)
options = SpatialElementBoundaryOptions()
rows, non_walls, missing = [], [], 0
by_type = (
    {}
)  # wall type name -> [type mark, thickness, fire values, segments, rooms, length mm]
for room, code, by_name in sorted(
    rooms, key=lambda r: (U.level_name(doc, r[0]) or "", r[0].Number)
):
    room_cols = [
        U.level_name(doc, room) or "",
        room.Number,
        text(room, BuiltInParameter.ROOM_NAME),
        text(room, BuiltInParameter.ROOM_PHASE),
        code,
        by_name,
    ]
    lengths, order = {}, []
    for loop in room.GetBoundarySegments(options) or []:
        for segment in loop:
            resolved = resolve(segment)
            if resolved is None or resolved[2] is None:
                missing += 1
                continue
            key = (resolved[0], U.idv(resolved[2].Id))
            if key not in lengths:
                order.append((key, resolved))
                lengths[key] = 0.0
            lengths[key] += segment.GetCurve().Length
    for key, (model, src_doc, element) in order:
        row = room_cols + element_row(model, src_doc, element, doors)
        length = U.to_mm(lengths[key])
        row[COLUMNS.index("Lunghezza sul confine mm")] = "%.0f" % length
        rows.append(row)
        if isinstance(element, Wall) and src_doc is doc:
            values, coverage = behind(room, element, candidates)
            start = COLUMNS.index(BEHIND[0])
            row[start : start + len(BEHIND)] = values
            if coverage is None or coverage < COVERED:
                uncovered.append(
                    (
                        room.Number,
                        row[COLUMNS.index("Tipo")],
                        row[COLUMNS.index("Id elemento")],
                        length,
                        coverage,
                    )
                )
        if not isinstance(element, Wall):
            non_walls.append((room.Number, row[COLUMNS.index("Categoria")], length))
            continue
        name = row[COLUMNS.index("Tipo")]
        entry = by_type.setdefault(
            name,
            [
                row[COLUMNS.index("Contrassegno tipo")],
                row[COLUMNS.index("Spessore mm")],
                row[COLUMNS.index("Descrizione tipo")],
                0,
                set(),
                0.0,
            ],
        )
        entry[3] += 1
        entry[4].add(room.Number)
        entry[5] += length

# ---------------------------------------------------------------- export

path = forms.save_file(
    file_ext="csv",
    default_name="BMC_MuriREI_scale_%s" % doc.Title,
    title="Salva l'estrazione dei muri REI (CSV per Excel)",
)
if not path:
    script.exit()
with codecs.open(path, "w", encoding="utf-8") as f:
    f.write(codecs.BOM_UTF8.decode("utf-8"))  # BOM: Excel reads the file as UTF-8
    for line in [COLUMNS] + rows:
        f.write(";".join(csv_cell(v) for v in line) + "\r\n")

# ---------------------------------------------------------------- report

output.print_md("# BMC - Muri REI vani scala e filtri fumo")
output.print_md(
    "**File:** `%s`  \n%d locali, %d righe (locale x elemento di confine). Sola lettura: modello non modificato."
    % (path, len(rooms), len(rows))
)
output.print_md("## Tipi di muro sul perimetro")
output.print_table(
    table_data=[
        [name, e[0], e[1], e[2], e[3], len(e[4]), "%.1f" % (e[5] / 1000.0)]
        for name, e in sorted(by_type.items(), key=lambda kv: -kv[1][5])
    ],
    columns=[
        "Tipo",
        "Contrassegno",
        "Spessore mm",
        "Descrizione tipo",
        "Segmenti",
        "Locali",
        "Lunghezza m",
    ],
)
output.print_md("## Tratti senza muro strutturale dietro (copertura < %d%%)" % COVERED)
if not candidates:
    output.print_md(
        "**Nessun muro nei modelli collegati caricati:** carica il modello strutturale e rilancia."
    )
if counts:
    output.print_md(
        "Muri letti dai modelli collegati: %s"
        % ", ".join("%s %d" % (k, v) for k, v in sorted(counts.items()))
    )
if not candidates:
    pass
elif uncovered:
    output.print_md(
        "Qui la compartimentazione dipende dal muro del modello AR (tramezzo) o manca."
    )
    output.print_table(
        table_data=[
            [num, t, i, "%.0f" % length, "-" if cov is None else "%.0f" % cov]
            for num, t, i, length, cov in uncovered
        ],
        columns=["Locale", "Tipo", "Id", "Lunghezza sul confine mm", "Copertura %"],
    )
else:
    output.print_md("Tutti i muri hanno il muro strutturale dietro.")
mismatch = [
    (r.Number, text(r, BuiltInParameter.ROOM_NAME), c, n)
    for r, c, n in rooms
    if c != n  # different, or one of the two missing
]
if mismatch:
    output.print_md("## Locali con nome e codice funzione incoerenti o incompleti")
    output.print_table(
        table_data=[[num, name, c or "-", n or "-"] for num, name, c, n in mismatch],
        columns=["Numero", "Nome", "Funzione da codice", "Funzione da nome"],
    )
if non_walls:
    output.print_md(
        "## Confini che non sono muri (linee di separazione = perimetro non chiuso da muri)"
    )
    output.print_table(
        table_data=[[num, cat, "%.0f" % length] for num, cat, length in non_walls],
        columns=["Locale", "Categoria", "Lunghezza mm"],
    )
if skipped:
    output.print_md(
        "**Locali esclusi (non posizionati o non chiusi):** %s"
        % ", ".join(r.Number for r in skipped)
    )
if missing:
    output.print_md(
        "**%d segmenti di confine senza elemento leggibile** (modello collegato scaricato?)."
        % missing
    )
