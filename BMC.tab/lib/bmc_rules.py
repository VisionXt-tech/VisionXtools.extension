"""BMC project rules (PGI R06 22/09/2026, PIR R06, WBS R06): the single source for every BMC.tab script.

Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
pyRevit adds BMC.tab/lib to the search path of every button of the tab.
Decisions that override the documents are marked "decision"; open doubts are listed in
docs/_corrente/R06_analisi.md section 8.
"""

import json
import math
import os
import re

from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter

LOCAL_DATA_DIR = os.path.join(
    os.path.dirname(__file__), "data"
)  # dev copy, not in the repo
DATA_MARKER = "wbs_R06.json"
DATA_TITLE = "BMC - cartella dati"

# project data, filled by load_data() only in the commands that need it
DATA_DIR = None
SHARED_PARAMS = None  # shared parameter files R06 (name -> file, guid, type, group)
SHARED_DIR = None
WBS_DOMAIN = None
WBS_DESCRIPTION = None
WBS_L8_PARENT = None


def _is_data_dir(path):
    return bool(path) and os.path.isfile(os.path.join(path, DATA_MARKER))


def _find_data_dir(required):
    """Project data (client shared parameter TXTs, WBS) is not in the public repo: each user points
    once to the shared project folder and the choice is kept in the pyRevit user config.
    """
    if _is_data_dir(LOCAL_DATA_DIR):
        return LOCAL_DATA_DIR
    try:
        from pyrevit import forms, script
    except ImportError:  # CPython self-checks without pyRevit
        return None
    cfg = script.get_config("BMC")
    path = cfg.get_option("data_dir", "")
    if _is_data_dir(path):
        return path
    msg = (
        "Questo comando usa la cartella dati BMC condivisa di commessa "
        "(contiene %s, shared_params_R06.json e shared_params\\). "
        "La scelta viene ricordata." % DATA_MARKER
    )
    if required:
        go = forms.alert(msg, title=DATA_TITLE, ok=True, cancel=True)
    else:
        pick = "Seleziona cartella"
        go = (
            forms.alert(
                msg + " Senza, i controlli che la usano vengono saltati.",
                title=DATA_TITLE,
                options=[pick, "Continua senza dati"],
            )
            == pick
        )
    path = forms.pick_folder(title="Cartella dati BMC") if go else None
    if not _is_data_dir(path):
        if required and path:
            forms.alert(
                "Cartella non valida: non contiene %s." % DATA_MARKER,
                title=DATA_TITLE,
                exitscript=True,
            )
        if required:
            script.exit()
        return None
    cfg.set_option("data_dir", path)
    script.save_config()
    return path


def load_data(required=True):
    """Load the project data on demand. required: without a valid folder the command ends quietly;
    otherwise returns False and the caller skips what needs the data."""
    global DATA_DIR, SHARED_PARAMS, SHARED_DIR, WBS_DOMAIN, WBS_DESCRIPTION, WBS_L8_PARENT
    if DATA_DIR:
        return True
    path = _find_data_dir(required)
    if not path:
        if required:
            raise RuntimeError("cartella dati BMC non trovata (%s)" % LOCAL_DATA_DIR)
        return False

    def load(name):
        with open(os.path.join(path, name)) as f:
            return json.load(f)

    SHARED_PARAMS = load("shared_params_R06.json")
    SHARED_DIR = os.path.join(path, "shared_params")
    wbs = load(DATA_MARKER)
    WBS_DOMAIN = [set(wbs["levels"]["L%d" % i].keys()) for i in range(1, 9)]
    WBS_DESCRIPTION = [wbs["levels"]["L%d" % i] for i in range(1, 9)]
    WBS_L8_PARENT = wbs["l8_parent_l7"]
    DATA_DIR = path
    return True


# ---------------------------------------------------------------- general

ND = "nd"  # decision 21/09/2026: canonical "information not available" (not "ND")
ND_VARIANTS = ["ND", "n.d.", "N.D.", "Nd"]
YESNO_TAG = "_SN_"  # PIR Si/No parameters; referents 25/09/2026: never set -> "No" (0), set values kept
MAX_FILE_MB = 300  # PGI 3.11.1

PBP_EXPECTED = {"north_south": 0.0, "east_west": 0.0, "elevation": 260.46, "angle": 0.0}
COORD_TOL = 0.001

# ---------------------------------------------------------------- phases (PGI 3.12.4-3.12.6)

PHASES_EXPECTED = ["Stato di Fatto", "Strip-out", "Post Strip-out", "Stato di Progetto"]
PHASES_NO_CREATION = ["Strip-out", "Post Strip-out"]
SHOW_ALL = [
    "Show All",
    "Mostra tutto",
    "Show All/Mostra tutto",
    "Mostra Tutto/Show All",
]

S, O, D = "ShowByCategory", "ShowOverriden", "DontShow"
# New, Existing, Demolished, Temporary
PHASE_FILTERS_EXPECTED = [
    (SHOW_ALL, (S, O, O, O)),
    (["Strip-out"], (D, S, D, D)),
    (["Comparativa"], (O, S, O, D)),
    (["Solo Demolizione"], (D, D, S, D)),
    (["Solo Stato di Progetto"], (S, D, D, D)),
]
# (phase, phase filter) combinations allowed in views (PGI 3.12.6); None = no filter
VIEW_PHASING = set(
    [
        ("Stato di Fatto", None),
        ("Strip-out", "Strip-out"),
        ("Strip-out", "Solo Demolizione"),
        ("Post Strip-out", "Solo Demolizione"),
        ("Stato di Progetto", "Solo Demolizione"),
        ("Stato di Progetto", "Solo Stato di Progetto"),
        ("Stato di Progetto", "Comparativa"),
        ("Post Strip-out", "Comparativa"),
        ("Strip-out", "Comparativa"),
    ]
    + [("Stato di Progetto", name) for name in SHOW_ALL]
)

# ---------------------------------------------------------------- worksets (PGI 3.11.8)

WS_COORD = [
    "_ELEMENTI NASCOSTI",
    "_FILE DWG COLLEGATI",
    "_FILE IFC COLLEGATI",
    "_FILE RVT COLLEGATI",
    "_GRIGLIE E LIVELLI CONDIVISI",
]
WS_GRIDS, WS_RVT, WS_IFC, WS_DWG = (
    "_GRIGLIE E LIVELLI CONDIVISI",
    "_FILE RVT COLLEGATI",
    "_FILE IFC COLLEGATI",
    "_FILE DWG COLLEGATI",
)
WS_MODEL = (
    "AR_ARCHITETTURA"  # decision 23/09/2026: one modelling workset, PGI prefix AR_
)
WS_AREAS = "AR_AREE"  # PGI 3.11.8: areas on their own workset

# ---------------------------------------------------------------- naming (PGI 3.11.3-3.11.7)

EDIFICI = "XXX|E0[1-5]|E08|E09|E10|E0A|E0B|E0C|E0D|ELN|EIN"
FASI = "00|PG|CD|PP|VF|PA|PD|PE|PC|AB|AP|SO|PS"
RE_FILE = re.compile(
    r"^CMC-(BMA|BCS|BRE|OOU|XXX)-(XXX|XX|CM|CC|AU|RL|IN|LN|CR)-("
    + EDIFICI
    + r")-("
    + FASI
    + r")-(AR|XX)-MOD-(M2|M3|MC|RL)-\d{2}(_[A-Za-z0-9.]+)?$"  # optional local-file suffix
)
RE_LEVEL = re.compile(
    r"^(" + EDIFICI + r")_AR_(B0\dA|L0\d[AS]|R0\dA|XXXX)_(pf|pr|int)_([+-]\d+\.\d{2})$"
)
LEVEL_SEA = "COO - LIVELLO MARE"  # meeting 21/09/2026: sea level keeps its name
RE_LEVEL_COO = re.compile(r"^COO_\S(.*\S)?$")  # other coordination levels COO_...
LEVEL_STRUCTURAL_KINDS = ["pr", "int"]  # PGI 3.11.4: structural flag on

STRAT_CODES = ["CV", "CI", "CE", "CS", "PV", "PO", "PI", "XV", "XO", "XI"]
RE_STRAT_TYPE = re.compile(
    r"^AR_([A-Z]{3})_([A-Z]{2})([123])\.(\d{2})([A-Z])_(.+)_(\d+)$"
)
RE_FAMILY = re.compile(r"^AR_([A-Z]{3})_([A-Z]{2})([123])\.(\d{2})([A-Z])_(.+)$")
RE_TYPE = re.compile(
    r"^AR_([A-Z]{3})_([A-Z]{2})([123])\.(\d{2})([A-Z])_(.+)_(\d+(?:x\d+)*)$"
)
RE_CODE_PREFIX = re.compile(r"^(AR_[A-Z]{3}_([A-Z]{2}[123]\.\d{2}[A-Z]))(_|$)")
RE_TYPE_MARK = re.compile(r"^([A-Z]{2})([123])\.(\d{2})([A-Z])$")
RE_MATERIAL = re.compile(r"^([A-Z]{3}\.\d{2})_")

# Revit category -> PGI 3.11.6 category code (AR discipline)
CAT_CODE = {
    BuiltInCategory.OST_Walls: "MUR",
    BuiltInCategory.OST_Floors: "PAV",
    BuiltInCategory.OST_Ceilings: "CON",
    BuiltInCategory.OST_Doors: "POR",
    BuiltInCategory.OST_Windows: "FIN",
    BuiltInCategory.OST_Railings: "RIN",
    BuiltInCategory.OST_StairsRailing: "RIN",
    BuiltInCategory.OST_Furniture: "ARR",
    BuiltInCategory.OST_Casework: "ARF",
    BuiltInCategory.OST_PlumbingFixtures: "API",
    BuiltInCategory.OST_Stairs: "SCA",
    BuiltInCategory.OST_CurtainWallPanels: "PFC",
    BuiltInCategory.OST_Parking: "POA",
    BuiltInCategory.OST_SpecialityEquipment: "ATS",
    BuiltInCategory.OST_GenericModel: "FOR",
    BuiltInCategory.OST_Ramps: "RAI",
    BuiltInCategory.OST_Columns: "PIL",
    BuiltInCategory.OST_StructuralFraming: "TES",
    BuiltInCategory.OST_Roads: "STD",
    BuiltInCategory.OST_Rooms: "LOC",
}
STRAT_CATS = [
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Roofs,
]
STRAT_CAT_CODE = {
    BuiltInCategory.OST_Walls: "MUR",
    BuiltInCategory.OST_Floors: "PAV",
    BuiltInCategory.OST_Ceilings: "CON",
}


def strat_cat_code(bic, aa):
    """Category code in a stratigraphic type name. Referents 25/09/2026: roofs remodelled as floors
    use category COP and stratigraphy code CS (AR_COP_CS1.01A_...)."""
    if bic == BuiltInCategory.OST_Floors and aa == "CS":
        return "COP"
    return STRAT_CAT_CODE.get(bic)


LOADABLE_CATS = [
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_Furniture,
    BuiltInCategory.OST_Casework,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_SpecialityEquipment,
    BuiltInCategory.OST_Parking,
]
FORBIDDEN_CATS = [
    BuiltInCategory.OST_Roofs
]  # meeting 21/09/2026: roofs remodelled as floors

# ---------------------------------------------------------------- views (PGI 3.11.9-3.11.10)

VIEW_TEMI = ["WIP", "COO", "TAV"]
P_TEMA = "VI_TE_Tema della vista"
P_ARGOMENTO = "VI_TE_Argomento della vista"
# names of views placed on sheets: PA100-L01a.XX, SE50-Tb.XX, PR50-Na.XX, AP100-LEV01a.GEA
RE_VIEW_SHEET = re.compile(r"^(PA|SE|PR|AP)\d+-[A-Za-z0-9]+\.\S.*$")
RE_VIEW_3D = re.compile(r"^3D(WIP|COO|TAV)-\S.*$")
RE_VIEW_SCHEDULE = re.compile(r"^AB(WIP|COO|TAV)-(MA|QU|ET|EV)_\S.*$")
RE_VIEW_TEMPLATE = re.compile(r"^(3D|SZ|PA)_(WIP|COO|TAV)_\d+_\S.*$")

# ---------------------------------------------------------------- parameters (PIR R06, P. Esecutivo)
# (name, "T" = type parameter per PIR "Definisce il tipo = Si", "I" = instance)

P_KEYNOTE, P_TYPE_MARK, P_DESCRIPTION = "Keynote", "Type Mark", "Description"
BUILTIN = {
    "Keynote": BuiltInParameter.KEYNOTE_PARAM,
    "Type Mark": BuiltInParameter.ALL_MODEL_TYPE_MARK,
    "Description": BuiltInParameter.ALL_MODEL_DESCRIPTION,
    "Name": BuiltInParameter.ROOM_NAME,
    "Number": BuiltInParameter.ROOM_NUMBER,
}

WBS_LEVELS = ["WBS_TE_WBS Livello %d" % i for i in range(1, 9)]
WBS_CONCAT = "WBS_TE_WBS"
WBS_SEP = "."
WBS_L8_SEP = ","  # meeting 21/09/2026: several technical elements in level 8
P_EPU = "EPU_TE_Codice EPU"
P_UNIQUE = "PRO_TE_Contrassegno univoco"
P_HOST = "SPE_TE_Codice elemento tecnico HOST"
P_HOST_THK = "SPE_TE_Codice elemento tecnico e spessore HOST"
P_KG = "QTO_NU_Kg al metro"

PRO_TYPE = [(P_KEYNOTE, "T"), (P_TYPE_MARK, "T"), (P_DESCRIPTION, "T")]
WBS_1_5 = [(n, "I") for n in WBS_LEVELS[:5]]
WBS_ALL = [(n, "I") for n in WBS_LEVELS] + [(WBS_CONCAT, "I")]
EPU = [(P_EPU, "T")]
PRE = [
    ("PRE_TE_Isolamento acustico", "T"),
    ("PRE_TE_Trasmittanza termica", "T"),
    ("PRE_TE_Resistenza al fuoco", "T"),
]
VVF = [("VVF_SN_Compartimento", "I"), ("VVF_TE_Resistenza al fuoco", "I")]
VVF_EXIT = [("VVF_SN_Uscita di sicurezza", "I")]
SPE_FRAME = [
    ("SPE_TE_Telaio e Mostre", "T"),
    ("SPE_TE_Struttura Anta", "T"),
    ("SPE_TE_Finitura Anta", "T"),
    ("SPE_TE_Apertura Ante", "T"),
]
SPE_HOST = [(P_HOST, "I")]
SPE_HOST_THK = [(P_HOST_THK, "I")]
SPE_DOOR = [
    ("SPE_SN_Chiudi Porta", "I"),
    ("SPE_TE_Maniglia Interna_Tipologia", "I"),
    ("SPE_TE_Maniglia Esterna_Tipologia", "I"),
    ("SPE_SN_Elettromagnete", "I"),
    ("SPE_TE_Serratura_Tipologia", "I"),
    ("SPE_SN_Fermaporta", "I"),
    ("SPE_SN_Badge con elettroserratura", "I"),
    ("SPE_SN_Contatto magnetico", "I"),
    ("SPE_SN_Cerniera 180 gradi", "I"),
]
SPE_WINDOW = [("SPE_TE_Sistema oscurante_Tipologia", "I")]
UNIQUE = [(P_UNIQUE, "I")]
# referents 25/09/2026: Contrassegno univoco = prefix + progressive number (P01, P02...; P100 after P99);
# windows F01... (decision 25/09/2026, the answer names only P)
UNIQUE_PREFIX = {BuiltInCategory.OST_Doors: "P", BuiltInCategory.OST_Windows: "F"}
RE_UNIQUE = re.compile(r"^([A-Z])(\d{2,})$")
RE_HOST_THK = re.compile(r"^[A-Z]{2}[123]\.\d{2}[A-Z]_\d{3,}$")


def clockwise_order(items):
    """items: [(x, y, payload)] on one level. Payloads ordered clockwise starting from north around their
    centroid (decision 25/09/2026 for the Contrassegno univoco: levels bottom-up, then clockwise from north).
    x east, y north of the project: project north = true north because the PBP angle is 0 (checker 01.1).
    """
    if not items:
        return []
    cx = sum(i[0] for i in items) / float(len(items))
    cy = sum(i[1] for i in items) / float(len(items))
    # ponytail: bearing around the centroid; with a courtyard building internal doors near the centre
    # may interleave, upgrade to a per-wing sweep if the referents want it
    keyed = [
        ((math.atan2(x - cx, y - cy) + 2 * math.pi) % (2 * math.pi), n, p)
        for n, (x, y, p) in enumerate(items)
    ]
    return [p for _b, _n, p in sorted(keyed, key=lambda k: k[:2])]


VVF_ROOM_KINDS = {"VS": "vano scala", "FF": "filtro fumo"}


def vvf_room_kinds(number, name):
    """(function from the number code PIANO.FUNZIONE.NN, function from the name) as VS / FF / ''.
    Both are returned because they can disagree (e.g. "Filtro fumo L02A.VS.17")."""
    parts = (number or "").split(".")
    code = parts[1] if len(parts) == 3 and parts[1] in VVF_ROOM_KINDS else ""
    low = (name or "").strip().lower()
    by_name = [
        k for k, label in sorted(VVF_ROOM_KINDS.items()) if low.startswith(label)
    ]
    return code, by_name[0] if by_name else ""


def wall_behind(
    a0, a1, half_a, side, b0, b1, half_b, max_gap, min_overlap, angle_tol=0.02
):
    """Plan check (mm): is wall b parallel to wall a and on the given side of it within max_gap?
    a0, a1, b0, b1: (x, y) ends of the location lines; half_a, half_b: half widths; side: +1 / -1 along the
    normal n = (-dy, dx) of a0 -> a1, 0 = either side. Returns (gap between the faces, overlap along a) or None.
    Gap 0 = lining against the structure; a few mm of overlap are tolerated (modelling).
    """
    ax, ay = a1[0] - a0[0], a1[1] - a0[1]
    la = (ax * ax + ay * ay) ** 0.5
    bx, by = b1[0] - b0[0], b1[1] - b0[1]
    lb = (bx * bx + by * by) ** 0.5
    if la == 0 or lb == 0:
        return None
    dx, dy = ax / la, ay / la
    if abs(dx * by / lb - dy * bx / lb) > angle_tol:  # not parallel
        return None
    mx, my = (b0[0] + b1[0]) / 2.0 - a0[0], (b0[1] + b1[1]) / 2.0 - a0[1]
    s = -dy * mx + dx * my  # signed distance of b's axis along n
    if side and s * side <= 0:
        return None
    gap = abs(s) - half_a - half_b
    if gap < -5 or gap > max_gap:
        return None
    t = sorted(
        [
            (b0[0] - a0[0]) * dx + (b0[1] - a0[1]) * dy,
            (b1[0] - a0[0]) * dx + (b1[1] - a0[1]) * dy,
        ]
    )
    overlap = min(la, t[1]) - max(0.0, t[0])
    if overlap < min_overlap:
        return None
    return max(gap, 0.0), overlap


COMPOSITION = "Composizione:"


def layer_thickness_cm(width_mm):
    """1,5 / 1,25 / 45: centimetres with a decimal comma, no trailing zeros."""
    return ("%.2f" % (width_mm / 10.0)).rstrip("0").rstrip(".").replace(".", ",")


def stratigraphy_description(current, layers):
    """Type Description rebuilt from the real layers (decision 29/09/2026): the text before "Composizione:" is
    kept, then every layer as full material name (the old descriptions cut it with "...") and thickness in cm,
    joined by " + "; zero-thickness layers (membranes, paint) without thickness.
    layers: [(material name, width mm)]."""
    intro = (current or "").strip()
    if intro == ND:
        intro = ""
    if COMPOSITION in intro:
        intro = intro.split(COMPOSITION)[0]
    intro = " ".join(
        intro.split()
    )  # collapses the double spaces left before "Composizione:"
    parts = [
        "%s (%s cm)" % (name, layer_thickness_cm(w)) if round(w, 1) > 0 else name
        for name, w in layers
    ]
    return ("%s %s %s" % (intro, COMPOSITION, " + ".join(parts))).strip()


def unique_mark(prefix, number):
    return "%s%02d" % (prefix, number)


def host_thickness_code(host_type_mark, width_mm):
    """Referents 25/09/2026: SPE_TE_Codice elemento tecnico e spessore HOST = Type Mark_thickness (PV2.01K_112),
    thickness in mm as in the type name (layer widths are feet: round to 0.001 mm first, half up).
    """
    return "%s_%03d" % (host_type_mark, int(round(round(width_mm, 3))))


UTI = [
    ("Name", "I"),
    ("Number", "I"),
    ("UTI_TE_Abbreviazione Locale", "I"),
    ("UTI_TE_Edificio", "I"),
    ("UTI_TE_Scala", "I"),
    ("UTI_TE_Piano", "I"),
    ("UTI_TE_Tipologia appartamento", "I"),
    ("UTI_TE_Codice appartamento", "I"),
    ("UTI_TE_Codice reparto", "I"),
]
FNT = [
    ("FNT_TE_Tinteggiatura", "I"),
    ("FNT_TE_Zoccolino", "I"),
    ("FNT_TE_Finitura a soffitto", "I"),
]
RAI = [
    ("RAI_NU_Superficie aerante", "I"),
    ("RAI_NU_Superficie illuminante", "I"),
    ("RAI_NU_Rapporto aerante", "I"),
    ("RAI_NU_Rapporto illuminante", "I"),
    ("RAI_SN_Aerazione meccanica", "I"),
    ("RAI_SN_Illuminazione artificiale", "I"),
    ("RAI_SN_Verifica RAI richiesta", "I"),
]
QTO_KG = [(P_KG, "T")]
COMMON = PRO_TYPE + WBS_ALL + EPU

# PIR R06 "Oggetti (App. idrauliche, Arredi, Arredi fissi, Posti auto, Scale)" -> PGI 3.11.6 API, ARR, ARF, POA,
# SCA = Revit Plumbing Fixtures, Furniture, Casework ("Arredi fissi"), Parking, Stairs. Furniture Systems,
# Specialty Equipment (ATS) and Planting are not in the PIR group (audit 24/09: docs/_sorgenti/audit_categorie.py).
OGG_CATS = [
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_Furniture,
    BuiltInCategory.OST_Casework,
    BuiltInCategory.OST_Parking,
    BuiltInCategory.OST_Stairs,
]

# PIR group: (code, label, categories, wall kind filter, parameters). Parameters only on main
# elements (meeting 21/09/2026): no mullions, balusters, runs, landings.
PIR_GROUPS = [
    # PIR "Aperture" = PGI FOR "Forometrie (modelli generici)"; Revit shaft openings are openings too
    (
        "APE",
        "Aperture",
        [BuiltInCategory.OST_GenericModel, BuiltInCategory.OST_ShaftOpening],
        None,
        PRO_TYPE,
    ),
    ("MUR", "Muri", [BuiltInCategory.OST_Walls], "basic", COMMON + PRE + VVF),
    (
        "FAC",
        "Facciate continue",
        [BuiltInCategory.OST_Walls],
        "curtain",
        COMMON + PRE + VVF,
    ),
    ("PAV", "Pavimenti", [BuiltInCategory.OST_Floors], None, COMMON + PRE + VVF),
    ("CON", "Controsoffitti", [BuiltInCategory.OST_Ceilings], None, COMMON + PRE + VVF),
    (
        "PFC",
        "Pannelli CW",
        [BuiltInCategory.OST_CurtainWallPanels],
        None,
        COMMON + PRE + VVF + SPE_HOST,
    ),
    (
        "POR",
        "Porte",
        [BuiltInCategory.OST_Doors],
        None,
        COMMON
        + UNIQUE
        + PRE
        + VVF
        + VVF_EXIT
        + SPE_FRAME
        + SPE_HOST
        + SPE_HOST_THK
        + SPE_DOOR,
    ),
    (
        "FIN",
        "Finestre",
        [BuiltInCategory.OST_Windows],
        None,
        COMMON
        + UNIQUE
        + PRE
        + VVF
        + VVF_EXIT
        + SPE_FRAME
        + SPE_HOST
        + SPE_HOST_THK
        + SPE_WINDOW,
    ),
    (
        "RIN",
        "Ringhiere",
        [BuiltInCategory.OST_Railings, BuiltInCategory.OST_StairsRailing],
        None,
        COMMON + QTO_KG,
    ),
    ("OGG", "Oggetti", OGG_CATS, None, COMMON),
    ("LOC", "Locali", [BuiltInCategory.OST_Rooms], None, WBS_1_5 + UTI + FNT + RAI),
]


def pir_names(params):
    return [name for name, _ in params]


def all_pir_parameters():
    """Every shared (non built-in) parameter required by the PIR, in PIR order, with its level."""
    seen, result = set(), []
    for _code, _label, _cats, _kind, params in PIR_GROUPS:
        for name, level in params:
            if name not in BUILTIN and name not in seen:
                seen.add(name)
                result.append((name, level))
    return result


# decision 23/09/2026: every R06 TXT except MEP and ST is used by the ARC model
SHARED_FILES = {
    "AR": "BMC_Parametri condivisi AR.txt",
    "GE": "BMC_Parametri condivisi GE.txt",
    "SPACE": "BMC_Parametri condivisi SPACE.txt",
    "BROWSER": "BMC_Parametri condivisi BROWSER.txt",
    "TAVOLE": "BMC_Parametri condivisi TAVOLE.txt",
    "Supporto_BMC_Parametri condivisi": "Supporto_BMC_Parametri condivisi.txt",
}
# parameters to remove (meeting 21/09/2026: not in the PIR / not in the R06 TXT files)
LEGACY_PARAMS = [
    "SGAS_WBS_Level 1",
    "SGAS_WBS_Level 2",
    "SGAS_WBS_Level 3",
    "SGAS_WBS_Level 4",
    "SGAS_WBS_Level 5",
    "SGAS_WBS_Level 6",
    "SGAS_WBS_Level 7",
    "SGAS_View Status",
    "SGAS_View Topic",
    "SGAS_View Type",
    "SUP X 1/8",
    "UTI_TE_Codice",
    "VI_TE_Status della vista",
    "RAI_ARE_Superficie illuminante",
]

# where the parameters outside PIR_ARC are bound
AREA_PARAMS = [  # PIR_SUP e AREE (instance, Areas)
    "UTI_TE_Edificio",
    "UTI_TE_Piano",
    "UTI_TE_Scala",
    "UTI_TE_Tipologia superficie",
    "UTI_TE_Funzione",
]
# PIR Viste: views. PGI 3.11.9: Tema also organises the schedules (LIV.1); schedules LIV.2 is the category and
# sheets are organised by name, so Argomento only on views and nothing on sheets.
VIEW_CATS = [BuiltInCategory.OST_Views]
VIEW_CATS_BY_PARAM = {
    P_TEMA: [BuiltInCategory.OST_Views, BuiltInCategory.OST_Schedules]
}
SHEET_CATS = [
    BuiltInCategory.OST_Sheets
]  # TAVOLE: decision 23/09/2026, sheets stay in the ARC model
SAG_CATS = [  # Supporto: decision 23/09/2026 (no PIR row)
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_GenericModel,
]


def binding_targets():
    """name -> (level "I"/"T", [BuiltInCategory]) for every parameter of the SHARED_FILES TXTs (needs load_data).
    PIR_ARC groups first, then PIR_SUP e AREE, Viste, Tavole and Supporto."""
    targets = {}
    for _code, _label, cats, _kind, params in PIR_GROUPS:
        for name, level in params:
            if name in BUILTIN:
                continue
            old_level, old_cats = targets.get(name, (level, []))
            targets[name] = (
                old_level,
                old_cats + [c for c in cats if c not in old_cats],
            )
    for name in AREA_PARAMS:
        level, cats = targets.get(name, ("I", []))
        targets[name] = (level, cats + [BuiltInCategory.OST_Areas])
    for name, ref in SHARED_PARAMS.items():
        if ref["file"] not in SHARED_FILES or name in targets:
            continue
        if ref["file"] == "BROWSER":
            targets[name] = ("I", list(VIEW_CATS_BY_PARAM.get(name, VIEW_CATS)))
        elif ref["file"] == "TAVOLE":
            targets[name] = ("I", list(SHEET_CATS))
        elif ref["file"].startswith("Supporto"):
            targets[name] = ("I", list(SAG_CATS))
    return targets


# legacy values worth keeping: (source, target, only these categories or None, kind of conversion)
# applied only where the target is empty; every value is in the backup JSON anyway
MIGRATIONS = [
    ("SGAS_View Status", P_TEMA, None, "text"),
    ("VI_TE_Status della vista", P_TEMA, None, "text"),
    ("SGAS_View Topic", P_ARGOMENTO, None, "text"),
    (
        "SGAS_WBS_Level 1",
        "UTI_TE_Tipologia superficie",
        [BuiltInCategory.OST_Areas],
        "text",
    ),
    (
        "RAI_ARE_Superficie illuminante",
        "RAI_NU_Superficie illuminante",
        None,
        "area_to_m2",
    ),
]

# ---------------------------------------------------------------- WBS (BGMC_WBS_R06.xlsx)


def wbs_level_errors(values):
    """values: list of 8 strings ('' = empty). Returns a list of messages for values out of domain (needs load_data)."""
    errors = []
    for i, value in enumerate(values):
        if not value or value == ND:
            continue
        codes = [v.strip() for v in value.split(WBS_L8_SEP)] if i == 7 else [value]
        for code in codes:
            if code not in WBS_DOMAIN[i]:
                errors.append("L%d `%s` non ammesso" % (i + 1, code))
            elif i == 7 and values[6] and WBS_L8_PARENT.get(code) != values[6]:
                errors.append(
                    "L8 `%s` appartiene a L7 `%s`, non a `%s`"
                    % (code, WBS_L8_PARENT.get(code), values[6])
                )
    return errors


# Expected (L7, L8) per kind of element, from the WBS R06 descriptions. Proposal of 24/09 validated by the
# referents on 25/09/2026 (external doors changed): the first pair is the one Compila WBS writes when
# nothing better exists; the checker (06.WBS8) warns when an element has a valid pair that is not in its list.
#   key: (PIR group, discriminator) - walls/floors: AA code of the Type Mark, doors: "CW" when hosted by a
#   curtain wall, else function digit (1 external, 2 internal), objects: Revit category, others "".
_FACADE = ("3IN", "402")
_WALL_502, _WALL_506 = ("3AR", "502"), ("3AR", "506")
WBS_EXPECTED = {
    ("MUR", "CV"): [_FACADE],
    ("MUR", "CE"): [_FACADE],
    ("MUR", "CI"): [_FACADE],
    ("MUR", "CS"): [_FACADE],
    ("MUR", "XV"): [("3IN", "471"), _FACADE],  # external partitions: parapets or facade
    ("MUR", "XO"): [("3IN", "471"), _FACADE],
    ("MUR", "XI"): [("3IN", "471"), _FACADE],
    ("MUR", "PV"): [
        _WALL_502
    ],  # partitions / linings with plasterboard, bricks, insulation
    ("MUR", "PO"): [_WALL_502],
    ("MUR", "PI"): [_WALL_502],
    ("MUR", "RIV"): [
        _WALL_506
    ],  # internal walls made only of finishes (e.g. PV2.01L tiles 12 mm)
    ("MUR", ""): [_WALL_502, _WALL_506, _FACADE],
    ("FAC", ""): [("3IN", "404")],  # curtain walls: glazed facades
    ("PAV", "CS"): [("3IN", "454")],  # roofs remodelled as floors: opaque roofs
    ("PAV", ""): [("3AR", "506")],
    ("CON", ""): [("3AR", "506")],
    ("PFC", ""): [("3IN", "404")],
    # referents 25/09/2026: external doors are opaque facade (402), glazed facade (404) only in curtain walls
    ("POR", "1"): [_FACADE],
    ("POR", "CW"): [("3IN", "404")],
    ("POR", "2"): [("3AR", "508")],
    ("POR", ""): [("3AR", "508"), _FACADE, ("3IN", "404")],
    ("FIN", ""): [("3IN", "404"), ("3IN", "406")],
    ("RIN", "1"): [("3IN", "471")],
    ("RIN", "2"): [("3AR", "515")],  # 471 is under Involucro (validated 25/09/2026)
    ("RIN", ""): [("3IN", "471"), ("3AR", "515")],
    ("OGG", "OST_PlumbingFixtures"): [("3AR", "510")],
    ("OGG", "OST_Furniture"): [("9FF", "B10")],
    ("OGG", "OST_Casework"): [("9FF", "B11")],
    ("OGG", "OST_Stairs"): [
        ("3AR", "506")
    ],  # ARC stair = finishes, structure in ST (validated 25/09/2026)
    ("OGG", "OST_Parking"): [("3AR", "515")],
}
FINISH_PREFIXES = (
    "FIN",
    "TIN",
    "INT",
)  # layers that make a wall a cladding (rivestimento)


def wbs_expected(
    group, bic, type_mark_value, layer_codes=None, type_name=None, in_curtain_wall=False
):
    """Expected (L7, L8) pairs for an element ([] = no rule). layer_codes: material codes of the type layers
    (e.g. ["FIN.07"]): an internal wall made only of finishes is a cladding (506), not a partition (502).
    in_curtain_wall: the element (door) is hosted by a curtain wall.
    Rule: docs/20260924/Proposta_WBS_L7_L8.md, validated 25/09/2026."""
    if group == "POR" and in_curtain_wall:
        return WBS_EXPECTED[("POR", "CW")]
    m = RE_TYPE_MARK.match(type_mark_value or "")
    if not m and type_name:
        # OLD_/WIP_ types and empty Type Marks: the code is still in the name (OLD_AR_MUR_CV1.01B_...)
        n = RE_CODE_PREFIX.match(re.sub(r"^(OLD_|WIP_)+", "", type_name))
        m = RE_TYPE_MARK.match(n.group(2)) if n else None
    if group in ("MUR", "PAV") and m:
        key = (group, m.group(1))
        if (
            group == "MUR"
            and m.group(1) in ("PV", "PO", "PI")
            and layer_codes
            and all(c[:3] in FINISH_PREFIXES for c in layer_codes)
        ):
            key = ("MUR", "RIV")
    elif group in ("POR", "RIN") and m:
        key = (group, m.group(2))
    elif group == "OGG":
        key = (group, str(bic))
    else:
        key = (group, "")
    return WBS_EXPECTED.get(key) or WBS_EXPECTED.get((group, ""), [])


def wbs_l4_from_l6(l6):
    """AG above ground, BG basement (L6 i01, i02)."""
    if not l6:
        return None
    return "BG" if l6.startswith("i") else "AG"


def room_abbreviation(number, name):
    """PGI 3.11.12: from the room number PIANO.FUNZIONE.NN (L00A.BG.04 -> BG), else the residential table."""
    parts = (number or "").split(".")
    if len(parts) == 3 and parts[1].isalpha():
        return parts[1]
    return ROOM_ABBREVIATIONS.get((name or "").strip())


def _wbs_valid(i, value):
    return bool(value) and value != ND and value in WBS_DOMAIN[i]


def _wbs_l8_valid(value, l7):
    codes = [c.strip() for c in (value or "").split(WBS_L8_SEP) if c.strip()]
    return bool(codes) and all(
        c in WBS_DOMAIN[7] and WBS_L8_PARENT.get(c) == l7 for c in codes
    )


def wbs_merge(current, proposed, n=8):
    """Final WBS values: valid current values are kept, empty/invalid ones take the proposal.
    L8 must belong to the final L7; without a valid proposal the current value is left as it is
    (never cleared) and the checker keeps reporting it."""
    final = [(current[i] if i < len(current) else "") or "" for i in range(n)]
    for i in range(min(n, 7)):
        if not _wbs_valid(i, final[i]) and proposed[i] and _wbs_valid(i, proposed[i]):
            final[i] = proposed[i]
    if (
        n == 8
        and not _wbs_l8_valid(final[7], final[6])
        and _wbs_l8_valid(proposed[7], final[6])
    ):
        final[7] = proposed[7]
    return final


def wbs_l6_from_level(level_name):
    """WBS level 6 from a level name: E0C_AR_L01A_pf_+0.59 -> L01, B01A -> i01, R01A -> None."""
    m = RE_LEVEL.match(level_name or "")
    if not m:
        return None
    code = m.group(2)
    if code.startswith("B"):
        return "i%s" % code[1:3]
    if code.startswith("L"):
        return code[:3]
    return None


# ---------------------------------------------------------------- rooms (PGI 3.11.12)

ROOM_ABBREVIATIONS = {
    "Antibagno": "Ab",
    "Bagno": "B",
    "Balcone": "Ba",
    "Cabina armadio": "Ca",
    "Camera singola": "Cs",
    "Camera doppia": "Cd",
    "Cucina": "K",
    "Disimpegno": "D",
    "Ingresso": "I",
    "Locale wc": "Wc",
    "Lavanderia": "La",
    "Loggia": "L",
    "Ripostiglio": "R",
    "Soggiorno": "S",
    "Soggiorno/cucina": "S+k",
    "Studio": "St",
    "Terrazzo": "T",
    "Cantina": "Ca",
    "Box": "Bo",
    "Posto Auto": "Pa",
    "Posto Moto": "Pm",
}
